/**
 * Conga AI Assistance — API Server
 *
 * POST /api/chat        — AI chat endpoint (Azure OpenAI with keyword fallback)
 * POST /api/metadata    — parse uploaded .docx and return content controls
 * GET  /api/proxy       — proxy SharePoint page, strip X-Frame-Options so it embeds
 * GET  /api/health      — health check
 */

require("dotenv").config();

const express  = require("express");
const cors     = require("cors");
const multer   = require("multer");
const AdmZip   = require("adm-zip");
const fetch    = require("node-fetch");
const https    = require("https");
const fs       = require("fs");
const path     = require("path");
const zlib     = require("zlib");
const { AzureOpenAI, OpenAI } = require("openai");

// ── Logger ────────────────────────────────────────────────────────────────────

const LOGS_DIR = path.join(__dirname, "logs");
if (!fs.existsSync(LOGS_DIR)) fs.mkdirSync(LOGS_DIR, { recursive: true });

function getLogFilePath() {
  const date = new Date().toISOString().split("T")[0]; // YYYY-MM-DD
  return path.join(LOGS_DIR, `${date}.log`);
}

function writeLog(level, category, data) {
  const entry = { ts: new Date().toISOString(), level, category, ...data };
  const line  = JSON.stringify(entry) + "\n";
  try { fs.appendFileSync(getLogFilePath(), line, "utf8"); } catch {}
  const prefix = `[${entry.ts}] [${level.padEnd(5)}] [${category}]`;
  if (level === "ERROR") console.error(prefix, data.message || data.error || "");
  else console.log(prefix, data.message || "");
}

const logger = {
  info:  (cat, data) => writeLog("INFO",  cat, data),
  warn:  (cat, data) => writeLog("WARN",  cat, data),
  error: (cat, data) => writeLog("ERROR", cat, data),
  debug: (cat, data) => writeLog("DEBUG", cat, data),
};

// ── Azure OpenAI client ───────────────────────────────────────────────────────
// Supports two endpoint styles:
//   Classic:  https://<resource>.openai.azure.com          → AzureOpenAI client (needs api-version)
//   Foundry:  https://<project>.services.ai.azure.com/...  → OpenAI client (no api-version)
const AZURE_ENDPOINT   = process.env.AZURE_OPENAI_ENDPOINT   || "";
const AZURE_API_KEY    = process.env.AZURE_OPENAI_API_KEY    || "";
const AZURE_DEPLOYMENT = process.env.AZURE_OPENAI_DEPLOYMENT || "gpt-4o";
const AZURE_API_VER    = process.env.AZURE_OPENAI_API_VERSION || "2024-12-01-preview";

function buildAiClient() {
  if (!AZURE_ENDPOINT || !AZURE_API_KEY) return null;

  // Endpoint styles:
  //   /openai/v1 path  → new Azure OpenAI compatibility API (no api-version, api-key header)
  //   services.ai.azure.com → AI Foundry inference endpoint
  //   .openai.azure.com only → classic Azure OpenAI (needs api-version)
  const hasV1Path  = /\/openai\/v1\/?$/.test(AZURE_ENDPOINT);
  const isFoundry  = AZURE_ENDPOINT.includes("services.ai.azure.com");

  if (hasV1Path || isFoundry) {
    // Ensure baseURL ends with /v1 (no trailing slash — SDK appends /chat/completions)
    let baseURL = AZURE_ENDPOINT.replace(/\/+$/, "");
    if (!baseURL.endsWith("/v1")) baseURL += "/v1";
    console.log(`AI Foundry / v1-compatible endpoint detected → baseURL: ${baseURL}`);
    return new OpenAI({
      baseURL,
      apiKey: AZURE_API_KEY,
      defaultHeaders: { "api-key": AZURE_API_KEY }   // Azure requires this header
    });
  }

  // Classic Azure OpenAI endpoint (https://<resource>.openai.azure.com)
  return new AzureOpenAI({ endpoint: AZURE_ENDPOINT, apiKey: AZURE_API_KEY, apiVersion: AZURE_API_VER });
}

const aiClient = buildAiClient();

if (aiClient) {
  console.log(`Azure OpenAI connected → deployment: ${AZURE_DEPLOYMENT}`);
} else {
  console.warn("Azure OpenAI not configured — using keyword fallback. Set AZURE_OPENAI_* in api/.env");
}

// ── Azure AI Agent (Assistants API) client ────────────────────────────────────
// Run `node create-agent.js` once to create the agent, then set AZURE_AGENT_ID in .env.
// When configured, uses persistent threads for true multi-turn conversation memory.
// Falls back to stateless chat.completions if AZURE_AGENT_ID is not set.
const AZURE_AGENT_ID = process.env.AZURE_AGENT_ID || "";

function buildAssistantsClient() {
  if (!AZURE_API_KEY || !AZURE_AGENT_ID) return null;
  // Prefer the Foundry project endpoint (services.ai.azure.com) so agents appear
  // in the Foundry portal. Fall back to stripping the OpenAI endpoint if not set.
  const foundryEndpoint = process.env.AZURE_FOUNDRY_ENDPOINT || "";
  const baseEndpoint = foundryEndpoint
    ? foundryEndpoint.replace(/\/+$/, "")
    : AZURE_ENDPOINT.replace(/\/openai\/v1\/?$/, "").replace(/\/v1\/?$/, "").replace(/\/+$/, "");
  return new AzureOpenAI({ endpoint: baseEndpoint, apiKey: AZURE_API_KEY, apiVersion: "2025-01-01-preview" });
}

const assistantsClient = buildAssistantsClient();
if (assistantsClient) {
  console.log(`Azure AI Agent connected → agent: ${AZURE_AGENT_ID}`);
} else {
  console.log("Azure AI Agent not configured — set AZURE_AGENT_ID in .env to enable persistent threads");
}

// ── Tool-calling mode flag ─────────────────────────────────────────────────────
// When MCP tools are discovered, we use Chat Completions with function calling
// so the model can invoke Conga API tools + document tools in a single loop.
const AGENT_ID = AZURE_AGENT_ID;
if (AGENT_ID) {
  console.log(`Agent ID configured: ${AGENT_ID} (using Chat Completions + tool calling)`);
}

// In-memory thread store: threadId → { created, lastUsed }
const activeThreads = new Map();

// ── Foundry Toolbox MCP client ─────────────────────────────────────────────
const FOUNDRY_TOOLBOX_URL = process.env.FOUNDRY_TOOLBOX_URL || "";

async function callFoundryToolbox(method, params = {}) {
  if (!FOUNDRY_TOOLBOX_URL) {
    throw new Error("FOUNDRY_TOOLBOX_URL not configured in .env");
  }

  const body = {
    jsonrpc: "2.0",
    id: `req-${Date.now()}-${Math.floor(Math.random() * 100000)}`,
    method,
    params,
  };

  const res = await fetch(FOUNDRY_TOOLBOX_URL, {
    method: "POST",
    headers: {
      "api-key": AZURE_API_KEY,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  const text = await res.text().catch(() => "");
  if (!res.ok) throw new Error(`Foundry toolbox ${res.status}: ${text}`);
  return text ? JSON.parse(text) : { success: true };
}

if (FOUNDRY_TOOLBOX_URL) {
  console.log(`Foundry Toolbox MCP → ${FOUNDRY_TOOLBOX_URL.split("?")[0]}...`);
}

// ── Conga token provider (auto-auth for review tools) ─────────────────────
const CONGA_CLIENT_ID     = process.env.CONGA_CLIENT_ID || "";
const CONGA_CLIENT_SECRET = process.env.CONGA_CLIENT_SECRET || "";
const CONGA_SCOPE         = process.env.CONGA_SCOPE || "sign";
const CONGA_AUTH_URL      = process.env.CONGA_AUTH_URL || "https://login-rlsdev.congacloud.io/api/v1/auth/connect/token";

let cachedCongaToken   = "";
let congaTokenExpiresAt = 0;
let inflightTokenReq    = null;

async function getCongaToken() {
  if (cachedCongaToken && Date.now() < congaTokenExpiresAt - 30000) {
    return cachedCongaToken;
  }
  if (inflightTokenReq) return inflightTokenReq;

  inflightTokenReq = (async () => {
    if (!CONGA_CLIENT_ID || !CONGA_CLIENT_SECRET) {
      throw new Error("CONGA_CLIENT_ID / CONGA_CLIENT_SECRET not configured");
    }
    const body = new URLSearchParams({
      client_id:     CONGA_CLIENT_ID,
      client_secret: CONGA_CLIENT_SECRET,
      scope:         CONGA_SCOPE,
      grant_type:    "client_credentials",
    });
    const res = await fetch(CONGA_AUTH_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });
    const text = await res.text();
    if (!res.ok) throw new Error(`Conga auth failed (${res.status}): ${text}`);
    const json = JSON.parse(text);
    cachedCongaToken    = json.access_token;
    const expiresIn     = json.expires_in || 3600;
    congaTokenExpiresAt = Date.now() + expiresIn * 1000;
    logger.info("CONGA_AUTH", { message: `Token acquired, expires in ${expiresIn}s` });
    return cachedCongaToken;
  })().finally(() => { inflightTokenReq = null; });

  return inflightTokenReq;
}

/** Inject Conga bearer token for congaendreviewtool___ tool calls */
async function maybeInjectAuth(toolName, args) {
  const needsAuth = toolName.startsWith("congaendreviewtool___");
  if (!needsAuth) return args;
  if (args.Authorization && String(args.Authorization).trim()) return args;
  try {
    const token = await getCongaToken();
    return { ...args, Authorization: `Bearer ${token}` };
  } catch (err) {
    logger.warn("TOKEN", { message: `Auth injection failed: ${err.message}` });
    return args;
  }
}

/**
 * Call getReviewAuthorization via MCP to obtain the ExternalToken needed
 * by the endReviewerReview endpoint (Office365 flow).
 */
async function getReviewExternalToken(reviewtype) {
  const token = await getCongaToken();
  const authResult = await callFoundryToolbox("tools/call", {
    name: "congaendreviewtool___getReviewAuthorization",
    arguments: {
      reviewtype: reviewtype || "Office365",
      Authorization: `Bearer ${token}`,
    },
  });

  // Extract external token from MCP response
  let externalToken = "";
  try {
    const payload = authResult?.result;

    // Log the content parts specifically (not the huge _meta block)
    if (payload?.content) {
      const contentTexts = payload.content.map(c => c.text || "").join("");
      logger.info("REVIEW_AUTH", { message: "Response content", content: contentTexts.substring(0, 2000) });
    } else {
      logger.info("REVIEW_AUTH", { message: "Full response (no content array)", result: JSON.stringify(authResult).substring(0, 2000) });
    }

    if (payload?.content) {
      // MCP content array format: [{ type: "text", text: "..." }]
      const textContent = payload.content.find(c => c.type === "text");
      if (textContent && textContent.text) {
        const raw = textContent.text.trim();
        // Try parsing as JSON
        try {
          const parsed = JSON.parse(raw);
          // Conga Review API returns: { Value: { Token: { AuthToken: "..." } } }
          externalToken = parsed?.Value?.Token?.AuthToken
            || parsed?.value?.token?.authToken
            || parsed.externalToken || parsed.ExternalToken
            || parsed.token || parsed.Token
            || parsed.access_token || parsed.accessToken
            || parsed.authUrl || parsed.authURL || parsed.auth_url
            || "";
          // If parsed is a plain string, use it directly
          if (!externalToken && typeof parsed === "string") externalToken = parsed;
          // If parsed has a nested data/value object
          if (!externalToken && parsed.data) {
            externalToken = parsed.data.externalToken || parsed.data.token || parsed.data.access_token || "";
          }
          // If parsed has a redirectUrl with token in query params
          if (!externalToken && (parsed.redirectUrl || parsed.url)) {
            const url = parsed.redirectUrl || parsed.url;
            const match = url.match(/[?&](?:token|code|access_token)=([^&]+)/i);
            if (match) externalToken = decodeURIComponent(match[1]);
            else externalToken = url; // pass the whole URL as external token
          }
        } catch {
          // Not JSON — use the raw text as the token
          externalToken = raw;
        }
      }
    } else if (typeof payload === "string") {
      externalToken = payload;
    }
  } catch (err) {
    logger.warn("REVIEW_AUTH", { message: `Could not parse external token: ${err.message}` });
  }

  if (externalToken) {
    logger.info("REVIEW_AUTH", { message: `ExternalToken obtained (${externalToken.length} chars)` });
  } else {
    logger.warn("REVIEW_AUTH", { message: "ExternalToken could not be extracted from response" });
  }
  return externalToken;
}

// Warm up Conga token on startup if credentials are set
if (CONGA_CLIENT_ID && CONGA_CLIENT_SECRET) {
  getCongaToken().catch(() => {});
  console.log("Conga auth configured → tokens will be provisioned automatically");
}

// ── Tool definitions for the Hosted Agent ──────────────────────────────────
// These are passed to createAndPoll so the agent knows it can call tools.

const DOCUMENT_TOOLS = [
  {
    type: "function",
    function: {
      name: "navigate_to_clause",
      description: "Navigate to a specific clause or field in the document by tag name.",
      parameters: {
        type: "object",
        properties: { tag: { type: "string", description: "The tag name of the clause or field to navigate to." } },
        required: ["tag"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "update_clause",
      description: "Update/replace text within a clause. Provide the exact text to find and its replacement.",
      parameters: {
        type: "object",
        properties: {
          tag:     { type: "string", description: "The tag name of the clause to update." },
          find:    { type: "string", description: "The exact text to find (must not be empty)." },
          replace: { type: "string", description: "The replacement text." },
        },
        required: ["tag", "find", "replace"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_clause",
      description: "Insert a brand new clause or section into the document.",
      parameters: {
        type: "object",
        properties: {
          tag:  { type: "string", description: "Name/tag for the new clause." },
          text: { type: "string", description: "Full text content of the new clause." },
        },
        required: ["tag", "text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "delete_clause",
      description: "Delete a clause or content control from the document.",
      parameters: {
        type: "object",
        properties: { tag: { type: "string", description: "The tag name of the clause to delete." } },
        required: ["tag"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "list_clauses",
      description: "List all content controls (clauses and fields) found in the document.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "search_clause",
      description: "Search for a clause or field by keyword.",
      parameters: {
        type: "object",
        properties: { keyword: { type: "string", description: "The keyword to search for." } },
        required: ["keyword"],
      },
    },
  },
];

// ── Dynamic MCP tool discovery ────────────────────────────────────────────
// Fetches available tools from the Foundry Toolbox MCP endpoint on startup
// and converts them to OpenAI function-tool format so the agent can call them.

let mcpToolDefs = []; // populated by discoverMcpTools()

function convertMcpToolToOpenAI(mcpTool) {
  const schema = mcpTool.inputSchema || { type: "object", properties: {} };
  // Remove auth-related params since we auto-inject them server-side
  const AUTO_INJECTED = new Set(["Authorization", "ExternalToken"]);
  const required = (schema.required || []).filter(r => !AUTO_INJECTED.has(r));
  // Also strip them from properties so the model never tries to supply them
  const properties = { ...(schema.properties || {}) };
  for (const key of AUTO_INJECTED) delete properties[key];
  return {
    type: "function",
    function: {
      name: mcpTool.name,
      description: mcpTool.description || `MCP tool: ${mcpTool.name}`,
      parameters: {
        type: "object",
        properties,
        ...(required.length ? { required } : {}),
      },
    },
  };
}

async function discoverMcpTools() {
  if (!FOUNDRY_TOOLBOX_URL) return;
  try {
    const response = await callFoundryToolbox("tools/list");
    const tools = response?.result?.tools || [];
    // Only keep review tools (congaendreviewtool___*); skip congaapitools___ and congaauth___
    const KEEP_PREFIXES = ["congaendreviewtool___"];
    const filtered = tools.filter(t => KEEP_PREFIXES.some(p => t.name.startsWith(p)));
    mcpToolDefs = filtered.map(convertMcpToolToOpenAI);
    console.log(`MCP tool discovery → found ${tools.length} tools, kept ${mcpToolDefs.length}: ${mcpToolDefs.map(t => t.function.name).join(", ")}`);
  } catch (err) {
    console.error(`MCP tool discovery failed: ${err.message}`);
  }
}

// Run discovery on startup
discoverMcpTools();

const app     = express();
const upload  = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

// In production set ALLOWED_ORIGIN to your Azure Static Web App URL.
// In development (no env var) all origins are allowed for convenience.
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN;
app.use(cors(
  ALLOWED_ORIGIN
    ? { origin: ALLOWED_ORIGIN, methods: ["GET", "POST", "OPTIONS"], allowedHeaders: ["Content-Type"] }
    : {}
));
app.use(express.json());

// ── Request logger middleware ─────────────────────────────────────────────────
app.use((req, _res, next) => {
  logger.info("REQUEST", { message: `${req.method} ${req.path}`, ip: req.ip });
  next();
});

// ── GET /api/proxy ────────────────────────────────────────────────────────────
// Fetches a SharePoint URL server-side and strips the X-Frame-Options /
// Content-Security-Policy frame-ancestors headers so the page can be embedded.
//
// Usage: /api/proxy?url=https://appextremes.sharepoint.com/...
//
// NOTE: This only works for PUBLIC or pre-authenticated sessions.
// SharePoint will redirect to login if the server has no auth cookie.
// For a demo, open SharePoint in the same browser first (cookies are separate).
// For a real solution, use Microsoft Graph API with a service account token.

app.get("/api/proxy", async (req, res) => {
  const targetUrl = req.query.url;
  if (!targetUrl) return res.status(400).send("Missing ?url= parameter");

  // Only allow SharePoint / OneDrive targets for security
  const allowed = /^https:\/\/([\w-]+\.(sharepoint\.com|onedrive\.com|office\.com|live\.com))/i;
  if (!allowed.test(targetUrl)) {
    return res.status(403).send("Only SharePoint/OneDrive URLs are allowed.");
  }

  try {
    // Forward cookies from the browser request so SharePoint auth works
    const cookieHeader = req.headers["cookie"] || "";

    const spRes = await fetch(targetUrl, {
      headers: {
        "User-Agent": req.headers["user-agent"] || "Mozilla/5.0",
        "Accept":     req.headers["accept"]      || "text/html",
        "Cookie":     cookieHeader
      },
      redirect: "follow",
      agent: new https.Agent({ rejectUnauthorized: true })
    });

    // Copy status + headers, removing the frame-blocking ones
    const STRIP_HEADERS = new Set([
      "x-frame-options",
      "content-security-policy",
      "content-security-policy-report-only"
    ]);

    spRes.headers.forEach((value, name) => {
      if (!STRIP_HEADERS.has(name.toLowerCase())) {
        try { res.setHeader(name, value); } catch {}
      }
    });

    // Allow framing from localhost
    res.setHeader("Content-Security-Policy", "frame-ancestors 'self' http://localhost:*");
    res.status(spRes.status);

    // Stream body back
    spRes.body.pipe(res);

  } catch (err) {
    res.status(500).send("Proxy error: " + err.message);
  }
});

// ── Keyword → likely tag fragments mapping ───────────────────────────────────

const KEYWORD_MAP = {
  payment:         ["payment", "paymentterms", "payment_clause", "fee"],
  termination:     ["termination", "termination_clause", "terminat"],
  obligation:      ["obligation", "obligations", "duty"],
  confidential:    ["confidentiality", "confidential", "nda", "non_disclosure"],
  liability:       ["liability", "limitation_of_liability", "limit_liability"],
  warranty:        ["warranty", "warranties", "representation"],
  indemnif:        ["indemnification", "indemnity", "indemnif"],
  intellectual:    ["ip", "intellectual_property", "ownership"],
  governing:       ["governing_law", "jurisdiction", "dispute"],
  force:           ["force_majeure", "acts_of_god"],
  renewal:         ["renewal", "auto_renew", "evergreen"],
  amendment:       ["amendment", "change_order", "modification"]
};

// ── Helpers ───────────────────────────────────────────────────────────────────

function findClause(keyword, contentControls) {
  const kw = keyword.toLowerCase().trim();

  // 1. Exact substring match on tag or title
  for (const cc of contentControls) {
    const tag   = (cc.tag   || "").toLowerCase();
    const title = (cc.title || "").toLowerCase();
    if (tag.includes(kw) || title.includes(kw)) return cc;
  }

  // 2. Keyword map match
  for (const [key, fragments] of Object.entries(KEYWORD_MAP)) {
    if (kw.includes(key) || key.includes(kw)) {
      for (const cc of contentControls) {
        const tag = (cc.tag || "").toLowerCase();
        if (fragments.some((f) => tag.includes(f))) return cc;
      }
    }
  }

  return null;
}

function formatClauseList(contentControls) {
  if (!contentControls.length) {
    return "No content controls found in the document. Make sure a Conga contract document is open in Word.";
  }
  const lines = contentControls.map((cc) => {
    const name = cc.tag || cc.title || "Unnamed";
    return `• ${name}  (ID: ${cc.id})`;
  });
  return `Found **${contentControls.length}** content controls:\n\n${lines.join("\n")}`;
}

// ── Parse Conga custom XML metadata from DOCX ────────────────────────────────
// Based on the reference metadata-viewer project, Conga CLM stores metadata in
// customXml/item*.xml (root-level in the DOCX ZIP, NOT under word/).
// Each <Node p2:id="[ccId]"> holds a base64 + GZIP-compressed XML string
// containing a <Metadata> element with all field properties as children.
// Document-level properties are stored in a <Properties> element.

/**
 * Dynamically extract ALL children of <Metadata> from decompressed XML.
 * Returns a plain object with every property present in that metadata node.
 */
function extractAllMetadataFields(xml) {
  // Find <Metadata>...</Metadata> block
  const metaMatch = xml.match(/<Metadata>([\s\S]*?)<\/Metadata>/);
  if (!metaMatch) return null;

  const inner = metaMatch[1];
  const meta = {};

  // Extract every <TagName>value</TagName> pair
  const tagRegex = /<([A-Za-z][A-Za-z0-9_]*)>([^<]*)<\/\1>/g;
  let m;
  while ((m = tagRegex.exec(inner)) !== null) {
    meta[m[1]] = m[2].trim();
  }

  // Must have at least a Type to be a valid Conga metadata node
  return meta["Type"] ? meta : null;
}

/**
 * Scan all customXml/item*.xml entries in the DOCX zip for Conga metadata.
 * Key fixes vs previous version:
 *   1. Path is customXml/item*.xml (root), NOT word/customXml/item*.xml
 *   2. Decompression is GZIP (gunzipSync), not zlib inflate
 *   3. All Metadata children extracted dynamically, not hardcoded fields
 *
 * Returns a Map<ccIdString, metaObject> for merging into content-control objects.
 * Also returns documentProps if a <Properties> node is found.
 */
function parseCustomXmlMetadata(zip) {
  const result = new Map();
  let documentProps = null;
  const CONGA_NS = "http://www.apttus.com/schemas";

  try {
    const entries = zip.getEntries().filter((e) =>
      /^customXml\/item\d+\.xml$/i.test(e.entryName)   // root-level, no word/ prefix
    );

    for (const entry of entries) {
      const xml = entry.getData().toString("utf8");
      if (!xml.includes(CONGA_NS)) continue;

      // ── CASE 1: Clause / Field metadata nodes ────────────────────────────
      const nodeRegex = /<Node\b[^>]*\bp2:id="([^"]+)"[^>]*>([\s\S]*?)<\/Node>/g;
      let nodeMatch;
      while ((nodeMatch = nodeRegex.exec(xml)) !== null) {
        const ccIdRaw = nodeMatch[1].trim();
        const encoded = nodeMatch[2].trim();
        if (!encoded) continue;

        let metaXml = null;
        try {
          const buf = Buffer.from(encoded, "base64");
          metaXml = zlib.gunzipSync(buf).toString("utf8");   // GZIP, not inflate
        } catch {
          // Fallback: may be plain (uncompressed) XML in older docs
          metaXml = encoded;
        }

        if (!metaXml) continue;

        const meta = extractAllMetadataFields(metaXml);
        if (!meta) continue;

        // Index by raw id AND unsigned-32 equivalent (Word can return signed ints)
        const register = (key) => { if (key) result.set(key, meta); };
        register(ccIdRaw);
        const asNum = parseInt(ccIdRaw, 10);
        if (!isNaN(asNum)) {
          register((asNum >>> 0).toString());
          register(asNum.toString());
        }
      }

      // ── CASE 2: Document-level properties ────────────────────────────────
      const propsMatch = xml.match(/<Properties>([\s\S]*?)<\/Properties>/);
      if (propsMatch) {
        const encoded = propsMatch[1].trim();
        if (encoded) {
          try {
            const buf = Buffer.from(encoded, "base64");
            const propsXml = zlib.gunzipSync(buf).toString("utf8");
            // Extract all root-element children as key/value pairs
            const props = {};
            const tagRegex = /<([A-Za-z][A-Za-z0-9_]*)>([^<]*)<\/\1>/g;
            let pm;
            while ((pm = tagRegex.exec(propsXml)) !== null) {
              props[pm[1]] = pm[2].trim();
            }
            if (Object.keys(props).length) documentProps = props;
          } catch { /* non-fatal */ }
        }
      }
    }
  } catch (err) {
    // Non-fatal — return whatever was collected
  }

  return { metaMap: result, documentProps };
}

// ── Build system prompt from document context ────────────────────────────────

function buildSystemPrompt(contentControls, docText, documentProps) {
  const ccLines = contentControls.map((cc) => {
    const type    = cc.type  || "Control";
    // Use human-readable clauseName when available; fall back to numeric tag
    const name    = cc.clauseName
      ? `${cc.clauseName} [Tag: ${cc.tag}]`
      : (cc.tag || cc.title || "Unnamed");
    const preview = cc.text  ? cc.text.substring(0, 300) : "(no text)";

    // Surface key metadata flags so the AI understands the contract state
    const flags = [
      cc.SubType            ? `SubType:${cc.SubType}`                   : null,
      cc.SFObjectName       ? `Object:${cc.SFObjectName}`               : null,
      cc.SFSourceAPI        ? `Field:${cc.SFSourceAPI}`                 : null,
      cc.Smart  === "true"  ? "Smart"                                   : null,
      cc.Dirty  === "true"  ? "Modified"                                : null,
      cc.Readonly === "true" || cc.cannotEdit ? "ReadOnly"              : null,
      cc.MarkedForDeletion === "true" ? "MarkedForDeletion"             : null,
      cc.Action             ? `Action:${cc.Action}`                     : null,
    ].filter(Boolean).join(", ");

    return `- "${name}" [${type}]${flags ? ` (${flags})` : ""}: ${preview}`;
  }).join("\n");

  // Include document-level properties (agreement type, status, etc.) if present
  const docPropsSection = documentProps
    ? `\nDocument Properties:\n${Object.entries(documentProps).map(([k,v]) => `  ${k}: ${v}`).join("\n")}`
    : "";

  const docSnippet = docText ? docText.substring(0, 3000) : "";

  return `You are Conga AI Assistance, a contract intelligence assistant for Conga CLM contracts.
${documentProps ? `\nDocument Properties:\n${Object.entries(documentProps).map(([k,v]) => `  ${k}: ${v}`).join("\n")}\n` : ""}
The document currently open has the following content controls (clauses and fields):
${ccLines || "(none detected yet)"}

Full document text (first 3000 chars):
${docSnippet || "(not available)"}

Your job:
1. Answer any question about this contract accurately, using only the information above.
2. If the user asks to find or navigate to a clause, include a JSON block at the END of your reply:
   {"action":"navigate","tag":"<exact tag name>"}
3. If the user asks to update/replace existing text in a clause, include:
   {"action":"update","tag":"<clause tag>","find":"<exact text to find>","replace":"<replacement text>"}
   IMPORTANT: "find" must NEVER be empty. It must contain the exact text you are replacing.
4. If the user asks to INSERT or ADD new text into a clause (not replace), use the text immediately before the insertion point as "find" and include the original text plus the new text as "replace".
   Example: insert "This is mandatory." after the first sentence → {"action":"update","tag":"967718","find":"first sentence text.","replace":"first sentence text. This is mandatory."}
5. If the user asks to insert or add a BRAND NEW clause or section to the document, include:
   {"action":"insert","tag":"<new clause name>","text":"<full text of the new clause>"}
6. If the user asks to list clauses or fields, list them clearly.
7. Be concise. Do not hallucinate content not present in the document.

IMPORTANT — Conga API tools:
• Authentication (Authorization header, ExternalToken) is handled AUTOMATICALLY by the server. NEVER ask the user for tokens, credentials, client_id, client_secret, or any auth details.
• For congaendreviewtool___endReviewerReview: just call it with reviewtype, reviewid, and reviewerid. The server will automatically fetch the ExternalToken (via getReviewAuthorization) and inject the Authorization header. NEVER ask the user for ExternalToken or Authorization.
• For congaendreviewtool___getReviewAuthorization: just call it with reviewtype. Authorization is injected automatically.
• NEVER call congaauth___createBearerToken manually — the app handles authentication behind the scenes.
• Only ask the user for BUSINESS parameters (reviewid, reviewerid, reviewtype, etc.).`;
}

// ── Build document context string for agent streaming ─────────────────────────

function buildDocumentContext(contentControls, docText, documentProps) {
  const ccLines = contentControls.map((cc) => {
    const type    = cc.type  || "Control";
    // Use human-readable clauseName when available; fall back to numeric tag
    const name    = cc.clauseName
      ? `${cc.clauseName} [Tag: ${cc.tag}]`
      : (cc.tag || cc.title || "Unnamed");
    const text    = cc.text  || "(no text)";
    const flags = [
      cc.SubType             ? `SubType:${cc.SubType}`                : null,
      cc.SFObjectName        ? `Object:${cc.SFObjectName}`            : null,
      cc.SFSourceAPI         ? `Field:${cc.SFSourceAPI}`              : null,
      cc.Smart  === "true"   ? "Smart"                                : null,
      cc.Dirty  === "true"   ? "Modified"                             : null,
      cc.Readonly === "true" || cc.cannotEdit ? "ReadOnly"            : null,
      cc.MarkedForDeletion === "true" ? "MarkedForDeletion"           : null,
      cc.Action              ? `Action:${cc.Action}`                  : null,
    ].filter(Boolean).join(", ");
    return `- "${name}" [${type}]${flags ? ` (${flags})` : ""}: ${text}`;
  }).join("\n");

  const docPropsSection = documentProps
    ? `\nDocument Properties:\n${Object.entries(documentProps).map(([k, v]) => `  ${k}: ${v}`).join("\n")}`
    : "";

  return `[CURRENT DOCUMENT CONTEXT — use this to answer the user]${docPropsSection}
Content controls (clauses and fields) in the open document:
${ccLines || "(none detected — make sure a Conga contract document is open in Word)"}`;
}

// ── Parse action JSON block from AI reply ─────────────────────────────────────

function parseActionFromReply(reply) {
  try {
    const match = reply.match(/\{\s*"action"\s*:.+?\}/s);
    if (match) return JSON.parse(match[0]);
  } catch {}
  return null;
}

// ── Chat Completions + Tool Calling loop ─────────────────────────────────────
// Handles tool calls from Chat Completions API. Document tools are resolved
// locally; MCP tools are proxied to the Foundry Toolbox endpoint.

async function executeToolCall(toolCall, contentControls) {
  const name = toolCall.function.name;
  let args;
  try { args = JSON.parse(toolCall.function.arguments); } catch { args = {}; }
  let action = null;
  let result = "";

  switch (name) {
    case "navigate_to_clause": {
      const found = findClause(args.tag, contentControls);
      if (found) {
        action = { type: "navigate", tag: found.tag || found.title, contentControlId: found.id };
        result = JSON.stringify({ success: true, tag: found.tag, id: found.id, text: (found.text || "").substring(0, 200) });
      } else {
        result = JSON.stringify({ success: false, error: `Clause "${args.tag}" not found in the document.` });
      }
      break;
    }
    case "update_clause": {
      const found = findClause(args.tag, contentControls);
      if (found) {
        action = { type: "update", tag: args.tag, find: args.find, replace: args.replace };
        result = JSON.stringify({ success: true, tag: found.tag, id: found.id });
      } else {
        result = JSON.stringify({ success: false, error: `Clause "${args.tag}" not found.` });
      }
      break;
    }
    case "insert_clause": {
      action = { type: "insert", tag: args.tag, text: args.text };
      result = JSON.stringify({ success: true, tag: args.tag });
      break;
    }
    case "delete_clause": {
      const found = findClause(args.tag, contentControls);
      if (found) {
        action = { type: "delete", tag: found.tag || found.title, contentControlId: found.id };
        result = JSON.stringify({ success: true, tag: found.tag, id: found.id });
      } else {
        result = JSON.stringify({ success: false, error: `Clause "${args.tag}" not found.` });
      }
      break;
    }
    case "list_clauses": {
      result = formatClauseList(contentControls);
      break;
    }
    case "search_clause": {
      const found = findClause(args.keyword, contentControls);
      if (found) {
        result = JSON.stringify({ found: true, tag: found.tag, id: found.id, type: found.type, text: (found.text || "").substring(0, 300) });
      } else {
        result = JSON.stringify({ found: false, message: `No clause matching "${args.keyword}" found.` });
      }
      break;
    }
    case "congaendreviewtool___endReviewerReview": {
      // Orchestrated flow: get ExternalToken via authorize, then call end review
      const reviewtype = args.reviewtype || "Office365";
      const reviewid = args.reviewid || "";
      const reviewerid = args.reviewerid || "";
      if (!reviewid || !reviewerid) {
        result = JSON.stringify({ success: false, error: "reviewid and reviewerid are required" });
        break;
      }
      try {
        // Step 1: Get ExternalToken via getReviewAuthorization
        let externalToken = args.ExternalToken || "";
        if (!externalToken) {
          logger.info("REVIEW_FLOW", { message: `Fetching ExternalToken for reviewtype=${reviewtype}` });
          externalToken = await getReviewExternalToken(reviewtype);
          if (!externalToken) {
            logger.warn("REVIEW_FLOW", { message: "ExternalToken could not be obtained from authorize endpoint" });
          }
        }
        // Step 2: Call endReviewerReview with all params + ExternalToken
        const enrichedArgs = await maybeInjectAuth(name, {
          ...args,
          reviewtype,
          reviewid,
          reviewerid,
          ExternalToken: externalToken,
        });
        logger.info("REVIEW_FLOW", { message: `Calling endReviewerReview: reviewid=${reviewid}, reviewerid=${reviewerid}` });
        const mcpResult = await callFoundryToolbox("tools/call", { name, arguments: enrichedArgs });
        logger.info("REVIEW_FLOW", { message: "endReviewerReview completed", result: JSON.stringify(mcpResult).substring(0, 500) });
        result = JSON.stringify({ success: true, result: mcpResult });
      } catch (err) {
        logger.error("REVIEW_FLOW", { message: `endReviewerReview failed: ${err.message}` });
        result = JSON.stringify({ success: false, error: err.message });
      }
      break;
    }
    default: {
      // Route unknown tools through Foundry Toolbox MCP endpoint
      if (FOUNDRY_TOOLBOX_URL) {
        try {
          const enrichedArgs = await maybeInjectAuth(name, args);
          const mcpResult = await callFoundryToolbox("tools/call", {
            name,
            arguments: enrichedArgs,
          });
          logger.info("MCP_TOOL", { message: `MCP tool: ${name}`, result: JSON.stringify(mcpResult).substring(0, 500) });
          result = JSON.stringify({ success: true, result: mcpResult });
        } catch (err) {
          logger.error("MCP_TOOL", { message: `MCP tool failed: ${name}`, error: err.message });
          result = JSON.stringify({ success: false, error: err.message });
        }
      } else {
        logger.warn("TOOL_CALL", { message: `Unknown tool: ${name}`, args });
        result = JSON.stringify({ error: `Unknown tool: ${name}`, args });
      }
    }
  }

  return { result, action };
}

/**
 * Run Chat Completions with a tool-call loop.
 * The model can call document tools or MCP tools; we execute them and feed
 * the results back until the model produces a final text response.
 */
async function chatWithTools(systemPrompt, userMessage, contentControls, history) {
  const allTools = [...DOCUMENT_TOOLS, ...mcpToolDefs];
  const messages = history && history.length
    ? [{ role: "system", content: systemPrompt }, ...history]
    : [
        { role: "system", content: systemPrompt },
        { role: "user",   content: userMessage },
      ];

  const actions = [];
  let maxIterations = 10; // safety limit

  while (maxIterations-- > 0) {
    const completion = await aiClient.chat.completions.create({
      model: AZURE_DEPLOYMENT,
      messages,
      tools: allTools.length ? allTools : undefined,
      temperature: 0.2,
      max_tokens: 1500,
    });

    const choice = completion.choices[0];
    const assistantMsg = choice.message;

    // If no tool calls, we're done
    if (!assistantMsg.tool_calls || assistantMsg.tool_calls.length === 0) {
      return {
        reply: assistantMsg.content || "No response.",
        actions,
        usage: completion.usage,
      };
    }

    // Push assistant message (with tool_calls) to conversation
    messages.push({
      role: "assistant",
      content: assistantMsg.content ?? "",
      tool_calls: assistantMsg.tool_calls,
    });

    const toolNames = assistantMsg.tool_calls.map(tc => tc.function.name);
    logger.info("TOOL_CALLS", { message: `Model calling tools: ${toolNames.join(", ")}` });

    // Execute each tool call and push results
    for (const tc of assistantMsg.tool_calls) {
      const { result, action } = await executeToolCall(tc, contentControls);
      if (action) actions.push(action);
      messages.push({
        role: "tool",
        tool_call_id: tc.id,
        content: result,
      });
      logger.debug("TOOL_RESULT", { message: `Tool ${tc.function.name}`, result: result.substring(0, 300) });
    }
  }

  // If we exhausted iterations, return the last content
  return { reply: "Tool call loop exceeded maximum iterations.", actions, usage: null };
}

// ── POST /api/chat ────────────────────────────────────────────────────────────

app.post("/api/chat", async (req, res) => {
  const { message = "", contentControls = [], docText = "", threadId: reqThreadId = null } = req.body;
  let threadId = reqThreadId;
  const reqId = Date.now().toString(36);

  logger.info("CHAT_REQ", {
    message: `Chat request [${reqId}]`,
    reqId,
    userMessage: message,
    contentControlCount: contentControls.length,
    contentControls: contentControls.map(cc => ({ id: cc.id, tag: cc.tag, type: cc.type })),
    docTextLength: docText.length,
    mode: aiClient ? "azure-openai" : "keyword-fallback"
  });

  // ── Agent path (Azure AI Foundry — Assistants API with persistent threads) ──
  if (assistantsClient && AZURE_AGENT_ID) {
    try {
      if (!threadId) {
        const thread = await assistantsClient.beta.threads.create();
        threadId = thread.id;
      }

      await assistantsClient.beta.threads.messages.create(threadId, {
        role: "user",
        content: message
      });

      const docContext = buildDocumentContext(contentControls, docText, null);

      logger.debug("AGENT_REQ", {
        message: `Agent run [${reqId}]`,
        reqId, threadId,
        agentId: AZURE_AGENT_ID,
        docContextLength: docContext.length
      });

      const run = await assistantsClient.beta.threads.runs.createAndPoll(threadId, {
        assistant_id: AZURE_AGENT_ID,
        additional_instructions: docContext
      });

      if (run.status !== "completed") {
        throw new Error(`Agent run ended with status: ${run.status}. ${run.last_error?.message || ""}`);
      }

      const msgList  = await assistantsClient.beta.threads.messages.list(threadId, { limit: 1, order: "desc" });
      const latestMsg = msgList.data[0];
      const rawReply  = latestMsg?.content
        ?.filter((c) => c.type === "text")
        .map((c) => c.text.value)
        .join("") || "No response from agent.";

      logger.info("AGENT_RES", {
        message: `Agent response [${reqId}]`,
        reqId, threadId,
        promptTokens: run.usage?.prompt_tokens,
        completionTokens: run.usage?.completion_tokens,
        rawReply
      });

      const action     = parseActionFromReply(rawReply);
      const cleanReply = rawReply.replace(/\{\s*"action"\s*:.+?\}/s, "").trim();

      const response = { reply: cleanReply, threadId };

      if (action?.action === "navigate") {
        const found = findClause(action.tag, contentControls);
        response.navigateTo       = action.tag;
        response.contentControlId = found?.id;
      }
      if (action?.action === "update") {
        response.updateAction = { tag: action.tag, find: action.find, replace: action.replace };
      }
      if (action?.action === "insert") {
        response.insertAction = { tag: action.tag, text: action.text };
      }

      logger.info("CHAT_RES", {
        message: `Chat response [${reqId}]`,
        reqId,
        replyLength: cleanReply.length,
        action: action?.action || null,
        navigateTo: response.navigateTo || null
      });

      return res.json(response);

    } catch (err) {
      logger.error("AGENT_ERR", { message: `Agent error [${reqId}]`, reqId, threadId, error: err.message });
      // Fall through to stateless chat.completions fallback below
    }
  }

  // ── Chat Completions + Tool Calling ─────────────────────────────────────
  const hasTools = DOCUMENT_TOOLS.length > 0 || mcpToolDefs.length > 0;
  if (aiClient && hasTools) {
    try {
      const systemPrompt = buildSystemPrompt(contentControls, docText, null);

      logger.debug("TOOL_CHAT_REQ", {
        message: `Tool-calling request [${reqId}]`,
        reqId,
        deployment: AZURE_DEPLOYMENT,
        toolCount: DOCUMENT_TOOLS.length + mcpToolDefs.length,
        userMessage: message,
        threadId
      });

      const { reply, actions, usage } = await chatWithTools(
        systemPrompt, message, contentControls
      );

      logger.info("TOOL_CHAT_RES", {
        message: `Tool-calling response [${reqId}]`,
        reqId,
        replyLength: reply.length,
        actions: actions.map(a => a.type),
        promptTokens: usage?.prompt_tokens,
        completionTokens: usage?.completion_tokens,
      });

      // Build response for client
      const response = { reply, threadId };

      for (const action of actions) {
        if (action.type === "navigate") {
          response.navigateTo       = action.tag;
          response.contentControlId = action.contentControlId;
        }
        if (action.type === "update") {
          response.updateAction = { tag: action.tag, find: action.find, replace: action.replace };
        }
        if (action.type === "insert") {
          response.insertAction = { tag: action.tag, text: action.text };
        }
        if (action.type === "delete") {
          response.deleteAction = { tag: action.tag, contentControlId: action.contentControlId };
        }
      }

      return res.json(response);

    } catch (err) {
      logger.error("TOOL_CHAT_ERR", { message: `Tool-calling error [${reqId}]`, reqId, error: err.message });
      // Fall through to simple prompt agent
    }
  }

  // ── Prompt Agent fallback (Chat Completions API) ────────────────────────
  if (aiClient) {
    try {
      const systemPrompt = buildSystemPrompt(contentControls, docText, null);

      logger.debug("OPENAI_REQ", {
        message: `OpenAI request [${reqId}]`,
        reqId,
        deployment: AZURE_DEPLOYMENT,
        systemPromptLength: systemPrompt.length,
        userMessage: message
      });

      const completion = await aiClient.chat.completions.create({
        model: AZURE_DEPLOYMENT,
        messages: [
          { role: "system",  content: systemPrompt },
          { role: "user",    content: message }
        ],
        temperature: 0.2,
        max_tokens: 800
      });

      const rawReply = completion.choices[0]?.message?.content || "No response from AI.";
      const usage    = completion.usage || {};

      logger.info("OPENAI_RES", {
        message: `OpenAI response [${reqId}]`,
        reqId,
        promptTokens: usage.prompt_tokens,
        completionTokens: usage.completion_tokens,
        totalTokens: usage.total_tokens,
        rawReply
      });

      const action      = parseActionFromReply(rawReply);
      const cleanReply  = rawReply.replace(/\{\s*"action"\s*:.+?\}/s, "").trim();

      const response = { reply: cleanReply, threadId };

      if (action?.action === "navigate") {
        const found = findClause(action.tag, contentControls);
        response.navigateTo       = action.tag;
        response.contentControlId = found?.id;
      }

      if (action?.action === "update") {
        response.updateAction = { tag: action.tag, find: action.find, replace: action.replace };
      }

      if (action?.action === "insert") {
        response.insertAction = { tag: action.tag, text: action.text };
      }

      return res.json(response);

    } catch (err) {
      logger.error("OPENAI_ERR", { message: `OpenAI error [${reqId}]`, reqId, error: err.message });
      // Fall through to keyword fallback on AI error
    }
  }

  // ── Keyword fallback (no AI configured or AI error) ──────────────────────
  const msg = message.toLowerCase();

  if (msg.includes("list") || msg.includes("all clause") || msg.includes("show clause") || msg.includes("what clause")) {
    return res.json({ reply: formatClauseList(contentControls), threadId });
  }

  const findMatch =
    msg.match(/(?:find|locate|go to|navigate to|jump to)\s+(?:the\s+)?(.+?)(?:\s+clause)?$/) ||
    msg.match(/where is\s+(?:the\s+)?(.+?)(?:\s+clause)?$/);

  if (findMatch) {
    const keyword = findMatch[1].replace(/clause/i, "").trim();
    const found   = findClause(keyword, contentControls);
    if (found) {
      return res.json({
        reply: `Found **${found.tag || found.title}** (ID: ${found.id}).\n\n"${(found.text || "").substring(0, 200)}…"`,
        contentControlId: found.id,
        navigateTo: found.tag || found.title,
        threadId
      });
    }
    return res.json({ reply: `Could not find "${keyword}". Try "List all clauses" to see what's available.`, threadId });
  }

  const updateMatch = msg.match(/(?:change|update|replace|set)\s+(.+?)\s+(?:from\s+)?["']?(.+?)["']?\s+to\s+["']?(.+?)["']?$/);
  if (updateMatch) {
    return res.json({
      reply: `Updating "${updateMatch[2]}" → "${updateMatch[3]}". Click Apply below.`,
      updateAction: { tag: updateMatch[1], find: updateMatch[2], replace: updateMatch[3] },
      threadId
    });
  }

  res.json({
    reply: `⚠️ Azure OpenAI is not configured. Add your keys to api/.env to enable AI responses.\n\nKeyword mode — try:\n• "List all clauses"\n• "Find the payment clause"\n• "Change 30 days to 60 days"`,
    threadId
  });
});

// ── POST /api/parse-xml-metadata ──────────────────────────────────────────────
// Accepts raw custom XML strings from Office.js customXmlParts and returns a
// map of { [ccId]: clauseName } by decoding the Conga GZIP-compressed metadata.

app.post("/api/parse-xml-metadata", (req, res) => {
  const { xmlStrings = [] } = req.body;
  const tagMap = {};

  for (const xml of xmlStrings) {
    if (typeof xml !== "string") continue;
    const CONGA_NS = "http://www.apttus.com/schemas";
    if (!xml.includes(CONGA_NS)) continue;

    const nodeRegex = /<Node\b[^>]*\bp2:id="([^"]+)"[^>]*>([\s\S]*?)<\/Node>/g;
    let nodeMatch;
    while ((nodeMatch = nodeRegex.exec(xml)) !== null) {
      const ccIdRaw = nodeMatch[1].trim();
      const encoded = nodeMatch[2].trim();
      if (!encoded) continue;

      let metaXml = null;
      try {
        const buf = Buffer.from(encoded, "base64");
        metaXml = zlib.gunzipSync(buf).toString("utf8");
      } catch {
        metaXml = encoded;
      }
      if (!metaXml) continue;

      const tagMatch = metaXml.match(/<Tag>([^<]*)<\/Tag>/);
      if (!tagMatch) continue;
      const clauseName = tagMatch[1].trim();
      if (!clauseName) continue;

      const register = (key) => { if (key) tagMap[key] = clauseName; };
      register(ccIdRaw);
      const asNum = parseInt(ccIdRaw, 10);
      if (!isNaN(asNum)) {
        register((asNum >>> 0).toString());
        register(asNum.toString());
      }
    }
  }

  res.json({ tagMap });
});

// ── POST /api/risk-scan ──────────────────────────────────────────────────────
// Analyses every clause in the document and returns a risk score for each one.

app.post("/api/risk-scan", async (req, res) => {
  const { contentControls = [] } = req.body;
  const reqId = Date.now().toString(36);

  const SKIP_TYPES = new Set(["field", "repeat", "segment"]);
  const clauses = contentControls.filter(
    (cc) => !SKIP_TYPES.has((cc.type || "").toLowerCase()) && (cc.tag || cc.title)
  );

  if (!clauses.length) return res.json({ clauses: [] });

  if (!aiClient && !assistantsClient) {
    return res.json({
      clauses: clauses.map((cc) => ({
        tag: cc.tag || cc.title, id: cc.id, riskLevel: "unknown",
        reason: "AI not configured", riskFactors: []
      }))
    });
  }

  const clauseList = clauses.map((cc, i) =>
    `[${i + 1}] Tag: "${cc.tag || cc.title}" (ID: ${cc.id})\nText: ${cc.text || "(empty)"}`
  ).join("\n\n");

  const systemPrompt = `You are a contract risk analyst. Analyse each clause below and return ONLY a valid JSON array — no prose, no markdown, no code fences.

Each element must have exactly these fields:
  "tag"         — the Tag value from the input (string)
  "id"          — the ID value from the input (number)
  "riskLevel"   — one of: "high", "medium", "low"
  "reason"      — one concise sentence explaining the risk rating (max 12 words)
  "riskFactors" — array of up to 3 short risk factor labels

Risk guidelines:
  HIGH   — uncapped liability, unilateral termination without cause, broad IP assignment, indemnification without limits, auto-renewal with <30 day opt-out window
  MEDIUM — payment terms >60 days, liquidated damages, exclusivity clauses, non-compete >1 year, limited warranties
  LOW    — standard confidentiality, reasonable notice periods, mutual obligations, balanced termination

Return ONLY the JSON array.`;

  try {
    const completion = await aiClient.chat.completions.create({
      model: AZURE_DEPLOYMENT,
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user",   content: clauseList }
      ],
      temperature: 0.1,
      max_tokens: 2000
    });

    let raw = completion.choices[0]?.message?.content || "[]";
    raw = raw.replace(/```json\s*/gi, "").replace(/```\s*/g, "").trim();
    let parsed;
    try {
      const obj = JSON.parse(raw);
      parsed = Array.isArray(obj) ? obj : (obj.clauses || obj.results || Object.values(obj).find(v => Array.isArray(v)) || []);
    } catch {
      const match = raw.match(/\[[\s\S]*\]/);
      parsed = match ? JSON.parse(match[0]) : [];
    }

    logger.info("RISK_SCAN", { message: `Risk scan [${reqId}] — ${parsed.length} clauses scored`, reqId });
    return res.json({ clauses: parsed });

  } catch (err) {
    logger.error("RISK_SCAN_ERR", { message: `Risk scan error [${reqId}]`, reqId, error: err.message });
    try {
      const completion = await aiClient.chat.completions.create({
        model: AZURE_DEPLOYMENT,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user",   content: clauseList }
        ],
        temperature: 0.1,
        max_tokens: 2000
      });
      let raw = completion.choices[0]?.message?.content || "[]";
      const match = raw.match(/\[[\s\S]*\]/);
      const parsed = match ? JSON.parse(match[0]) : [];
      return res.json({ clauses: parsed });
    } catch (err2) {
      logger.error("RISK_SCAN_ERR2", { message: `Risk scan retry failed [${reqId}]`, reqId, error: err2.message });
      return res.status(500).json({ error: "Risk scan failed", clauses: [] });
    }
  }
});

// ── POST /api/chat/stream ─────────────────────────────────────────────────────
// Streaming version of /api/chat — uses Server-Sent Events (SSE).

app.post("/api/chat/stream", async (req, res) => {
  const { message = "", contentControls = [], docText = "", threadId: reqThreadId = null } = req.body;
  let threadId = reqThreadId;
  const reqId = Date.now().toString(36);

  // Set SSE headers immediately so the browser can start reading
  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders();

  const sendEvent = (data) => {
    if (!res.writableEnded) res.write(`data: ${JSON.stringify(data)}\n\n`);
  };
  const endStream = () => { if (!res.writableEnded) res.end(); };

  // ── Agent path (Assistants API streaming) ──────────────────────────────────
  if (assistantsClient && AZURE_AGENT_ID) {
    try {
      if (!threadId) {
        const thread = await assistantsClient.beta.threads.create();
        threadId = thread.id;
      }

      await assistantsClient.beta.threads.messages.create(threadId, {
        role: "user",
        content: message
      });

      const docContext = buildDocumentContext(contentControls, docText, null);
      let fullText = "";

      const runStream = assistantsClient.beta.threads.runs.stream(threadId, {
        assistant_id: AZURE_AGENT_ID,
        additional_instructions: docContext
      });

      runStream.on("textDelta", (delta) => {
        const token = delta.value || "";
        if (token) {
          fullText += token;
          sendEvent({ type: "token", text: token });
        }
      });

      runStream.on("error", (err) => {
        logger.error("AGENT_STREAM_ERR", { message: `Agent stream error [${reqId}]`, reqId, error: err.message });
        sendEvent({ type: "error", message: "Stream error — please try again." });
        endStream();
      });

      await runStream.finalRun();

      const action     = parseActionFromReply(fullText);
      const cleanReply = fullText.replace(/\{\s*"action"\s*:.+?\}/s, "").trim();
      const done = { type: "done", reply: cleanReply, threadId };

      if (action?.action === "navigate") {
        const found = findClause(action.tag, contentControls);
        done.navigateTo       = action.tag;
        done.contentControlId = found?.id;
      }
      if (action?.action === "update") {
        done.updateAction = { tag: action.tag, find: action.find, replace: action.replace };
      }
      if (action?.action === "insert") {
        done.insertAction = { tag: action.tag, text: action.text };
      }

      sendEvent(done);
      endStream();
      return;

    } catch (err) {
      logger.error("AGENT_STREAM_FAIL", { message: `Agent stream failed [${reqId}], falling back`, reqId, error: err.message });
    }
  }

  // ── chat.completions path (streaming) ──────────────────────────────────────
  if (aiClient) {
    try {
      const systemPrompt = buildSystemPrompt(contentControls, docText, null);
      let fullText = "";

      const stream = await aiClient.chat.completions.create({
        model: AZURE_DEPLOYMENT,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user",   content: message }
        ],
        temperature: 0.2,
        max_tokens: 800,
        stream: true
      });

      for await (const chunk of stream) {
        const token = chunk.choices[0]?.delta?.content || "";
        if (token) {
          fullText += token;
          sendEvent({ type: "token", text: token });
        }
      }

      const action     = parseActionFromReply(fullText);
      const cleanReply = fullText.replace(/\{\s*"action"\s*:.+?\}/s, "").trim();
      const done = { type: "done", reply: cleanReply, threadId };

      if (action?.action === "navigate") {
        const found = findClause(action.tag, contentControls);
        done.navigateTo       = action.tag;
        done.contentControlId = found?.id;
      }
      if (action?.action === "update") {
        done.updateAction = { tag: action.tag, find: action.find, replace: action.replace };
      }
      if (action?.action === "insert") {
        done.insertAction = { tag: action.tag, text: action.text };
      }

      sendEvent(done);
      endStream();
      return;

    } catch (err) {
      logger.error("COMPLETIONS_STREAM_FAIL", { message: `Completions stream failed [${reqId}]`, reqId, error: err.message });
      sendEvent({ type: "error", message: "AI error — please try again." });
      endStream();
      return;
    }
  }

  // ── Keyword fallback (no AI) ──────────────────────────────────────────────
  const msg_lc = message.toLowerCase();
  let fallbackReply = `⚠️ Azure OpenAI is not configured. Add your keys to api/.env to enable AI responses.\n\nKeyword mode — try:\n• "List all clauses"\n• "Find the payment clause"`;
  if (msg_lc.includes("list") || msg_lc.includes("all clause") || msg_lc.includes("show clause")) {
    fallbackReply = formatClauseList(contentControls);
  }
  sendEvent({ type: "done", reply: fallbackReply, threadId });
  endStream();
});

// ── POST /api/metadata ────────────────────────────────────────────────────────
// Accepts a .docx file upload, extracts content controls from document.xml

app.post("/api/metadata", upload.single("file"), (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded" });

  try {
    const zip     = new AdmZip(req.file.buffer);
    const xmlEntry = zip.getEntry("word/document.xml");
    if (!xmlEntry) return res.json({ contentControls: [] });

    const xml = xmlEntry.getData().toString("utf8");
    const controls = [];

    // Extract sdt (structured document tags = content controls) with regex
    // w:tag   = the Conga name (e.g. "Payment Terms", "Contract End Date")
    // w:alias = the Conga type (e.g. "Clause", "Field", "Repeat")
    const sdtRegex = /<w:sdt\b[\s\S]*?<\/w:sdt>/g;
    let sdtMatch;
    while ((sdtMatch = sdtRegex.exec(xml)) !== null) {
      const sdt = sdtMatch[0];

      const tagMatch   = sdt.match(/<w:tag\s+w:val="([^"]+)"/);   // name
      const idMatch    = sdt.match(/<w:id\s+w:val="([^"]+)"/);    // numeric id
      const aliasMatch = sdt.match(/<w:alias\s+w:val="([^"]+)"/); // type: Clause/Field/Repeat

      // Skip non-Conga controls (no tag, no Conga alias)
      const CONGA_TYPES = ["Clause", "Field", "Repeat", "Segment"];
      const alias = aliasMatch ? aliasMatch[1] : "";
      const tag   = tagMatch   ? tagMatch[1]   : "";
      if (!tag && !CONGA_TYPES.includes(alias)) continue;

      // Get plain text from the content
      const textParts = [];
      const tRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
      let tMatch;
      while ((tMatch = tRegex.exec(sdt)) !== null) textParts.push(tMatch[1]);

      controls.push({
        id:    idMatch ? parseInt(idMatch[1], 10) : Math.floor(Math.random() * 10000),
        tag,
        type:  alias,                    // "Clause", "Field", "Repeat", or "Segment"
        title: tag ? `${tag}${alias ? " (" + alias + ")" : ""}` : alias || "Unnamed",
        text:  textParts.join(" ").substring(0, 500)
      });
    }

    // ── Enrich with Conga custom XML metadata ──────────────────────────────
    // customXml/item*.xml (root-level in ZIP) holds GZIP-compressed Metadata
    // for each content control, plus optional document-level Properties.
    const { metaMap, documentProps } = parseCustomXmlMetadata(zip);
    for (const ctrl of controls) {
      const meta = metaMap.get(ctrl.id.toString()) || metaMap.get((ctrl.id >>> 0).toString());
      if (meta) {
        // Use Type from custom XML when available (more precise than w:alias)
        if (meta.Type) ctrl.type = meta.Type;
        // Merge ALL metadata fields directly onto the control object
        // (keys match the raw Metadata XML element names: Tag, SubType, SFSource, etc.)
        Object.assign(ctrl, meta);
      }
    }

    res.json({ contentControls: controls, total: controls.length, documentProps });
    logger.info("METADATA", {
      message: `Parsed metadata: ${controls.length} controls found`,
      total: controls.length,
      clauses: controls.filter(c => (c.type || "").toLowerCase() === "clause").length,
      fields:  controls.filter(c => (c.type || "").toLowerCase() === "field").length,
      repeats: controls.filter(c => (c.type || "").toLowerCase() === "repeat").length,
      customXmlEnriched: controls.filter(c => c.Tag).length,
      hasDocumentProps: !!documentProps
    });

  } catch (err) {
    logger.error("METADATA", { message: "Failed to parse document", error: err.message });
    res.status(500).json({ error: "Failed to parse document: " + err.message });
  }
});

// ── POST /api/end-review ──────────────────────────────────────────────────────
// Directly calls the Conga End Review MCP tool using env-configured IDs.
// No AI involved — just calls the tool and returns the result.

app.post("/api/end-review", async (req, res) => {
  const reviewtype = "Office365";
  const reviewid   = process.env.CONGA_REVIEW_ID   || "";
  const reviewerid = process.env.CONGA_REVIEWER_ID  || "";
  const reqId = Date.now().toString(36);

  if (!reviewid || !reviewerid) {
    return res.status(400).json({ success: false, error: "CONGA_REVIEW_ID and CONGA_REVIEWER_ID must be set in api/.env" });
  }

  try {
    // Step 1: Get ExternalToken
    logger.info("END_REVIEW", { message: `Getting ExternalToken for reviewtype=${reviewtype}`, reqId });
    const externalToken = await getReviewExternalToken(reviewtype);

    // Step 2: Call endReviewerReview with auth + external token
    const token = await getCongaToken();
    const args = {
      reviewtype,
      reviewid,
      reviewerid,
      ExternalToken: externalToken || "",
      Authorization: `Bearer ${token}`,
    };

    logger.info("END_REVIEW", { message: `Calling endReviewerReview: reviewid=${reviewid}, reviewerid=${reviewerid}`, reqId });
    const mcpResult = await callFoundryToolbox("tools/call", {
      name: "congaendreviewtool___endReviewerReview",
      arguments: args,
    });

    logger.info("END_REVIEW", { message: "endReviewerReview completed", reqId, result: JSON.stringify(mcpResult).substring(0, 500) });

    // Check JSON-RPC level error — extract a friendly message from the MCP wrapper
    if (mcpResult.error) {
      const rawErrMsg = mcpResult.error.message || JSON.stringify(mcpResult.error);
      logger.error("END_REVIEW_ERR", { message: `MCP JSON-RPC error [${reqId}]`, reqId, error: rawErrMsg });
      // Parse the MCP error to extract a user-friendly message
      let friendlyError = rawErrMsg;
      const detailMatch = rawErrMsg.match(/"detail"\s*:\s*"([^"]+)"/);
      if (detailMatch) {
        friendlyError = detailMatch[1];
      } else {
        const httpMatch = rawErrMsg.match(/HTTP error (\d+):\s*([^'"]+)/);
        if (httpMatch) {
          const status = httpMatch[1];
          if (status === "400") friendlyError = "The review may already be completed or is in an invalid state for this action.";
          else if (status === "401" || status === "403") friendlyError = "Authorization failed. Check your Conga credentials.";
          else if (status === "404") friendlyError = "Review or reviewer not found. Check the Review ID and Reviewer ID.";
          else friendlyError = `Conga API error (HTTP ${status}): ${httpMatch[2].trim()}`;
        }
      }
      return res.status(400).json({ success: false, error: friendlyError });
    }

    // Check MCP tool-level error (isError: true in result)
    const toolResult = mcpResult.result ?? mcpResult;
    if (toolResult.isError) {
      const errText = toolResult.content?.map(c => c.text).join(" ") || "Tool returned an error.";
      logger.error("END_REVIEW_ERR", { message: `MCP tool error [${reqId}]`, reqId, error: errText });
      return res.status(500).json({ success: false, error: errText });
    }

    return res.json({ success: true, message: "Review ended successfully.", result: mcpResult });

  } catch (err) {
    logger.error("END_REVIEW_ERR", { message: `End review failed [${reqId}]`, reqId, error: err.message });
    return res.status(500).json({ success: false, error: err.message });
  }
});

// ── POST /api/suggest-fix ─────────────────────────────────────────────────────
// Given a single risky clause, returns a rewritten version that mitigates the risk.
// Request:  { tag, id, text, riskLevel, reason, riskFactors }
// Response: { suggestedText }

app.post("/api/suggest-fix", async (req, res) => {
  const { tag, id, text, riskLevel, reason, riskFactors = [] } = req.body;

  if (!text) return res.status(400).json({ error: "Clause text is required" });
  if (!aiClient) return res.status(503).json({ error: "AI not configured" });

  const reqId = Date.now().toString(36);

  const systemPrompt = `You are a contract lawyer specialising in risk mitigation. Your job is to rewrite a single contract clause to reduce its legal and commercial risk while preserving its intent.

Rules:
- Return ONLY the rewritten clause text — no explanations, no preamble, no markdown
- Keep the same general structure and length as the original
- Fix the specific issues identified in the risk assessment
- Use balanced, professional contract language
- Do not add headings or labels`;

  const userPrompt = `Rewrite the following clause to reduce its risk.

Clause name: "${tag}"
Risk level: ${riskLevel}
Risk reason: ${reason}
Risk factors: ${riskFactors.join(", ")}

Original clause text:
${text}`;

  try {
    const completion = await aiClient.chat.completions.create({
      model: AZURE_DEPLOYMENT,
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user",   content: userPrompt }
      ],
      temperature: 0.2,
      max_tokens: 1000
    });

    const suggestedText = (completion.choices[0]?.message?.content || "").trim();
    logger.info("SUGGEST_FIX", { message: `Suggest fix [${reqId}] for clause id:${id}`, reqId });
    return res.json({ suggestedText });

  } catch (err) {
    logger.error("SUGGEST_FIX_ERR", { message: `Suggest fix error [${reqId}]`, reqId, error: err.message });
    return res.status(500).json({ error: "Failed to generate suggestion: " + err.message });
  }
});

// ── GET /api/health ───────────────────────────────────────────────────────────

app.get("/api/health", (_req, res) => {
  res.json({ status: "ok", service: "Conga AI Assistance API", timestamp: new Date().toISOString() });
});

// ── Start ─────────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  logger.info("STARTUP", { message: `Conga AI Assistance API started on port ${PORT}`, port: PORT, aiEnabled: !!aiClient, deployment: AZURE_DEPLOYMENT });
  console.log(`\nConga AI Assistance API  →  http://localhost:${PORT}`);
  console.log(`Health check         →  http://localhost:${PORT}/api/health`);
  console.log(`Logs                 →  ${LOGS_DIR}\n`);
});
