/**
 * DocSense AI — API Server
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
    const name    = cc.tag   || cc.title || "Unnamed";
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

  return `You are DocSense AI, a contract intelligence assistant for Conga CLM contracts.
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
7. Be concise. Do not hallucinate content not present in the document.`;
}

// ── Parse action JSON block from AI reply ─────────────────────────────────────

function parseActionFromReply(reply) {
  try {
    const match = reply.match(/\{\s*"action"\s*:.+?\}/s);
    if (match) return JSON.parse(match[0]);
  } catch {}
  return null;
}

// ── POST /api/chat ────────────────────────────────────────────────────────────

app.post("/api/chat", async (req, res) => {
  const { message = "", contentControls = [], docText = "" } = req.body;
  const reqId = Date.now().toString(36); // short request ID for correlation

  // Log the incoming chat request
  logger.info("CHAT_REQ", {
    message: `Chat request [${reqId}]`,
    reqId,
    userMessage: message,
    contentControlCount: contentControls.length,
    contentControls: contentControls.map(cc => ({ id: cc.id, tag: cc.tag, type: cc.type })),
    docTextLength: docText.length,
    mode: aiClient ? "azure-openai" : "keyword-fallback"
  });

  // ── AI path (Azure OpenAI) ──────────────────────────────────────────────
  if (aiClient) {
    try {
      const systemPrompt = buildSystemPrompt(contentControls, docText, null);

      logger.debug("OPENAI_REQ", {
        message: `OpenAI request [${reqId}]`,
        reqId,
        deployment: AZURE_DEPLOYMENT,
        systemPromptLength: systemPrompt.length,
        systemPrompt,
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

      // Strip action JSON from visible reply, parse it for navigation/update
      const action      = parseActionFromReply(rawReply);
      const cleanReply  = rawReply.replace(/\{\s*"action"\s*:.+?\}/s, "").trim();

      const response = { reply: cleanReply };

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
      logger.error("OPENAI_ERR", { message: `OpenAI error [${reqId}]`, reqId, error: err.message });
      // Fall through to keyword fallback on AI error
    }
  }

  // ── Keyword fallback (no AI configured or AI error) ──────────────────────
  const msg = message.toLowerCase();

  if (msg.includes("list") || msg.includes("all clause") || msg.includes("show clause") || msg.includes("what clause")) {
    return res.json({ reply: formatClauseList(contentControls) });
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
        navigateTo: found.tag || found.title
      });
    }
    return res.json({ reply: `Could not find "${keyword}". Try "List all clauses" to see what's available.` });
  }

  const updateMatch = msg.match(/(?:change|update|replace|set)\s+(.+?)\s+(?:from\s+)?["']?(.+?)["']?\s+to\s+["']?(.+?)["']?$/);
  if (updateMatch) {
    return res.json({
      reply: `Updating "${updateMatch[2]}" → "${updateMatch[3]}". Click Apply below.`,
      updateAction: { tag: updateMatch[1], find: updateMatch[2], replace: updateMatch[3] }
    });
  }

  res.json({
    reply: `⚠️ Azure OpenAI is not configured. Add your keys to api/.env to enable AI responses.\n\nKeyword mode — try:\n• "List all clauses"\n• "Find the payment clause"\n• "Change 30 days to 60 days"`
  });
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

// ── GET /api/health ───────────────────────────────────────────────────────────

app.get("/api/health", (_req, res) => {
  res.json({ status: "ok", service: "DocSense AI Mock API", timestamp: new Date().toISOString() });
});

// ── Start ─────────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  logger.info("STARTUP", { message: `DocSense AI API started on port ${PORT}`, port: PORT, aiEnabled: !!aiClient, deployment: AZURE_DEPLOYMENT });
  console.log(`\nDocSense AI Mock API  →  http://localhost:${PORT}`);
  console.log(`Health check         →  http://localhost:${PORT}/api/health`);
  console.log(`Logs                 →  ${LOGS_DIR}\n`);
});
