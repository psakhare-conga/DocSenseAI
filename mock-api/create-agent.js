/**
 * One-time script: creates the DocSense AI agent in Azure OpenAI
 * and prints the AGENT_ID to add to .env
 *
 * Run: node create-agent.js
 */
require("dotenv").config();
const { AzureOpenAI } = require("openai");

const ENDPOINT = process.env.AZURE_OPENAI_ENDPOINT || "";
const API_KEY  = process.env.AZURE_OPENAI_API_KEY  || "";

if (!ENDPOINT || !API_KEY) {
  console.error("Missing AZURE_OPENAI_ENDPOINT or AZURE_OPENAI_API_KEY in .env");
  process.exit(1);
}

// Assistants API requires the classic base endpoint (without /openai/v1 path)
// Strip any /openai/v1 or /v1 suffix — the SDK adds the correct paths
const baseEndpoint = ENDPOINT.replace(/\/openai\/v1\/?$/, "").replace(/\/v1\/?$/, "").replace(/\/+$/, "");

const client = new AzureOpenAI({
  endpoint:   baseEndpoint,
  apiKey:     API_KEY,
  apiVersion: "2024-05-01-preview",
});

// ── Static instructions stored permanently in the agent ─────────────────────
// Everything here is NOT billed as input tokens per call.
// Only dynamic document context (clause list) is sent per request.
const INSTRUCTIONS = `You are DocSense AI, a contract intelligence assistant for Conga CLM contracts.

BEHAVIOR RULES:
- Answer only from the document context provided by the user in each message.
- Never hallucinate clause content not present in the document.
- Be concise and professional.
- When listing clauses, use a numbered list with clause name and tag.
- If a clause is not found, say so clearly and suggest the user try "List all clauses".
- Maintain context across the conversation — remember what the user has asked before.

ACTIONS:
When the user requests navigation, updates, inserts, or deletes, emit exactly ONE JSON block
at the very END of your reply (after all human-readable text). Never embed JSON mid-reply.

1. Navigate to a clause:
   {"action":"navigate","tag":"<exact tag name>"}

2. Update / replace existing text in a clause:
   {"action":"update","tag":"<clause tag>","find":"<exact existing text>","replace":"<new text>"}
   CRITICAL: "find" must be the exact text currently in the clause. It must NEVER be empty.
   For insertions into an existing clause: set "find" to the sentence before the insertion
   point and "replace" to that sentence plus the new text appended.

3. Insert a brand-new clause into the document:
   {"action":"insert","tag":"<new clause name>","text":"<full text of the new clause>"}

4. Delete an existing clause:
   {"action":"delete","tag":"<exact tag name>"}

Only emit one action JSON block per reply. If the user asks for multiple changes,
handle them one at a time and confirm before proceeding to the next.`;

async function main() {
  console.log("Creating DocSense AI agent...");
  console.log(`Endpoint: ${baseEndpoint}`);

  try {
    const agent = await client.beta.assistants.create({
      model:       "gpt-4o",
      name:        "DocSense AI",
      description: "Contract intelligence assistant for Conga CLM contracts",
      instructions: INSTRUCTIONS,
      temperature: 0.2,
      top_p:       1.0,
    });

    console.log("\n✅ Agent created successfully!");
    console.log(`   Name  : ${agent.name}`);
    console.log(`   Model : ${agent.model}`);
    console.log(`   ID    : ${agent.id}`);
    console.log("\nAdd this to mock-api/.env:");
    console.log(`   AZURE_OPENAI_AGENT_ID=${agent.id}`);
    console.log("\nAgent instructions stored in Azure — not billed per call.");
  } catch (err) {
    console.error("Failed to create agent:", err.message);
    if (err.status) console.error("HTTP status:", err.status);
    process.exit(1);
  }
}

main();
