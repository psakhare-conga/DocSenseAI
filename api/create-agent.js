/**
 * One-time script: creates the DocSense AI agent in Azure AI Foundry
 * and prints the AGENT_ID to add to .env
 *
 * Run: node create-agent.js
 */
require("dotenv").config();
const fetch = require("node-fetch");

const PROJECT_ENDPOINT = process.env.AZURE_PROJECT_ENDPOINT || "";
const API_KEY          = process.env.AZURE_OPENAI_API_KEY   || "";
const API_VERSION      = "2025-05-15-preview";

if (!PROJECT_ENDPOINT || !API_KEY) {
  console.error("Missing AZURE_PROJECT_ENDPOINT or AZURE_OPENAI_API_KEY in .env");
  process.exit(1);
}

// ── Static instructions stored permanently in the agent ─────────────────────
// Everything here is NOT billed as input tokens per call.
// Only dynamic document context (clause list) is sent per request.
const INSTRUCTIONS = `You are Conga AI Assistance, a contract intelligence assistant for Conga CLM contracts.

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
  console.log("Creating Conga AI Assistance agent...");
  console.log(`Project endpoint: ${PROJECT_ENDPOINT}`);

  // Use direct REST call — the AzureOpenAI SDK adds /openai/ prefix which
  // is wrong for the Foundry project-scoped endpoint (/api/projects/...)
  const url = `${PROJECT_ENDPOINT.replace(/\/+$/, "")}/assistants?api-version=${API_VERSION}`;

  try {
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": API_KEY
      },
      body: JSON.stringify({
        model:        "gpt-4o",
        name:         "Conga AI Assistance",
        description:  "Contract intelligence assistant for Conga CLM contracts",
        instructions: INSTRUCTIONS,
        temperature:  0.2,
        top_p:        1.0
      })
    });

    const body = await res.json();
    if (!res.ok) {
      console.error("Failed to create agent:", JSON.stringify(body, null, 2));
      process.exit(1);
    }

    const agent = body;
    console.log("\n✅ Agent created successfully!");
    console.log(`   Name  : ${agent.name}`);
    console.log(`   Model : ${agent.model}`);
    console.log(`   ID    : ${agent.id}`);
    console.log("\nAdd this to api/.env and mock-api/.env:");
    console.log(`   AZURE_AGENT_ID=${agent.id}`);
    console.log("\nAgent instructions stored in Azure Foundry — visible at ai.azure.com → Agents");
  } catch (err) {
    console.error("Failed to create agent:", err.message);
    process.exit(1);
  }
}

main();
