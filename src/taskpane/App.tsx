/* global Office, Word */
import * as React from "react";
import { useState, useEffect, useRef } from "react";
import "./taskpane.css";
import { DocInfoPanel } from "./DocInfoPanel";

declare const process: { env: { API_URL: string } };
const API_URL = process.env.API_URL || "http://localhost:5000";

// ─── Types ──────────────────────────────────────────────────────────────────

interface ContentControl {
  id: number;
  tag: string;
  title: string;
  text: string;
  // Richer properties loaded from Office.js (mirrors XAuthor ContentControlInfo)
  type?: string;           // Conga type: "Clause" | "Field" | "Repeat" | "Segment"
  cannotEdit?: boolean;    // cc is read-only (locked clause)
  cannotDelete?: boolean;  // cc is protected from deletion
  placeholderText?: string;// field label / placeholder
  // Properties from Conga custom XML metadata (populated by /api/metadata)
  subType?: string;        // e.g. "HeaderValue", "Text", etc.
  sfSource?: string;       // Salesforce source object API path
  sfId?: string;           // Salesforce record ID
  sfObjectName?: string;   // e.g. "Apttus__Agreement__c"
  smart?: string;          // "true" | "false"
  dirty?: string;          // "true" = user-modified
  readonly?: string;       // "true" = read-only from metadata
  markedForDeletion?: string; // "true" = pending removal
  action?: string;         // "Inserted" | "Deleted" | "Replaced"
  clauseMarkup?: string;   // clause OOXML markup
}

interface Message {
  id: string;
  role: "user" | "assistant" | "system";
  text: string;
  contentControlId?: number;
}

interface UpdateAction {
  tag: string;
  find: string;
  replace: string;
}

interface InsertAction {
  tag: string;   // clause name/tag for the new clause
  text: string;  // full text of the new clause
}

interface DeleteAction {
  tag: string;
  contentControlId?: number;
}

// ─── Component ──────────────────────────────────────────────────────────────

export default function App() {
  const [activeTab, setActiveTab] = useState<"chat" | "docinfo">("chat");

  const [messages, setMessages] = useState<Message[]>([
    {
      id: "welcome",
      role: "assistant",
      text: "Hi! I'm DocSense AI. I can help you navigate and understand your contract.\n\nTry asking:\n• \"List all clauses\"\n• \"Find the payment clause\"\n• \"Summarize the termination clause\"\n• \"What are the obligations?\"\n• \"End Review\""
    }
  ]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [contentControls, setContentControls] = useState<ContentControl[]>([]);
  const [officeReady, setOfficeReady] = useState(false);
  const [showClauses, setShowClauses] = useState(false);
  const [threadId, setThreadId] = useState<string | null>(null);

  const messagesEndRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // ── Initialise Office.js ──────────────────────────────────────────────────

  useEffect(() => {
    Office.onReady(() => {
      setOfficeReady(true);
      loadDocumentMetadata();
    });
  }, []);

  // ── Auto-scroll to latest message ─────────────────────────────────────────

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  // ── Read content controls from the open Word document ────────────────────

  const loadDocumentMetadata = async () => {
    try {
      await Word.run(async (context) => {
        const ccs = context.document.contentControls;
        // Load richer properties matching XAuthor's ContentControlInfo model
        ccs.load("items/id,items/tag,items/title,items/text,items/cannotEdit,items/cannotDelete,items/placeholderText");
        await context.sync();

        const CONGA_TYPES = new Set(["clause", "field", "repeat", "segment"]);

        const loaded: ContentControl[] = ccs.items
          // Keep only Conga-tagged controls or those with a recognised Conga title
          .filter((cc) => cc.tag || CONGA_TYPES.has((cc.title || "").toLowerCase()))
          .map((cc) => ({
            id: cc.id,
            // cc.tag   = w:tag  — the clause/field name (e.g. "Payment Terms")
            // cc.title = w:alias — the Conga type  (e.g. "Clause", "Field", "Repeat")
            tag: cc.tag || "",
            title: cc.tag
              ? `${cc.tag}${cc.title ? " (" + cc.title + ")" : ""}`
              : cc.title || "Unnamed",
            text: cc.text ? cc.text.substring(0, 300) : "",
            type: cc.title || "",
            cannotEdit: cc.cannotEdit || false,
            cannotDelete: cc.cannotDelete || false,
            placeholderText: cc.placeholderText || ""
          }));

        setContentControls(loaded);

        const clauses = loaded.filter((cc) => (cc.type || "").toLowerCase() === "clause");
        const fields  = loaded.filter((cc) => (cc.type || "").toLowerCase() === "field");

        addSystemMessage(
          loaded.length > 0
            ? `Document loaded — found ${clauses.length} clause${clauses.length !== 1 ? "s" : ""} and ${fields.length} field${fields.length !== 1 ? "s" : ""}.`
            : "Document loaded — no Conga content controls found. Open a Conga contract document."
        );
      });
    } catch {
      addSystemMessage("Could not read document. Make sure a Word document is open.");
    }
  };

  // ── Navigate Word to a specific content control ──────────────────────────

  const navigateToClause = async (contentControlId: number) => {
    try {
      await Word.run(async (context) => {
        const cc = context.document.contentControls.getById(contentControlId);
        cc.select("Select");
        await context.sync();
      });
    } catch {
      addSystemMessage(`Could not navigate to content control ID ${contentControlId}.`);
    }
  };

  // ── Insert a brand-new clause at the end of the document ────────────────

  const insertClause = async (action: InsertAction) => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.insertParagraph("", "End");                              // blank line before
        const para = body.insertParagraph(action.text, "End");        // clause text
        para.styleBuiltIn = Word.BuiltInStyleName.normal;
        para.font.bold = false;
        await context.sync();
        addSystemMessage(`✅ New clause "${action.tag}" inserted at the end of the document.`);
        loadDocumentMetadata();
      });
    } catch (err) {
      addSystemMessage(`Failed to insert clause: ${err}`);
    }
  };

  // ── Update text inside a specific content control via Office.js ─────────

  const updateClause = async (action: UpdateAction) => {
    // Find the content control by tag
    const cc = contentControls.find(
      (c) => c.tag === action.tag ||
             c.tag.toLowerCase() === action.tag.toLowerCase() ||
             c.id === Number(action.tag)
    );
    if (!cc) {
      addSystemMessage(`Could not find clause with tag "${action.tag}" to update.`);
      return;
    }
    try {
      await Word.run(async (context) => {
        const wordCc = context.document.contentControls.getById(cc.id);
        wordCc.load("text");
        await context.sync();

        const original = wordCc.text;
        // If find is empty, treat as append to end of clause
        if (!action.find || action.find.trim() === "") {
          wordCc.insertText(" " + action.replace, "End");
          await context.sync();
          addSystemMessage(`\u2705 Inserted "${action.replace}" at the end of "${cc.tag}"`);
          loadDocumentMetadata();
          return;
        }
        const updated  = original.replace(
          new RegExp(action.find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi"),
          action.replace
        );

        if (updated === original) {
          addSystemMessage(`Could not find "${action.find}" inside the "${cc.tag}" clause.`);
          return;
        }

        wordCc.insertText(updated, "Replace");
        await context.sync();
        addSystemMessage(`✅ Updated: "${action.find}" → "${action.replace}" in "${cc.tag}"`);

        // Refresh the content controls state with new text
        loadDocumentMetadata();
      });
    } catch (err) {
      addSystemMessage(`Failed to update clause: ${err}`);
    }
  };

  const addSystemMessage = (text: string) =>
    setMessages((prev) => [
      ...prev,
      { id: `sys-${Date.now()}`, role: "system", text }
    ]);

  // ── Send a chat message to the API ────────────────────────────────────────

  const sendMessage = async (overrideText?: string) => {
    const text = (overrideText ?? input).trim();
    if (!text || loading) return;

    setInput("");
    setMessages((prev) => [
      ...prev,
      { id: `u-${Date.now()}`, role: "user", text }
    ]);
    setLoading(true);

    try {
      const response = await fetch(`${API_URL}/api/chat`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          message: text,
          contentControls,
          docText: contentControls.map(cc => `${cc.tag}: ${cc.text}`).join("\n"),
          threadId
        })
      });

      if (!response.ok) throw new Error(`API ${response.status}`);

      const data: { reply: string; contentControlId?: number; navigateTo?: string; updateAction?: UpdateAction; insertAction?: InsertAction; deleteAction?: DeleteAction; threadId?: string } = await response.json();

      // Persist thread ID for multi-turn conversation
      if (data.threadId) {
        setThreadId(data.threadId);
      }

      setMessages((prev) => [
        ...prev,
        {
          id: `a-${Date.now()}`,
          role: "assistant",
          text: data.reply || (data.updateAction ? `Updating "${data.updateAction.find}" → "${data.updateAction.replace}" in ${data.updateAction.tag}…` : data.insertAction ? `Inserting clause "${data.insertAction.tag}"…` : "Done."),
          contentControlId: data.contentControlId
        }
      ]);

      if (data.contentControlId) {
        await navigateToClause(data.contentControlId);
      }

      if (data.updateAction) {
        await updateClause(data.updateAction);
      }

      if (data.insertAction) {
        await insertClause(data.insertAction);
      }

      if (data.deleteAction) {
        const cc = contentControls.find(
          (c) => c.tag === data.deleteAction!.tag ||
                 c.tag.toLowerCase() === data.deleteAction!.tag.toLowerCase() ||
                 c.id === data.deleteAction!.contentControlId
        );
        if (cc) {
          try {
            await Word.run(async (context) => {
              const wordCc = context.document.contentControls.getById(cc.id);
              wordCc.delete(false);
              await context.sync();
              addSystemMessage(`✅ Deleted clause "${cc.tag}".`);
              loadDocumentMetadata();
            });
          } catch (err) {
            addSystemMessage(`Failed to delete clause: ${err}`);
          }
        } else {
          addSystemMessage(`Could not find clause "${data.deleteAction.tag}" to delete.`);
        }
      }
    } catch {
      setMessages((prev) => [
        ...prev,
        {
          id: `err-${Date.now()}`,
          role: "assistant",
          text: "Could not reach the API. Make sure the API server is running:\n  cd api && npm start"
        }
      ]);
    } finally {
      setLoading(false);
      inputRef.current?.focus();
    }
  };

  const quickAction = (prompt: string) => {
    setInput(prompt);
    inputRef.current?.focus();
  };

  const quickActionAndSend = (prompt: string) => {
    void sendMessage(prompt);
  };

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div className="app">

      {/* ── Header ── */}
      <div className="header">
        <div className="header-title">
          <span className="header-icon">⚡</span>
          <span>DocSense AI</span>
        </div>
        <div className="header-right">
          <button
            className="refresh-btn"
            onClick={loadDocumentMetadata}
            title="Re-read document"
          >
            ↻
          </button>
          <div
            className={`status-dot ${officeReady ? "ready" : "connecting"}`}
            title={officeReady ? "Word connected" : "Connecting to Word…"}
          />
        </div>
      </div>

      {/* ── Tab Bar ── */}
      <div className="tab-bar">
        <button
          className={`tab-btn${activeTab === "chat" ? " tab-btn-active" : ""}`}
          onClick={() => setActiveTab("chat")}
        >
          💬 Chat
        </button>
        <button
          className={`tab-btn${activeTab === "docinfo" ? " tab-btn-active" : ""}`}
          onClick={() => setActiveTab("docinfo")}
        >
          📋 Document Info
        </button>
      </div>

      {/* ── Document Info Tab ── */}
      {activeTab === "docinfo" && (
        <div className="tab-content">
          <DocInfoPanel onNavigate={navigateToClause} />
        </div>
      )}

      {/* ── Chat Tab ── */}
      {activeTab === "chat" && (
        <>

      {/* ── Clauses Panel ── */}
      {contentControls.length > 0 && (
        <div className="clauses-panel">
          <button
            className="clauses-toggle"
            onClick={() => setShowClauses((s) => !s)}
          >
            📄 {contentControls.length} Content Controls&nbsp;
            {showClauses ? "▲" : "▼"}
          </button>
          {showClauses && (
            <div className="clauses-list">
              {contentControls.map((cc) => (
                <div
                  key={cc.id}
                  className="clause-item"
                  onClick={() => navigateToClause(cc.id)}
                  title={`Click to navigate — ID: ${cc.id}`}
                >
                  <span className="clause-tag">{cc.title}</span>
                  <span className="clause-id">#{cc.id}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* ── Messages ── */}
      <div className="messages">
        {messages.map((msg) => (
          <div key={msg.id} className={`message message-${msg.role}`}>
            <div className="message-bubble">
              {msg.text.split("\n").map((line, i, arr) => (
                <React.Fragment key={i}>
                  {line}
                  {i < arr.length - 1 && <br />}
                </React.Fragment>
              ))}
            </div>
            {msg.contentControlId && (
              <button
                className="nav-btn"
                onClick={() => navigateToClause(msg.contentControlId!)}
              >
                📍 Navigate to clause
              </button>
            )}
          </div>
        ))}

        {/* Typing indicator */}
        {loading && (
          <div className="message message-assistant">
            <div className="message-bubble typing">
              <span />
              <span />
              <span />
            </div>
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      {/* ── Quick Actions ── */}
      <div className="quick-actions">
        <button onClick={() => quickAction("End review for reviewtype=Office365 reviewid= reviewerid=")}>End Review</button>
        <button onClick={() => quickAction("List all clauses in this document")}>
          List Clauses
        </button>
        <button onClick={() => quickAction("Find the payment clause")}>
          Payment
        </button>
        <button onClick={() => quickAction("Find the termination clause")}>
          Termination
        </button>
        <button onClick={() => quickAction("Summarize the obligations")}>
          Obligations
        </button>
        <button onClick={() => quickAction("Find the confidentiality clause")}>
          Confidentiality
        </button>
      </div>

      {/* ── Input ── */}
      <div className="input-area">
        <input
          ref={inputRef}
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && sendMessage()}
          placeholder="Ask about this document…"
          disabled={loading}
        />
        <button onClick={() => void sendMessage()} disabled={loading || !input.trim()}>
          {loading ? "…" : "Send"}
        </button>
      </div>

        </> /* end Chat tab */
      )}
    </div>
  );
}
