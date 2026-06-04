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
  clauseName?: string;     // human-readable name from Conga custom XML <Tag> field
}

interface Message {
  id: string;
  role: "user" | "assistant" | "system";
  text: string;
  contentControlId?: number;
  pendingEdit?: {
    action: UpdateAction;
    ccId: number;
    ccName: string;
  };
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

interface RiskResult {
  tag: string;
  id: number;
  riskLevel: "high" | "medium" | "low" | "unknown";
  reason: string;
  riskFactors: string[];
}

interface Suggestion {
  loading: boolean;
  text: string | null;  // null = not yet fetched
  visible: boolean;     // false = dismissed
}

// ─── Component ──────────────────────────────────────────────────────────────

export default function App() {
  const [activeTab, setActiveTab] = useState<"chat" | "docinfo">("chat");

  const [messages, setMessages] = useState<Message[]>([
    {
      id: "welcome",
      role: "assistant",
      text: "Hi! I'm Conga AI Assistance. I can help you navigate and understand your contract.\n\nTry asking:\n• \"List all clauses\"\n• \"Find the payment clause\"\n• \"Summarize the termination clause\"\n• \"What are the obligations?\""
    }
  ]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [contentControls, setContentControls] = useState<ContentControl[]>([]);
  const [officeReady, setOfficeReady] = useState(false);
  const [showClauses, setShowClauses] = useState(false);
  const [threadId, setThreadId] = useState<string | null>(null);
  const [isStreaming, setIsStreaming] = useState(false);
  const [riskResults, setRiskResults] = useState<RiskResult[] | null>(null);
  const [riskLoading, setRiskLoading] = useState(false);
  const [showRiskPanel, setShowRiskPanel] = useState(false);
  // suggestions keyed by clause id
  const [suggestions, setSuggestions] = useState<Record<number, Suggestion>>({});
  // tracks clause ids where fix has been applied this session
  const [fixedIds, setFixedIds] = useState<Set<number>>(new Set());

  const messagesEndRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  // Conga custom XML tag map: { [ccId]: clauseName } — persists across refreshes
  const tagMapRef = useRef<Record<string, string>>({});

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

  // ── Silently re-read content controls without posting a system message ──────
  // Used after updates/inserts to keep state fresh without spamming the chat.
  const refreshContentControls = async () => {
    try {
      await Word.run(async (context) => {
        const ccs = context.document.contentControls;
        ccs.load("items/id,items/tag,items/title,items/text,items/cannotEdit,items/cannotDelete,items/placeholderText");
        await context.sync();

        const CONGA_TYPES = new Set(["clause", "field", "repeat", "segment"]);
        const tagMap = tagMapRef.current;
        const loaded: ContentControl[] = ccs.items
          .filter((cc) => cc.tag || CONGA_TYPES.has((cc.title || "").toLowerCase()))
          .map((cc) => ({
            id: cc.id,
            tag: cc.tag || "",
            title: cc.tag
              ? `${cc.tag}${cc.title ? " (" + cc.title + ")" : ""}`
              : cc.title || "Unnamed",
            text: cc.text || "",
            type: cc.title || "",
            cannotEdit: cc.cannotEdit || false,
            cannotDelete: cc.cannotDelete || false,
            placeholderText: cc.placeholderText || "",
            clauseName: tagMap[cc.id.toString()] || tagMap[String(cc.id >>> 0)] || undefined
          }));
        setContentControls(loaded);
      });
    } catch { /* silent — non-critical */ }
  };

  const loadDocumentMetadata = async () => {
    setThreadId(null); // new document load = fresh conversation session
    let loaded: ContentControl[] = [];
    let xmlStrings: string[] = [];
    let clauseCount = 0, fieldCount = 0;

    try {
      await Word.run(async (context) => {
        const ccs = context.document.contentControls;
        // Load richer properties matching XAuthor's ContentControlInfo model
        ccs.load("items/id,items/tag,items/title,items/text,items/cannotEdit,items/cannotDelete,items/placeholderText");
        const parts = context.document.customXmlParts;
        parts.load("items/id"); // prime the collection
        await context.sync();

        const CONGA_TYPES = new Set(["clause", "field", "repeat", "segment"]);

        loaded = ccs.items
          // Keep only Conga-tagged controls or those with a recognised Conga title
          .filter((cc) => cc.tag || CONGA_TYPES.has((cc.title || "").toLowerCase()))
          .map((cc) => ({
            id: cc.id,
            // cc.tag   = w:tag  — numeric record ID in Conga CLM docs
            // cc.title = w:alias — the Conga type (e.g. "Clause", "Field", "Repeat")
            tag: cc.tag || "",
            title: cc.tag
              ? `${cc.tag}${cc.title ? " (" + cc.title + ")" : ""}`
              : cc.title || "Unnamed",
            text: cc.text || "",
            type: cc.title || "",
            cannotEdit: cc.cannotEdit || false,
            cannotDelete: cc.cannotDelete || false,
            placeholderText: cc.placeholderText || ""
          }));

        clauseCount = loaded.filter((cc) => (cc.type || "").toLowerCase() === "clause").length;
        fieldCount  = loaded.filter((cc) => (cc.type || "").toLowerCase() === "field").length;

        // Read raw custom XML parts — Conga stores GZIP-compressed metadata here
        const xmlResults = parts.items.map((p) => p.getXml());
        await context.sync();
        xmlStrings = xmlResults.map((r) => r.value).filter(Boolean);
      });
    } catch {
      addSystemMessage("Could not read document. Make sure a Word document is open.");
      return;
    }

    // Enrich content controls with human-readable clause names from Conga metadata
    if (xmlStrings.length) {
      try {
        const resp = await fetch(`${API_URL}/api/parse-xml-metadata`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ xmlStrings })
        });
        if (resp.ok) {
          const { tagMap = {} } = await resp.json();
          tagMapRef.current = tagMap; // persist for use in refreshContentControls
          loaded = loaded.map((cc) => ({
            ...cc,
            clauseName: tagMap[cc.id.toString()] || tagMap[String(cc.id >>> 0)] || undefined
          }));
        }
      } catch { /* non-fatal — display falls back to text snippet */ }
    }

    setContentControls(loaded);
    addSystemMessage(
      loaded.length > 0
        ? `Document loaded — found ${clauseCount} clause${clauseCount !== 1 ? "s" : ""} and ${fieldCount} field${fieldCount !== 1 ? "s" : ""}.`
        : "Document loaded — no Conga content controls found. Open a Conga contract document."
    );
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
        refreshContentControls();
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

        // Append-only path (no find text) — apply directly as tracked change
        if (!action.find || action.find.trim() === "") {
          let prevMode: Word.ChangeTrackingMode = Word.ChangeTrackingMode.off;
          try {
            context.document.load("changeTrackingMode");
            await context.sync();
            prevMode = context.document.changeTrackingMode as Word.ChangeTrackingMode;
          } catch { /* proceed */ }
          context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
          wordCc.insertText(" " + action.replace, "End");
          await context.sync();
          context.document.changeTrackingMode = prevMode;
          await context.sync();
          addSystemMessage(`✅ Tracked change: appended "${action.replace}" to "${cc.clauseName || cc.tag}"`);
          refreshContentControls();
          return;
        }

        // Word's search API rejects strings longer than ~255 chars.
        // If find text is too long, skip search validation — the confirmation
        // card will use a full CC replace instead of a ranged search.
        const isFullReplace = action.find.length > 200;

        if (!isFullReplace) {
          // Range search — pre-validate the text exists before showing the card
          const searchResults = wordCc.search(action.find, { matchCase: false, matchWholeWord: false });
          searchResults.load("items/text");
          await context.sync();

          if (searchResults.items.length === 0) {
            addSystemMessage(`Could not find "${action.find}" inside the "${cc.clauseName || cc.tag}" clause.`);
            return;
          }
        }
        // Text found (or full-replace path) — fall through to show confirmation card
      });

      // Show confirmation card in chat with diff preview
      setMessages(prev => [
        ...prev,
        {
          id: `pending-${Date.now()}`,
          role: "system" as const,
          text: action.find.length > 200
            ? `Proposed full rewrite of "${cc.clauseName || cc.tag}":`
            : `Proposed edit in "${cc.clauseName || cc.tag}":`,
          pendingEdit: {
            action,
            ccId: cc.id,
            ccName: cc.clauseName || cc.tag
          }
        }
      ]);
    } catch (err) {
      addSystemMessage(`Failed to find text in clause: ${err}`);
    }
  };

  // ── Confirm a pending edit — apply as a tracked change in Word ─────────────────

  const confirmEdit = async (msgId: string, pending: NonNullable<Message["pendingEdit"]>) => {
    // Remove buttons immediately so user can't double-click
    setMessages(prev => prev.map(m =>
      m.id === msgId ? { ...m, pendingEdit: undefined, text: `⏳ Applying tracked change in "${pending.ccName}"…` } : m
    ));
    try {
      await Word.run(async (context) => {
        // Read the document's current tracking mode so we can restore it after.
        // Fallback to off if inaccessible (e.g. protected doc).
        let prevMode: Word.ChangeTrackingMode = Word.ChangeTrackingMode.off;
        try {
          context.document.load("changeTrackingMode");
          await context.sync();
          prevMode = context.document.changeTrackingMode as Word.ChangeTrackingMode;
        } catch {
          // proceed with default
        }

        // Enable track changes so Word records the edit as a redline
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

        const wordCc = context.document.contentControls.getById(pending.ccId);

        if (pending.action.find.length > 200) {
          // Full rewrite — find text is too long for Word's search API.
          // Replace the entire content control content instead.
          wordCc.insertText(pending.action.replace, "Replace");
        } else {
          // Ranged replace — only the matched text is changed; surrounding
          // formatting (bold, italics, font size) is preserved.
          const searchResults = wordCc.search(pending.action.find, { matchCase: false, matchWholeWord: false });
          searchResults.load("items/text");
          await context.sync();

          if (searchResults.items.length === 0) {
            // Restore original mode before returning
            context.document.changeTrackingMode = prevMode;
            await context.sync();
            setMessages(prev => prev.map(m =>
              m.id === msgId ? { ...m, text: `❌ Could not find "${pending.action.find}" in "${pending.ccName}".` } : m
            ));
            return;
          }

          searchResults.items[0].insertText(pending.action.replace, "Replace");
        }

        await context.sync();

        // Restore original mode — if doc was already in trackAll, it stays that way
        context.document.changeTrackingMode = prevMode;
        await context.sync();
      });

      setMessages(prev => prev.map(m =>
        m.id === msgId
          ? { ...m, text: `✅ Tracked change applied in "${pending.ccName}" — accept or reject in Word’s Review ribbon.` }
          : m
      ));
      refreshContentControls();
    } catch (err) {
      setMessages(prev => prev.map(m =>
        m.id === msgId ? { ...m, text: `❌ Failed to apply change: ${err}` } : m
      ));
    }
  };

  // ── Cancel a pending edit ─────────────────────────────────────────────────

  const cancelEdit = (msgId: string, ccName: string) => {
    setMessages(prev => prev.map(m =>
      m.id === msgId ? { ...m, pendingEdit: undefined, text: `✕ Edit cancelled for "${ccName}".` } : m
    ));
  };

  // ── Clause risk scan ──────────────────────────────────────────────────────

  const scanRisk = async () => {
    if (riskLoading || !contentControls.length) return;
    setRiskLoading(true);
    setShowRiskPanel(true);
    setRiskResults(null);
    setSuggestions({}); // clear previous suggestions on new scan
    setFixedIds(new Set()); // clear fixed markers on new scan
    try {
      const response = await fetch(`${API_URL}/api/risk-scan`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ contentControls })
      });
      if (!response.ok) throw new Error(`API ${response.status}`);
      const data = await response.json();
      // The AI may return {clauses:[...]}, {results:[...]}, or a bare array
      const raw = data.clauses ?? data.results ?? data;
      const arr: RiskResult[] = Array.isArray(raw) ? raw : Object.values(raw).find(v => Array.isArray(v)) as RiskResult[] ?? [];
      setRiskResults(arr);
    } catch {
      setRiskResults([]);
      addSystemMessage("Risk scan failed. Make sure the API server is running.");
    } finally {
      setRiskLoading(false);
    }
  };

  // ── Request AI-suggested fix for a risky clause ─────────────────────────

  const requestSuggestion = async (r: RiskResult) => {
    const cc = contentControls.find(c => c.id === r.id);
    setSuggestions(prev => ({ ...prev, [r.id]: { loading: true, text: null, visible: true } }));
    try {
      const response = await fetch(`${API_URL}/api/suggest-fix`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          tag: cc?.clauseName || r.tag,
          id: r.id,
          text: cc?.text || "",
          riskLevel: r.riskLevel,
          reason: r.reason,
          riskFactors: r.riskFactors
        })
      });
      const data = await response.json();
      setSuggestions(prev => ({
        ...prev,
        [r.id]: { loading: false, text: data.suggestedText || null, visible: true }
      }));
    } catch {
      setSuggestions(prev => ({
        ...prev,
        [r.id]: { loading: false, text: null, visible: true }
      }));
    }
  };

  // ── Apply a suggested fix directly into the Word content control ──────────

  const applySuggestion = async (r: RiskResult, suggestedText: string) => {
    const cc = contentControls.find(c => c.id === r.id);
    if (!cc) return;
    try {
      await Word.run(async (context) => {
        const wordCc = context.document.contentControls.getById(cc.id);

        // Read current tracking mode so we can restore it after.
        // Use a fallback in case the property isn't accessible (e.g. protected doc).
        let prevMode: Word.ChangeTrackingMode = Word.ChangeTrackingMode.off;
        try {
          context.document.load("changeTrackingMode");
          await context.sync();
          prevMode = context.document.changeTrackingMode as Word.ChangeTrackingMode;
        } catch {
          // If we can't read the mode, assume off and proceed
        }

        // Enable track changes — the full rewrite is recorded as a redline.
        // Word preserves the content control's existing paragraph formatting
        // automatically in the tracked markup, so no manual font reapplication needed.
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        wordCc.insertText(suggestedText, "Replace");
        await context.sync();

        // Restore original mode — if doc was already in trackAll, it stays that way
        context.document.changeTrackingMode = prevMode;
        await context.sync();
      });
      setSuggestions(prev => ({ ...prev, [r.id]: { ...prev[r.id], visible: false } }));
      setFixedIds(prev => new Set(prev).add(r.id));
      refreshContentControls();
      addSystemMessage(`✅ Tracked change applied to "${cc.clauseName || cc.tag}" — accept or reject in Word's Review ribbon.`);
    } catch (err) {
      addSystemMessage(`Failed to apply fix: ${err}`);
    }
  };

  const addSystemMessage = (text: string) =>
    setMessages((prev) => [
      ...prev,
      { id: `sys-${Date.now()}`, role: "system", text }
    ]);

  // ── Send a chat message to the API (streaming) ────────────────────────────

  const sendMessage = async () => {
    const text = input.trim();
    if (!text || loading) return;

    setInput("");
    const assistantMsgId = `a-${Date.now()}`;

    setMessages((prev) => [
      ...prev,
      { id: `u-${Date.now()}`, role: "user", text }
    ]);
    setLoading(true);
    setIsStreaming(false);

    try {
      const response = await fetch(`${API_URL}/api/chat/stream`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          message: text,
          contentControls,
          docText: contentControls.map(cc => `${cc.tag}: ${cc.text}`).join("\n"),
          threadId
        })
      });

      if (!response.ok || !response.body) throw new Error(`API ${response.status}`);

      const reader  = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer       = "";
      let streamedText = "";
      let placeholderAdded = false;

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const parts = buffer.split("\n\n");
        buffer = parts.pop() ?? "";

        for (const part of parts) {
          if (!part.startsWith("data: ")) continue;
          let event: { type: string; text?: string; reply?: string; threadId?: string; contentControlId?: number; navigateTo?: string; updateAction?: UpdateAction; insertAction?: InsertAction; message?: string };
          try { event = JSON.parse(part.slice(6)); } catch { continue; }

          if (event.type === "token") {
            streamedText += event.text ?? "";
            if (!placeholderAdded) {
              // First token — add the placeholder message and start streaming
              placeholderAdded = true;
              setIsStreaming(true);
              setMessages((prev) => [
                ...prev,
                { id: assistantMsgId, role: "assistant", text: streamedText }
              ]);
            } else {
              setMessages((prev) =>
                prev.map((m) => m.id === assistantMsgId ? { ...m, text: streamedText } : m)
              );
            }

          } else if (event.type === "done") {
            const finalText = event.reply ?? streamedText;
            if (!placeholderAdded) {
              // No tokens were streamed (e.g. keyword fallback) — add message now
              setMessages((prev) => [
                ...prev,
                { id: assistantMsgId, role: "assistant", text: finalText, contentControlId: event.contentControlId }
              ]);
            } else {
              // Replace streamed text with clean reply (action JSON stripped out)
              setMessages((prev) =>
                prev.map((m) =>
                  m.id === assistantMsgId
                    ? { ...m, text: finalText, contentControlId: event.contentControlId }
                    : m
                )
              );
            }

            if (event.threadId) setThreadId(event.threadId);
            if (event.contentControlId) await navigateToClause(event.contentControlId);
            if (event.updateAction)     await updateClause(event.updateAction);
            if (event.insertAction)     await insertClause(event.insertAction);

          } else if (event.type === "error") {
            const errText = event.message ?? "An error occurred. Please try again.";
            if (!placeholderAdded) {
              setMessages((prev) => [...prev, { id: assistantMsgId, role: "assistant", text: errText }]);
            } else {
              setMessages((prev) =>
                prev.map((m) => m.id === assistantMsgId ? { ...m, text: errText } : m)
              );
            }
          }
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
      setIsStreaming(false);
      inputRef.current?.focus();
    }
  };

  const quickAction = (prompt: string) => {
    setInput(prompt);
    inputRef.current?.focus();
  };

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div className="app">

      {/* ── Header ── */}
      <div className="header">
        <div className="header-title">
          <span className="header-icon">⚡</span>
          <span>Conga AI Assistance</span>
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

      {/* ── Risk Panel ── */}
      {showRiskPanel && (
        <div className="risk-panel">
          <div className="risk-panel-header">
            <span>⚠️ Clause Risk Scan</span>
            <div className="risk-panel-header-right">
              {riskResults && Array.isArray(riskResults) && !riskLoading && (
                <span className="risk-summary">
                  {["high","medium","low"].map(level => {
                    const count = riskResults.filter(r => r.riskLevel === level).length;
                    return count > 0 ? (
                      <span key={level} className={`risk-count risk-count-${level}`}>
                        {level === "high" ? "🔴" : level === "medium" ? "🟡" : "🟢"} {count}
                      </span>
                    ) : null;
                  })}
                </span>
              )}
              <button className="risk-panel-close" onClick={() => setShowRiskPanel(false)}>✕</button>
            </div>
          </div>

          {riskLoading && (
            <div className="risk-loading">
              <div className="risk-loading-dots"><span /><span /><span /></div>
              <span>Analysing clauses…</span>
            </div>
          )}

          {!riskLoading && riskResults && Array.isArray(riskResults) && riskResults.length === 0 && (
            <div className="risk-empty">No clauses found to analyse.</div>
          )}

          {!riskLoading && riskResults && Array.isArray(riskResults) && riskResults.length > 0 && (
            <div className="risk-list">
              {["high", "medium", "low"].map(level =>
                riskResults
                  .filter(r => r.riskLevel === level)
                  .map(r => {
                    const cc = contentControls.find(c => c.id === r.id);
                    // Priority: Conga metadata Tag > placeholderText > text snippet > raw tag
                    const isNumericTag = /^\d+$/.test(cc?.tag || "");
                    const displayName =
                      cc?.clauseName
                      || (cc?.placeholderText && cc.placeholderText.trim())
                      || (!isNumericTag && cc?.tag)
                      || (cc?.text ? cc.text.trim().split(/\s+/).slice(0, 5).join(" ") + "…" : null)
                      || r.tag;
                    const isFixed = fixedIds.has(r.id);
                    return (
                    <div
                      key={r.id}
                      className={`risk-item risk-item-${r.riskLevel}${isFixed ? " risk-item-fixed" : ""}`}
                    >
                      <div className="risk-item-top">
                        <span className="risk-badge risk-badge-${r.riskLevel}">
                          {isFixed ? "✅ Fixed" : r.riskLevel === "high" ? "🔴 High" : r.riskLevel === "medium" ? "🟡 Medium" : "🟢 Low"}
                        </span>
                        <span
                          className="risk-tag"
                          onClick={() => navigateToClause(r.id)}
                          title="Click to navigate to this clause"
                          style={{ cursor: "pointer", flex: 1 }}
                        >{displayName}</span>
                        {(r.riskLevel === "high" || r.riskLevel === "medium") && !suggestions[r.id]?.visible && !isFixed && (
                          <button
                            className="suggest-fix-btn"
                            onClick={() => requestSuggestion(r)}
                            title="Get AI-suggested rewrite to reduce risk"
                          >✨ Fix</button>
                        )}
                      </div>
                      {!isFixed && <div className="risk-reason">{r.reason}</div>}
                      {!isFixed && r.riskFactors && r.riskFactors.length > 0 && (
                        <div className="risk-factors">
                          {r.riskFactors.map((f, i) => (
                            <span key={i} className="risk-factor-chip">{f}</span>
                          ))}
                        </div>
                      )}

                      {/* ── Suggestion panel ── */}
                      {suggestions[r.id]?.visible && (
                        <div className="suggestion-panel">
                          {suggestions[r.id]?.loading ? (
                            <div className="suggestion-loading">
                              <div className="risk-loading-dots"><span /><span /><span /></div>
                              <span>Generating fix…</span>
                            </div>
                          ) : suggestions[r.id]?.text ? (
                            <>
                              <div className="suggestion-label">✨ Suggested rewrite</div>
                              <div className="suggestion-text">{suggestions[r.id].text}</div>
                              <div className="suggestion-actions">
                                <button
                                  className="suggestion-apply-btn"
                                  onClick={() => applySuggestion(r, suggestions[r.id].text!)}
                                >✅ Apply</button>
                                <button
                                  className="suggestion-dismiss-btn"
                                  onClick={() => setSuggestions(prev => ({ ...prev, [r.id]: { ...prev[r.id], visible: false } }))}
                                >✕ Dismiss</button>
                              </div>
                            </>
                          ) : (
                            <div className="suggestion-error">
                              Could not generate suggestion.{" "}
                              <span
                                style={{ cursor: "pointer", textDecoration: "underline" }}
                                onClick={() => requestSuggestion(r)}
                              >Retry</span>
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                    );
                  })
              )}
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
            {/* ── Pending edit diff + confirm/cancel buttons ── */}
            {msg.pendingEdit && (
              <div className="pending-edit">
                <div className="pending-edit-diff">
                  <span className="pending-edit-old">“{msg.pendingEdit.action.find}”</span>
                  <span className="pending-edit-arrow">&#8594;</span>
                  <span className="pending-edit-new">“{msg.pendingEdit.action.replace}”</span>
                </div>
                <div className="pending-edit-actions">
                  <button
                    className="pending-confirm-btn"
                    onClick={() => confirmEdit(msg.id, msg.pendingEdit!)}
                  >✅ Apply as tracked change</button>
                  <button
                    className="pending-cancel-btn"
                    onClick={() => cancelEdit(msg.id, msg.pendingEdit!.ccName)}
                  >✕ Cancel</button>
                </div>
              </div>
            )}
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

        {/* Typing indicator — shown while connecting, hidden once tokens arrive */}
        {loading && !isStreaming && (
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
        <button
          className={`risk-scan-btn${riskLoading ? " risk-scan-btn-loading" : ""}`}
          onClick={scanRisk}
          disabled={riskLoading || !contentControls.length}
          title="Scan all clauses for risk"
        >
          {riskLoading ? "⏳ Scanning…" : "⚠️ Risk Scan"}
        </button>
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
        <button onClick={sendMessage} disabled={loading || !input.trim()}>
          {loading ? "…" : "Send"}
        </button>
      </div>

        </> /* end Chat tab */
      )}
    </div>
  );
}
