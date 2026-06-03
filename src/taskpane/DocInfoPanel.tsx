/* global Office, Word */
import * as React from "react";
import { useState, useCallback, useEffect } from "react";

// ─── Types ───────────────────────────────────────────────────────────────────
interface MetaProp {
  name: string;
  value: string;
}

interface DocInfoNode {
  id: string;         // signed ID (from Word CC)
  unsignedId: string; // unsigned 32-bit equivalent
  alias: string;      // w:alias / title = "Clause" | "Field" | "Repeat" | "Segment" | ""
  tag: string;        // w:tag
  text: string;       // first 200 chars
  props: MetaProp[];  // all Conga metadata properties
}

interface DocInfoPanelProps {
  onNavigate: (ccId: number) => void;
}

interface Summary {
  clauses: number;
  fields: number;
  repeats: number;
  segments: number;
}

// ─── Browser-native GZIP decompression (no extra dependency) ─────────────────
async function gunzip(base64: string): Promise<string> {
  const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
  // DecompressionStream is available in Edge/Chrome (Chromium-based Office Add-in host)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ds = new (window as any).DecompressionStream("gzip");
  const writer = ds.writable.getWriter();
  const reader = ds.readable.getReader();
  writer.write(bytes);
  writer.close();
  const chunks: Uint8Array[] = [];
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
  }
  const total = chunks.reduce((a: number, c: Uint8Array) => a + c.length, 0);
  const buf = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) { buf.set(c, off); off += c.length; }
  return new TextDecoder().decode(buf);
}

// Conga stores IDs both signed and unsigned — convert to match either
function toUnsigned(id: string): string {
  const n = parseInt(id, 10);
  return isNaN(n) ? id : (n >>> 0).toString();
}

// ─── Parse the Conga custom XML part ─────────────────────────────────────────
async function parseCustomXml(
  xml: string,
  metaMap: Map<string, MetaProp[]>
): Promise<{ docProps: Record<string, string> | null; objectTypeKey: string }> {
  let docProps: Record<string, string> | null = null;
  let objectTypeKey = "";
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");

  // Case 1: <Node p2:id="ccId">base64gzip</Node> — per-control metadata
  const nodes = doc.getElementsByTagName("Node");
  for (let i = 0; i < nodes.length; i++) {
    const node = nodes[i];
    const id = node.getAttribute("p2:id") || node.getAttribute("id") || "";
    const base64 = node.textContent?.trim() || "";
    if (!id || !base64) continue;
    try {
      const decompressed = await gunzip(base64);
      const metaDoc = parser.parseFromString(decompressed, "application/xml");
      const metaNode = metaDoc.getElementsByTagName("Metadata")[0];
      if (!metaNode) continue;
      const props: MetaProp[] = [];
      for (let j = 0; j < metaNode.children.length; j++) {
        const child = metaNode.children[j];
        const val = child.textContent || "";
        if (val) props.push({ name: child.nodeName, value: val });
      }
      if (props.length) {
        metaMap.set(id, props);
        metaMap.set(toUnsigned(id), props);
      }
    } catch {
      // malformed node — skip
    }
  }

  // Case 2: <Properties>base64gzip</Properties> — document-level properties
  const propsEl = doc.getElementsByTagName("Properties")[0];
  if (propsEl) {
    const base64 = propsEl.textContent?.trim() || "";
    if (base64) {
      try {
        const decompressed = await gunzip(base64);
        const propsDoc = parser.parseFromString(decompressed, "application/xml");
        const root = propsDoc.documentElement;
        const props: Record<string, string> = {};
        for (let j = 0; j < root.children.length; j++) {
          const child = root.children[j];
          const val = child.textContent || "";
          if (val) props[child.nodeName] = val;
        }
        if (Object.keys(props).length) {
          docProps = props;
          objectTypeKey = props["SF_OBJECT_TYPE_KEY"] || props["SFObjectType"] || "";
        }
      } catch {
        // malformed properties — skip
      }
    }
  }

  return { docProps, objectTypeKey };
}

// ─── Type badge ───────────────────────────────────────────────────────────────
function TypeBadge({ alias }: { alias: string }) {
  const lower = alias.toLowerCase();
  const labels: Record<string, string> = {
    clause: "Clause",
    field: "Field",
    repeat: "Repeat",
    segment: "Segment",
  };
  const label = labels[lower] || alias || "CC";
  return <span className={`di-type-badge di-type-${lower || "other"}`}>{label}</span>;
}

// ─── Main Component ───────────────────────────────────────────────────────────
export function DocInfoPanel({ onNavigate }: DocInfoPanelProps) {
  const [loading, setLoading] = useState(false);
  const [nodes, setNodes] = useState<DocInfoNode[]>([]);
  const [docProps, setDocProps] = useState<Record<string, string> | null>(null);
  const [objectTypeKey, setObjectTypeKey] = useState("");
  const [expandedIds, setExpandedIds] = useState<Set<string>>(new Set());
  const [summary, setSummary] = useState<Summary>({ clauses: 0, fields: 0, repeats: 0, segments: 0 });
  const [filter, setFilter] = useState<"all" | "clause" | "field" | "repeat" | "segment">("all");
  const [searchText, setSearchText] = useState("");
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      // ── Step 1: Read Conga custom XML metadata ──────────────────────────────
      const metaMap = new Map<string, MetaProp[]>();
      let docPropsResult: Record<string, string> | null = null;
      let objTypeKey = "";

      await new Promise<void>((resolve) => {
        Office.context.document.customXmlParts.getByNamespaceAsync(
          "http://www.apttus.com/schemas",
          (result: Office.AsyncResult<Office.CustomXmlPart[]>) => {
            if (
              result.status === Office.AsyncResultStatus.Succeeded &&
              result.value.length > 0
            ) {
              const part = result.value[0];
              part.getXmlAsync((xmlResult: Office.AsyncResult<string>) => {
                if (xmlResult.status === Office.AsyncResultStatus.Succeeded) {
                  parseCustomXml(xmlResult.value, metaMap)
                    .then(({ docProps: dp, objectTypeKey: otk }) => {
                      docPropsResult = dp;
                      objTypeKey = otk;
                      resolve();
                    })
                    .catch(() => resolve());
                } else {
                  resolve();
                }
              });
            } else {
              resolve();
            }
          }
        );
      });

      // ── Step 2: Read all content controls via Word.run ──────────────────────
      await Word.run(async (context) => {
        const ccs = context.document.contentControls;
        ccs.load(
          "items/id,items/tag,items/title,items/text,items/cannotEdit,items/cannotDelete,items/placeholderText"
        );
        await context.sync();

        const CONGA_TYPES = new Set(["clause", "field", "repeat", "segment"]);
        let clauses = 0, fields = 0, repeats = 0, segments = 0;

        const allNodes: DocInfoNode[] = [];

        for (const cc of ccs.items) {
          const alias = cc.title || "";
          const tag = cc.tag || "";
          const aliasLow = alias.toLowerCase();

          // Only include Conga controls (have a recognised alias or a tag)
          if (!CONGA_TYPES.has(aliasLow) && !tag) continue;

          if (aliasLow === "clause") clauses++;
          else if (aliasLow === "field") fields++;
          else if (aliasLow === "repeat") repeats++;
          else if (aliasLow === "segment") segments++;

          const idStr = cc.id.toString();
          const unsignedId = toUnsigned(idStr);

          // Conga metadata from custom XML (prefer unsigned match) — resolve FIRST
          // so we can use the metadata Tag (the real name) for display
          const xmlProps = metaMap.get(unsignedId) || metaMap.get(idStr) || [];

          // The Conga XML "Tag" field contains the real clause/field name.
          // Office.js cc.tag (w:tag) often stores a numeric ID, not the name.
          const metaTagName = xmlProps.find((p) => p.name === "Tag")?.value || "";
          const displayTag = metaTagName || tag; // prefer Conga name over raw w:tag

          // Base Office.js properties
          // We deduplicate against xmlProps so the Conga metadata values always win.
          const xmlPropNames = new Set(xmlProps.map((p) => p.name));
          const baseProps: MetaProp[] = [
            { name: "ID (Unsigned)", value: unsignedId },
            { name: "ID (Signed)", value: idStr },
            { name: "Alias / Type", value: alias },
            // Only keep the raw w:tag if it differs from the Conga Tag (i.e. it's informative)
            ...(tag && tag !== metaTagName ? [{ name: "Word Tag (w:tag)", value: tag }] : []),
            { name: "CannotEdit", value: String(cc.cannotEdit) },
            { name: "CannotDelete", value: String(cc.cannotDelete) },
            { name: "PlaceholderText", value: cc.placeholderText || "" },
          ].filter(
            (p) =>
              p.value !== "" &&
              p.value !== "false" &&
              p.value !== "undefined" &&
              !xmlPropNames.has(p.name) // xmlProps takes precedence — no duplicates
          );

          allNodes.push({
            id: idStr,
            unsignedId,
            alias,
            tag: displayTag, // display the real name, not the raw w:tag ID
            text: cc.text ? cc.text.substring(0, 250) : "",
            props: [...baseProps, ...xmlProps],
          });
        }

        setSummary({ clauses, fields, repeats, segments });
        setNodes(allNodes);
        setDocProps(docPropsResult);
        setObjectTypeKey(objTypeKey);
      });
    } catch (err) {
      console.error("DocInfo load error:", err);
      setError(String(err));
    } finally {
      setLoading(false);
    }
  }, []);

  // Load on first render
  useEffect(() => { load(); }, [load]);

  const toggleExpand = (id: string) => {
    setExpandedIds((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  };

  const filteredNodes = nodes.filter((n) => {
    const aliasMatch = filter === "all" || n.alias.toLowerCase() === filter;
    const textMatch =
      !searchText ||
      n.tag.toLowerCase().includes(searchText.toLowerCase()) ||
      n.alias.toLowerCase().includes(searchText.toLowerCase()) ||
      n.unsignedId.includes(searchText);
    return aliasMatch && textMatch;
  });

  return (
    <div className="di-panel">
      {/* ── Summary badges ── */}
      <div className="di-summary">
        {objectTypeKey && <div className="di-obj-type">{objectTypeKey}</div>}
        <div className="di-badges">
          <span
            className={`di-badge di-badge-clause ${filter === "clause" ? "di-badge-active" : ""}`}
            onClick={() => setFilter(filter === "clause" ? "all" : "clause")}
            title="Filter Clauses"
          >
            📄 {summary.clauses}
          </span>
          <span
            className={`di-badge di-badge-field ${filter === "field" ? "di-badge-active" : ""}`}
            onClick={() => setFilter(filter === "field" ? "all" : "field")}
            title="Filter Fields"
          >
            🔤 {summary.fields}
          </span>
          <span
            className={`di-badge di-badge-repeat ${filter === "repeat" ? "di-badge-active" : ""}`}
            onClick={() => setFilter(filter === "repeat" ? "all" : "repeat")}
            title="Filter Tables/Repeats"
          >
            🔁 {summary.repeats}
          </span>
          <span
            className={`di-badge di-badge-segment ${filter === "segment" ? "di-badge-active" : ""}`}
            onClick={() => setFilter(filter === "segment" ? "all" : "segment")}
            title="Filter Segments"
          >
            📦 {summary.segments}
          </span>
          <button className="di-refresh-btn" onClick={load} title="Reload metadata">↻</button>
        </div>
      </div>

      {/* ── Search ── */}
      <div className="di-search-row">
        <input
          className="di-search"
          placeholder="Search by tag, type or ID…"
          value={searchText}
          onChange={(e) => setSearchText(e.target.value)}
        />
        {searchText && (
          <button className="di-search-clear" onClick={() => setSearchText("")}>✕</button>
        )}
      </div>

      {/* ── Loading ── */}
      {loading && (
        <div className="di-loading">
          <div className="typing"><span /><span /><span /></div>
          <span>Reading document metadata…</span>
        </div>
      )}

      {/* ── Error ── */}
      {error && !loading && (
        <div className="di-error">⚠️ {error}</div>
      )}

      {/* ── Empty ── */}
      {!loading && !error && nodes.length === 0 && (
        <div className="di-empty">
          No Conga content controls found.<br />
          Open a Conga CLM contract document.
        </div>
      )}

      {/* ── Document Properties ── */}
      {!loading && docProps && (
        <div className="di-section">
          <div
            className="di-section-header"
            onClick={() => toggleExpand("__docprops__")}
          >
            <span className="di-chevron">{expandedIds.has("__docprops__") ? "▾" : "▸"}</span>
            <span className="di-section-title">📋 Document Properties</span>
          </div>
          {expandedIds.has("__docprops__") && (
            <div className="di-props-table-wrap">
              <table className="di-props-table">
                <tbody>
                  {Object.entries(docProps).map(([k, v]) => (
                    <tr key={k}>
                      <td className="di-prop-name">{k}</td>
                      <td className="di-prop-value">{v}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* ── Controls list ── */}
      {!loading && filteredNodes.length > 0 && (
        <div className="di-controls-list">
          {filteredNodes.length !== nodes.length && (
            <div className="di-filter-info">
              Showing {filteredNodes.length} of {nodes.length} controls
              <button className="di-filter-clear" onClick={() => { setFilter("all"); setSearchText(""); }}>
                Clear filter
              </button>
            </div>
          )}
          {filteredNodes.map((node) => {
            const isExpanded = expandedIds.has(node.id);
            return (
              <div key={node.id} className={`di-control-item ${isExpanded ? "di-control-expanded" : ""}`}>
                {/* Header row */}
                <div
                  className="di-control-header"
                  onClick={() => toggleExpand(node.id)}
                >
                  <span className="di-chevron">{isExpanded ? "▾" : "▸"}</span>
                  <TypeBadge alias={node.alias} />
                  <span className="di-control-tag" title={node.tag}>
                    {node.tag || "(no tag)"}
                  </span>
                  <span className="di-control-id">#{node.unsignedId}</span>
                  <button
                    className="di-nav-btn"
                    onClick={(e) => {
                      e.stopPropagation();
                      onNavigate(parseInt(node.id, 10));
                    }}
                    title="Navigate to this control in Word"
                  >
                    📍
                  </button>
                </div>

                {/* Expanded detail */}
                {isExpanded && (
                  <div className="di-control-detail">
                    {/* Text preview */}
                    {node.text && (
                      <div className="di-text-preview">{node.text}</div>
                    )}
                    {/* Metadata table */}
                    <table className="di-props-table">
                      <tbody>
                        {node.props.map((p) => (
                          <tr key={p.name}>
                            <td className="di-prop-name">{p.name}</td>
                            <td className="di-prop-value">
                              {p.value || <span className="di-prop-empty">—</span>}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
