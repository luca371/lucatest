import { useState, useRef, useEffect } from "react";
import mammoth from "mammoth";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { saveAs } from "file-saver";
import PptxGenJS from "pptxgenjs";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";
import "./App.css";

// ── text extraction ──────────────────────────────────────────────────────────
const extractText = async (file) => {
  if (file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
    const buf = await file.arrayBuffer();
    const { value } = await mammoth.extractRawText({ arrayBuffer: buf });
    return value.trim();
  }
  if (file.type.startsWith("text/")) return await file.text();
  if (file.type === "application/pdf") return "[PDF uploaded]";
  return `[File: ${file.name}]`;
};

// ── quick actions ────────────────────────────────────────────────────────────
const QUICK_ACTIONS = [
  {
    id: "risk",
    icon: "§",
    label: "Risk Scanner",
    display: "Scan documents for legal risks",
    color: "#e05c5c",
    prompt: "Perform a comprehensive legal risk analysis of the uploaded documents. Structure your response with:\n\n**EXECUTIVE SUMMARY**\n\n**HIGH RISK**\n\n**MEDIUM RISK**\n\n**LOW RISK**\n\n**RECOMMENDED ACTIONS**\n\nBe specific. Cite clause numbers or sections where applicable.",
    mode: "risk",
  },
  {
    id: "timeline",
    icon: "—",
    label: "Legal Timeline",
    display: "Generate a visual timeline of key dates",
    color: "#5c9ee0",
    prompt: "Extract all dates, deadlines, obligations, and key events from the uploaded documents. Generate an HTML visual timeline. Return ONLY a complete HTML block starting with <div — no explanation, no markdown fences. Use inline styles only. Dark theme: background #0e0e12, text #e8e8f0, accent #c9a84c.",
    mode: "visual",
  },
  {
    id: "slides",
    icon: "▦",
    label: "Slide Deck",
    display: "Generate a slide deck from documents",
    color: "#7c6af7",
    prompt: "Summarize the uploaded documents as a professional slide deck for a client presentation. Return ONLY valid JSON, nothing else, no markdown fences: {\"slides\":[{\"title\":\"string\",\"bullets\":[\"string\",\"string\"]}]}",
    mode: "slides",
  },
];

// ── helpers ──────────────────────────────────────────────────────────────────
const detectMode = (text) => {
  const t = text.toLowerCase();
  if (t.includes("slide") || t.includes("presentation") || t.includes("pptx")) return "slides";
  if (t.includes("timeline") || t.includes("visual") || t.includes("diagram")) return "visual";
  if (t.includes("risk") || t.includes("scan")) return "risk";
  return "legal";
};

const extractHtml = (text) => {
  const fence = text.match(/```html\n?([\s\S]*?)```/i);
  if (fence) return fence[1].trim();
  if (text.trim().startsWith("<")) return text.trim();
  return null;
};

const extractJson = (text) => {
  try {
    const fence = text.match(/```json\n?([\s\S]*?)```/i);
    return JSON.parse(fence ? fence[1] : text.trim());
  } catch {
    const match = text.match(/\{[\s\S]*\}/);
    if (match) try { return JSON.parse(match[0]); } catch {}
    return null;
  }
};

// ── pptx export ──────────────────────────────────────────────────────────────
const exportPptx = (slides) => {
  const pptx = new PptxGenJS();
  slides.forEach((slide, i) => {
    const s = pptx.addSlide();
    s.background = { color: "0E0E12" };
    s.addText(slide.title, { x: 0.5, y: 0.3, w: 9, h: 0.9, fontSize: 26, bold: true, color: "C9A84C", fontFace: "Georgia" });
    (slide.bullets || []).forEach((b, j) => {
      s.addText(`${b}`, { x: 0.7, y: 1.6 + j * 0.65, w: 8.5, h: 0.6, fontSize: 15, color: "E8E8F0", fontFace: "Calibri", bullet: true });
    });
    s.addText(`${i + 1}`, { x: 9, y: 5.3, w: 0.5, h: 0.3, fontSize: 10, color: "5A5A72", align: "right" });
  });
  pptx.writeFile({ fileName: "LegalDeck.pptx" });
};

// ── export word ───────────────────────────────────────────────────────────────
const exportDocx = async (content, index) => {
  const lines = content.split("\n").filter(Boolean);
  const children = lines.map((line) => {
    if (line.startsWith("# ")) return new Paragraph({ text: line.slice(2), heading: HeadingLevel.HEADING_1 });
    if (line.startsWith("## ")) return new Paragraph({ text: line.slice(3), heading: HeadingLevel.HEADING_2 });
    if (line.startsWith("### ")) return new Paragraph({ text: line.slice(4), heading: HeadingLevel.HEADING_3 });
    return new Paragraph({ children: [new TextRun({ text: line.replace(/\*\*/g, "") })] });
  });
  const doc = new Document({ sections: [{ children }] });
  const blob = await Packer.toBlob(doc);
  saveAs(blob, `LegalAnalysis_${index + 1}.docx`);
};

// ── sub-components ────────────────────────────────────────────────────────────
const SlidePreview = ({ slides, onExport }) => (
  <div className="slides-preview">
    <div className="slides-grid">
      {slides.map((slide, i) => (
        <div key={i} className="slide-card">
          <div className="slide-num">SLIDE {i + 1}</div>
          <div className="slide-title">{slide.title}</div>
          <ul className="slide-bullets">
            {(slide.bullets || []).map((b, j) => <li key={j}>{b}</li>)}
          </ul>
        </div>
      ))}
    </div>
    <button className="export-btn pptx-btn" onClick={onExport}>Download PPTX</button>
  </div>
);

const VisualFrame = ({ html }) => (
  <div className="visual-frame">
    <iframe
      srcDoc={`<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{background:#0e0e12;color:#e8e8f0;font-family:sans-serif;padding:1rem;margin:0;}</style></head><body>${html}</body></html>`}
      sandbox="allow-scripts"
      title="Visual"
      className="visual-iframe"
    />
  </div>
);

// ── main ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [docs, setDocs] = useState([]);
  const [question, setQuestion] = useState("");
  const [messages, setMessages] = useState([]);
  const [loading, setLoading] = useState(false);
  const [loadingFiles, setLoadingFiles] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const chatRef = useRef(null);
  const textareaRef = useRef(null);

  useEffect(() => {
    chatRef.current?.scrollTo({ top: 99999, behavior: "smooth" });
  }, [messages]);

  useEffect(() => {
    if (!textareaRef.current) return;
    textareaRef.current.style.height = "auto";
    textareaRef.current.style.height = Math.min(textareaRef.current.scrollHeight, 160) + "px";
  }, [question]);

  const handleFiles = async (e) => {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    setLoadingFiles(true);
    try {
      const newDocs = await Promise.all(
        files.map(async (file) => {
          let text = "";
          try { text = await extractText(file); } catch { text = "[Error reading file]"; }
          return { id: crypto.randomUUID(), name: file.name, text };
        })
      );
      setDocs((prev) => [...prev, ...newDocs]);
    } finally {
      setLoadingFiles(false);
      e.target.value = "";
    }
  };

  const removeDoc = (id) => setDocs((prev) => prev.filter((d) => d.id !== id));

  const sendMessage = async (promptText, mode, displayText) => {
    const modeToUse = mode || detectMode(promptText);
    // displayText is what shows in chat; promptText is what goes to the API
    const visibleText = displayText || promptText;
    const userMsg = { role: "user", content: promptText, display: visibleText };
    const history = [...messages, userMsg];
    setMessages([...history, { role: "assistant", content: "...", mode: modeToUse }]);
    setQuestion("");
    setLoading(true);

    const docsContext = docs.length
      ? docs.map((d) => `=== ${d.name} ===\n${d.text}`).join("\n\n---\n\n")
      : "";

    try {
      const res = await fetch("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          messages: history.map((m) => ({ role: m.role, content: m.content })),
          docsContext,
          mode: modeToUse,
        }),
      });

      const text = await res.text();
      let reply;
      try { const data = JSON.parse(text); reply = data.reply || data.error || text; }
      catch { reply = text; }

      setMessages((prev) => {
        const filtered = prev.filter((m) => !(m.role === "assistant" && m.content === "..."));
        return [...filtered, { role: "assistant", content: reply, mode: modeToUse }];
      });
    } catch (err) {
      setMessages((prev) => {
        const filtered = prev.filter((m) => !(m.role === "assistant" && m.content === "..."));
        return [...filtered, { role: "assistant", content: `Error: ${err.message}`, mode: "legal" }];
      });
    } finally {
      setLoading(false);
    }
  };

  const renderMessage = (m, i) => {
    // ── user message ──
    if (m.role === "user") {
      return (
        <div key={i} className="msg user">
          <div className="msg-role">You</div>
          <div className="msg-content user-content">{m.display || m.content}</div>
        </div>
      );
    }

    // ── assistant loading ──
    if (m.content === "...") {
      return (
        <div key={i} className="msg assistant">
          <div className="msg-role">LexAI</div>
          <div className="msg-content ai-content"><span className="dots" /></div>
        </div>
      );
    }

    // ── slides ──
    if (m.mode === "slides") {
      const json = extractJson(m.content);
      return (
        <div key={i} className="msg assistant">
          <div className="msg-role">LexAI</div>
          <div className="msg-content ai-content">
            {json?.slides
              ? <SlidePreview slides={json.slides} onExport={() => exportPptx(json.slides)} />
              : <ReactMarkdown remarkPlugins={[remarkGfm]} className="markdown">{m.content}</ReactMarkdown>
            }
          </div>
        </div>
      );
    }

    // ── visual ──
    if (m.mode === "visual") {
      const html = extractHtml(m.content);
      return (
        <div key={i} className="msg assistant">
          <div className="msg-role">LexAI</div>
          <div className="msg-content ai-content">
            {html
              ? <VisualFrame html={html} />
              : <ReactMarkdown remarkPlugins={[remarkGfm]} className="markdown">{m.content}</ReactMarkdown>
            }
          </div>
        </div>
      );
    }

    // ── default legal response ──
    return (
      <div key={i} className="msg assistant">
        <div className="msg-role">LexAI</div>
        <div className="msg-content ai-content">
          <ReactMarkdown remarkPlugins={[remarkGfm]} className="markdown">{m.content}</ReactMarkdown>
          <div className="msg-actions">
            <button className="export-btn" onClick={() => exportDocx(m.content, i)}>Export Word</button>
            <button className="export-btn copy-btn" onClick={() => navigator.clipboard.writeText(m.content)}>Copy</button>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="app">
      <header className="header">
        <button className="sidebar-toggle" onClick={() => setSidebarOpen((v) => !v)}>
          <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
            <rect y="2" width="16" height="1.5" rx="1"/>
            <rect y="7.25" width="16" height="1.5" rx="1"/>
            <rect y="12.5" width="16" height="1.5" rx="1"/>
          </svg>
        </button>
        <div className="header-logo">LEX<span>AI</span></div>
        <div className="header-tag">Senior Legal Intelligence</div>
        <button className="new-chat-btn" onClick={() => setMessages([])}>New Chat</button>
      </header>

      <div className="body">
        <aside className={`sidebar ${sidebarOpen ? "open" : "closed"}`}>
          <div className="section-label">Documents</div>
          <label className={`upload-btn ${loadingFiles ? "loading" : ""}`}>
            {loadingFiles ? "Processing…" : "Upload Files"}
            <input type="file" accept=".txt,.docx,.pdf,image/*" multiple onChange={handleFiles} style={{ display: "none" }} />
          </label>
          <div className="doc-list">
            {docs.length === 0 && <p className="doc-empty">No files uploaded.</p>}
            {docs.map((doc) => (
              <div key={doc.id} className="doc-item">
                <svg className="doc-icon-svg" viewBox="0 0 16 16" fill="currentColor">
                  <path d="M4 0h6l4 4v11a1 1 0 01-1 1H3a1 1 0 01-1-1V1a1 1 0 011-1z" opacity=".15"/>
                  <path d="M4 0h6l4 4v11a1 1 0 01-1 1H3a1 1 0 01-1-1V1a1 1 0 011-1zm0 1v13h9V5H9V1H4zm6 0v3h3l-3-3z"/>
                </svg>
                <span className="doc-name" title={doc.name}>{doc.name}</span>
                <button className="doc-remove" onClick={() => removeDoc(doc.id)}>
                  <svg width="10" height="10" viewBox="0 0 10 10" fill="currentColor">
                    <path d="M1 1l8 8M9 1l-8 8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
                  </svg>
                </button>
              </div>
            ))}
          </div>

          <div className="section-label" style={{ marginTop: "1.5rem" }}>Quick Actions</div>
          <div className="quick-actions">
            {QUICK_ACTIONS.map((a) => (
              <button
                key={a.id}
                className="action-btn"
                style={{ "--action-color": a.color }}
                onClick={() => sendMessage(a.prompt, a.mode, a.display)}
                disabled={loading || docs.length === 0}
                title={docs.length === 0 ? "Upload a document first" : a.display}
              >
                <span className="action-icon">{a.icon}</span>
                <span>{a.label}</span>
              </button>
            ))}
          </div>

          <div className="section-label" style={{ marginTop: "1.5rem" }}>Suggested</div>
          <div className="suggestions">
            {[
              "Summarize key obligations",
              "Who are the parties?",
              "List termination clauses",
              "Extract all defined terms",
              "Flag unusual provisions",
            ].map((s) => (
              <button key={s} className="suggestion-chip"
                onClick={() => { setQuestion(s); textareaRef.current?.focus(); }}>
                {s}
              </button>
            ))}
          </div>
        </aside>

        <main className="chat">
          <div className="chat-history" ref={chatRef}>
            {messages.length === 0 ? (
              <div className="chat-empty">
                <div className="empty-icon-wrap">
                  <svg width="32" height="32" viewBox="0 0 32 32" fill="none" stroke="currentColor" strokeWidth="1.5">
                    <circle cx="16" cy="16" r="14"/>
                    <path d="M10 16h12M16 10v12" strokeLinecap="round"/>
                  </svg>
                </div>
                <div className="empty-title">LexAI Legal Assistant</div>
                <p className="empty-sub">Upload documents and ask anything, or use Quick Actions to scan for risks, generate timelines, and create slide decks.</p>
                <div className="empty-chips">
                  {["Review this contract", "What are my risks?", "Summarize key terms"].map((c) => (
                    <button key={c} className="empty-chip"
                      onClick={() => { setQuestion(c); textareaRef.current?.focus(); }}>{c}</button>
                  ))}
                </div>
              </div>
            ) : (
              messages.map((m, i) => renderMessage(m, i))
            )}
          </div>

          <div className="input-bar">
            <textarea
              ref={textareaRef}
              value={question}
              onChange={(e) => setQuestion(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter" && !e.shiftKey) {
                  e.preventDefault();
                  if (question.trim()) sendMessage(question);
                }
              }}
              placeholder="Ask LexAI anything about your documents…"
              disabled={loading}
              rows={1}
            />
            <button
              className="send-btn"
              onClick={() => question.trim() && sendMessage(question)}
              disabled={loading || !question.trim()}
            >
              {loading
                ? <span className="send-dots" />
                : (
                  <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
                    <path d="M2 8l12-6-5 6 5 6z"/>
                  </svg>
                )
              }
            </button>
          </div>
        </main>
      </div>
    </div>
  );
}