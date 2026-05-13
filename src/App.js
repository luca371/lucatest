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
  if (file.type === "application/pdf") return "[PDF uploaded — text extraction requires backend]";
  return `[File: ${file.name}]`;
};

// ── quick actions ────────────────────────────────────────────────────────────
const QUICK_ACTIONS = [
  {
    id: "risk",
    icon: "⚖",
    label: "Risk Scanner",
    color: "#e05c5c",
    prompt: "Perform a comprehensive legal risk analysis of the uploaded documents. Structure your response with:\n\n**EXECUTIVE SUMMARY**\n\n**HIGH RISK** (tag each item with 🔴)\n\n**MEDIUM RISK** (tag each item with 🟡)\n\n**LOW RISK** (tag each item with 🟢)\n\n**RECOMMENDED ACTIONS**\n\nBe specific, cite clause numbers or sections where applicable.",
    mode: "risk",
  },
  {
    id: "timeline",
    icon: "📅",
    label: "Legal Timeline",
    color: "#5c9ee0",
    prompt: "Extract all dates, deadlines, obligations, and key events from the uploaded documents. Generate an HTML visual timeline. Return ONLY a complete HTML block (starting with <div) — no explanation, no markdown fences — that displays a clean, dark-themed vertical timeline with all events sorted chronologically. Use inline styles only.",
    mode: "visual",
  },
  {
    id: "slides",
    icon: "📊",
    label: "Slide Deck",
    color: "#7c6af7",
    prompt: `Summarize the uploaded documents as a professional slide deck for a client presentation. Return ONLY valid JSON in this exact format, nothing else:\n{"slides":[{"title":"string","bullets":["string","string"]}]}`,
    mode: "slides",
  },
];

// ── helpers ──────────────────────────────────────────────────────────────────
const detectMode = (text) => {
  const t = text.toLowerCase();
  if (t.includes("slide") || t.includes("presentation") || t.includes("pptx")) return "slides";
  if (t.includes("timeline") || t.includes("visual") || t.includes("diagram") || t.includes("chart")) return "visual";
  if (t.includes("risk") || t.includes("scan") || t.includes("analyze risks")) return "risk";
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
  pptx.theme = { headFontFace: "Georgia", bodyFontFace: "Calibri" };
  slides.forEach((slide, i) => {
    const s = pptx.addSlide();
    s.background = { color: "0E0E12" };
    s.addText(slide.title, { x: 0.5, y: 0.3, w: 9, h: 0.9, fontSize: 28, bold: true, color: "C9A84C", fontFace: "Georgia" });
    if (i === 0) {
      s.addShape(pptx.ShapeType.line, { x: 0.5, y: 1.3, w: 9, h: 0, line: { color: "C9A84C", width: 1 } });
    }
    (slide.bullets || []).forEach((b, j) => {
      s.addText(`• ${b}`, { x: 0.7, y: 1.6 + j * 0.7, w: 8.5, h: 0.6, fontSize: 16, color: "E8E8F0", fontFace: "Calibri" });
    });
  });
  pptx.writeFile({ fileName: "LegalDeck.pptx" });
};

// ── export to word ───────────────────────────────────────────────────────────
const exportDocx = async (content, index) => {
  const lines = content.split("\n").filter(Boolean);
  const children = lines.map((line) => {
    if (line.startsWith("# ")) return new Paragraph({ text: line.slice(2), heading: HeadingLevel.HEADING_1 });
    if (line.startsWith("## ")) return new Paragraph({ text: line.slice(3), heading: HeadingLevel.HEADING_2 });
    if (line.startsWith("### ")) return new Paragraph({ text: line.slice(4), heading: HeadingLevel.HEADING_3 });
    if (line.startsWith("**") && line.endsWith("**")) return new Paragraph({ children: [new TextRun({ text: line.replace(/\*\*/g, ""), bold: true })] });
    return new Paragraph({ text: line.replace(/\*\*/g, "") });
  });
  const doc = new Document({ sections: [{ children }] });
  const blob = await Packer.toBlob(doc);
  saveAs(blob, `LegalAnalysis_${index + 1}.docx`);
};

// ── slide preview ─────────────────────────────────────────────────────────────
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
    <button className="export-btn pptx" onClick={onExport}>⬇ Download PPTX</button>
  </div>
);

// ── visual frame ──────────────────────────────────────────────────────────────
const VisualFrame = ({ html }) => (
  <div className="visual-frame">
    <iframe
      srcDoc={`<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{background:#0e0e12;color:#e8e8f0;font-family:sans-serif;padding:1rem;}</style></head><body>${html}</body></html>`}
      sandbox="allow-scripts"
      title="Visual"
      className="visual-iframe"
    />
  </div>
);

// ── main component ────────────────────────────────────────────────────────────
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

  const sendMessage = async (promptText, mode) => {
    const modeToUse = mode || detectMode(promptText);
    const userMsg = { role: "user", content: promptText };
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
        body: JSON.stringify({ messages: history, docsContext, mode: modeToUse }),
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
        return [...filtered, { role: "assistant", content: `❌ ${err.message}`, mode: "legal" }];
      });
    } finally {
      setLoading(false);
    }
  };

  const renderAssistantContent = (msg, i) => {
    if (msg.content === "...") return <span className="dots" />;

    if (msg.mode === "slides") {
      const json = extractJson(msg.content);
      if (json?.slides) return <SlidePreview slides={json.slides} onExport={() => exportPptx(json.slides)} />;
    }

    if (msg.mode === "visual") {
      const html = extractHtml(msg.content);
      if (html) return <VisualFrame html={html} />;
    }

    return (
      <>
        <ReactMarkdown remarkPlugins={[remarkGfm]} className="markdown">
          {msg.content}
        </ReactMarkdown>
        <div className="msg-actions">
          <button className="export-btn" onClick={() => exportDocx(msg.content, i)}>⬇ Export Word</button>
          <button className="export-btn copy" onClick={() => navigator.clipboard.writeText(msg.content)}>⎘ Copy</button>
        </div>
      </>
    );
  };

  return (
    <div className="app">
      {/* HEADER */}
      <header className="header">
        <button className="sidebar-toggle" onClick={() => setSidebarOpen((v) => !v)}>☰</button>
        <div className="header-logo">LEX<span>AI</span></div>
        <div className="header-tag">Senior Legal Intelligence</div>
        <button className="new-chat-btn" onClick={() => setMessages([])}>+ New Chat</button>
      </header>

      <div className="body">
        {/* SIDEBAR */}
        <aside className={`sidebar ${sidebarOpen ? "open" : "closed"}`}>
          <div className="section-label">DOCUMENTS</div>
          <label className={`upload-btn ${loadingFiles ? "loading" : ""}`}>
            {loadingFiles ? "Processing…" : "+ Upload Files"}
            <input type="file" accept=".txt,.docx,.pdf,image/*" multiple onChange={handleFiles} style={{ display: "none" }} />
          </label>
          <div className="doc-list">
            {docs.length === 0 && <p className="doc-empty">No files uploaded.</p>}
            {docs.map((doc) => (
              <div key={doc.id} className="doc-item">
                <span className="doc-icon">📄</span>
                <span className="doc-name" title={doc.name}>{doc.name}</span>
                <button className="doc-remove" onClick={() => removeDoc(doc.id)}>✕</button>
              </div>
            ))}
          </div>

          <div className="section-label" style={{ marginTop: "1.5rem" }}>QUICK ACTIONS</div>
          <div className="quick-actions">
            {QUICK_ACTIONS.map((a) => (
              <button
                key={a.id}
                className="action-btn"
                style={{ "--action-color": a.color }}
                onClick={() => sendMessage(a.prompt, a.mode)}
                disabled={loading || docs.length === 0}
                title={docs.length === 0 ? "Upload a document first" : ""}
              >
                <span className="action-icon">{a.icon}</span>
                <span>{a.label}</span>
              </button>
            ))}
          </div>

          <div className="section-label" style={{ marginTop: "1.5rem" }}>SUGGESTED</div>
          <div className="suggestions">
            {[
              "Summarize key obligations",
              "Who are the parties?",
              "What are the termination clauses?",
              "Extract all defined terms",
              "Flag any unusual provisions",
            ].map((s) => (
              <button key={s} className="suggestion-chip" onClick={() => { setQuestion(s); textareaRef.current?.focus(); }}>
                {s}
              </button>
            ))}
          </div>
        </aside>

        {/* CHAT */}
        <main className="chat">
          <div className="chat-history" ref={chatRef}>
            {messages.length === 0 ? (
              <div className="chat-empty">
                <div className="empty-icon">⚖</div>
                <div className="empty-title">LexAI Legal Assistant</div>
                <p className="empty-sub">Upload documents and ask anything, or use Quick Actions to scan for risks, generate timelines, and create slide decks.</p>
                <div className="empty-chips">
                  {["Review this contract", "What are my risks?", "Generate a summary slide deck"].map((c) => (
                    <button key={c} className="empty-chip" onClick={() => { setQuestion(c); textareaRef.current?.focus(); }}>{c}</button>
                  ))}
                </div>
              </div>
            ) : (
              messages.map((m, i) => (
                <div key={i} className={`msg ${m.role}`}>
                  <div className="msg-role">{m.role === "user" ? "You" : "⚖ LexAI"}</div>
                  <div className="msg-content">
                    {renderAssistantContent(m, i)}
                    {m.role === "user" && m.content}
                  </div>
                </div>
              ))
            )}
          </div>

          <div className="input-bar">
            <textarea
              ref={textareaRef}
              value={question}
              onChange={(e) => setQuestion(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); if (question.trim()) sendMessage(question); } }}
              placeholder="Ask LexAI anything about your documents…"
              disabled={loading}
              rows={1}
            />
            <button className="send-btn" onClick={() => question.trim() && sendMessage(question)} disabled={loading || !question.trim()}>
              {loading ? <span className="send-dots" /> : "→"}
            </button>
          </div>
        </main>
      </div>
    </div>
  );
}