import { useState, useRef, useEffect } from "react";
import mammoth from "mammoth";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { saveAs } from "file-saver";
import PptxGenJS from "pptxgenjs";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";
import "./App.css";

// ── file extraction ──────────────────────────────────────────────────────────
const extractText = async (file) => {
  if (file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
    const buf = await file.arrayBuffer();
    const { value } = await mammoth.extractRawText({ arrayBuffer: buf });
    return value.trim();
  }
  if (file.type.startsWith("text/")) return await file.text();
  return `[File: ${file.name}]`;
};

// ── quick actions ────────────────────────────────────────────────────────────
const QUICK_ACTIONS = [
  {
    id: "risk",
    icon: "§",
    label: "Risk Scanner",
    display: "Scan document for legal risks",
    color: "#b84040",
    prompt: "Perform a comprehensive legal risk analysis of the uploaded documents. Use this structure:\n\n## Executive Summary\n\n## High Risk\n\n## Medium Risk\n\n## Low Risk\n\n## Recommended Actions\n\nCite clause numbers or sections where applicable. Be specific and actionable.",
    mode: "risk",
  },
  {
    id: "timeline",
    icon: "≡",
    label: "Timeline",
    display: "Extract key dates as a visual timeline",
    color: "#2a4a7a",
    prompt: "Extract all dates, deadlines, obligations, and key events from the documents. Return ONLY a self-contained HTML block starting with <div — no markdown fences, no explanation. Style it as a clean vertical timeline. Use inline styles. Light background #f7f6f3, text #1a1916, accent #2d5016.",
    mode: "visual",
  },
  {
    id: "slides",
    icon: "▨",
    label: "Slide Deck",
    display: "Generate a slide deck from documents",
    color: "#2d5016",
    prompt: "Summarize the uploaded documents as a professional legal presentation. Return ONLY valid JSON, no markdown fences, no explanation: {\"title\":\"string\",\"subtitle\":\"string\",\"slides\":[{\"type\":\"content\",\"title\":\"string\",\"bullets\":[\"string\"]}]}. Max 8 slides. First slide type should be 'cover'. Last slide type 'closing'. Other slides type 'content'.",
    mode: "slides",
  },
];

// ── helpers ──────────────────────────────────────────────────────────────────
const detectMode = (text) => {
  const t = text.toLowerCase();
  if (t.includes("slide") || t.includes("presentation") || t.includes("deck")) return "slides";
  if (t.includes("timeline") || t.includes("visual") || t.includes("diagram")) return "visual";
  if (t.includes("risk scan") || t.includes("scan for risk")) return "risk";
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
const exportPptx = (deck) => {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";

  const DARK   = "1c2b1c";
  const LIGHT  = "f7f6f3";
  const GREEN  = "4a7c2f";
  const GREEN2 = "2d5016";
  const WHITE  = "e8f0e0";
  const GREY   = "8a9e7a";

  const slides = deck.slides || [];

  slides.forEach((slide) => {
    const s = pptx.addSlide();
    s.background = { color: DARK };

    // top accent bar
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.06, fill: { color: GREEN } });

    if (slide.type === "cover") {
      // cover layout
      s.addText(deck.title || slide.title, {
        x: 0.8, y: 1.6, w: 10, h: 1.2,
        fontSize: 36, bold: true, color: WHITE,
        fontFace: "Georgia", align: "left",
      });
      s.addText(deck.subtitle || "", {
        x: 0.8, y: 2.9, w: 10, h: 0.6,
        fontSize: 16, color: GREY, fontFace: "Calibri", align: "left",
      });
      s.addShape(pptx.ShapeType.line, { x: 0.8, y: 3.6, w: 4, h: 0, line: { color: GREEN, width: 1 } });
      s.addText("Pearson AI", {
        x: 0.8, y: 4.8, w: 5, h: 0.3,
        fontSize: 9, color: GREY, fontFace: "Calibri", italic: true,
      });
    } else if (slide.type === "closing") {
      s.addText(slide.title || "Thank You", {
        x: 0, y: 1.8, w: "100%", h: 1.2,
        fontSize: 32, bold: true, color: WHITE,
        fontFace: "Georgia", align: "center",
      });
      s.addShape(pptx.ShapeType.line, { x: 4, y: 3.1, w: 5, h: 0, line: { color: GREEN, width: 1 } });
      s.addText("Pearson AI  —  Legal Intelligence", {
        x: 0, y: 3.5, w: "100%", h: 0.4,
        fontSize: 10, color: GREY, fontFace: "Calibri", align: "center",
      });
    } else {
      // content layout
      s.addText(slide.title, {
        x: 0.6, y: 0.3, w: 11.4, h: 0.75,
        fontSize: 20, bold: true, color: WHITE,
        fontFace: "Georgia", align: "left",
      });
      s.addShape(pptx.ShapeType.line, { x: 0.6, y: 1.1, w: 11.8, h: 0, line: { color: GREEN2, width: 0.75 } });

      (slide.bullets || []).forEach((bullet, j) => {
        s.addShape(pptx.ShapeType.rect, {
          x: 0.6, y: 1.3 + j * 0.68, w: 0.04, h: 0.35,
          fill: { color: GREEN },
        });
        s.addText(bullet, {
          x: 0.85, y: 1.28 + j * 0.68, w: 11.2, h: 0.44,
          fontSize: 13, color: WHITE, fontFace: "Calibri",
          valign: "middle",
        });
      });

      // slide number bottom right
      s.addText(`${slides.indexOf(slide) + 1} / ${slides.length}`, {
        x: 11.5, y: 5.3, w: 1.4, h: 0.25,
        fontSize: 7, color: GREY, fontFace: "Calibri", align: "right",
      });
    }
  });

  pptx.writeFile({ fileName: "PearsonAI_Deck.pptx" });
};

// ── word export ───────────────────────────────────────────────────────────────
const exportDocx = async (content, index) => {
  const lines = content.split("\n").filter(Boolean);
  const children = lines.map((line) => {
    if (line.startsWith("# "))  return new Paragraph({ text: line.slice(2),  heading: HeadingLevel.HEADING_1 });
    if (line.startsWith("## ")) return new Paragraph({ text: line.slice(3),  heading: HeadingLevel.HEADING_2 });
    if (line.startsWith("### "))return new Paragraph({ text: line.slice(4),  heading: HeadingLevel.HEADING_3 });
    return new Paragraph({ children: [new TextRun({ text: line.replace(/\*\*/g, "") })] });
  });
  const doc = new Document({ sections: [{ children }] });
  const blob = await Packer.toBlob(doc);
  saveAs(blob, `PearsonAI_${index + 1}.docx`);
};

// ── slide preview card ────────────────────────────────────────────────────────
const SlideCard = ({ slide, index, total }) => (
  <div className="slide-card">
    <div className="slide-num">{String(index + 1).padStart(2, "0")} / {total}</div>
    <div className="slide-title">{slide.title}</div>
    <ul className="slide-bullets">
      {(slide.bullets || []).slice(0, 4).map((b, j) => <li key={j}>{b}</li>)}
    </ul>
    <div className="slide-footer">
      <span className="slide-footer-brand">Pearson AI</span>
    </div>
  </div>
);

const SlidePreview = ({ deck, onExport }) => {
  const slides = deck.slides || [];
  return (
    <div className="slides-preview">
      <div className="slides-header">{slides.length} slides — {deck.title}</div>
      <div className="slides-grid">
        {slides.map((slide, i) => (
          <SlideCard key={i} slide={slide} index={i} total={slides.length} />
        ))}
      </div>
      <button className="action-small" onClick={onExport}>Download PPTX</button>
    </div>
  );
};

const VisualFrame = ({ html }) => (
  <div className="visual-frame">
    <iframe
      srcDoc={`<!DOCTYPE html><html><head><meta charset="utf-8"><style>*{box-sizing:border-box}body{background:#f7f6f3;color:#1a1916;font-family:Inter,sans-serif;padding:1.5rem;margin:0;font-size:13px;}</style></head><body>${html}</body></html>`}
      sandbox="allow-scripts"
      title="Visual"
      className="visual-iframe"
    />
  </div>
);

// ── main ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [docs, setDocs]               = useState([]);
  const [question, setQuestion]       = useState("");
  const [messages, setMessages]       = useState([]);
  const [loading, setLoading]         = useState(false);
  const [loadingFiles, setLoadingFiles] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const chatRef    = useRef(null);
  const textareaRef = useRef(null);

  useEffect(() => {
    chatRef.current?.scrollTo({ top: 99999, behavior: "smooth" });
  }, [messages]);

  useEffect(() => {
    if (!textareaRef.current) return;
    textareaRef.current.style.height = "auto";
    textareaRef.current.style.height = Math.min(textareaRef.current.scrollHeight, 140) + "px";
  }, [question]);

  const handleFiles = async (e) => {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    setLoadingFiles(true);
    try {
      const newDocs = await Promise.all(files.map(async (file) => {
        let text = "";
        try { text = await extractText(file); } catch { text = "[Error reading file]"; }
        return { id: crypto.randomUUID(), name: file.name, text };
      }));
      setDocs((prev) => [...prev, ...newDocs]);
    } finally {
      setLoadingFiles(false);
      e.target.value = "";
    }
  };

  const removeDoc = (id) => setDocs((prev) => prev.filter((d) => d.id !== id));

  const sendMessage = async (promptText, mode, displayText) => {
    const modeToUse   = mode || detectMode(promptText);
    const visibleText = displayText || promptText;
    const userMsg     = { role: "user", content: promptText, display: visibleText };
    const history     = [...messages, userMsg];

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
      try { const d = JSON.parse(text); reply = d.reply || d.error || text; }
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

  // ── render each message ──────────────────────────────────────────────────
  const renderMsg = (m, i) => {
    if (m.role === "user") {
      return (
        <div key={i} className="msg user">
          <div className="msg-inner">
            <div className="msg-role">You</div>
            <div className="user-content">{m.display || m.content}</div>
          </div>
        </div>
      );
    }

    // loading
    if (m.content === "...") {
      return (
        <div key={i} className="msg assistant">
          <div className="msg-inner">
            <div className="msg-role">Pearson AI</div>
            <div className="ai-content">
              <div className="dots"><span/><span/><span/></div>
            </div>
          </div>
        </div>
      );
    }

    // slides
    if (m.mode === "slides") {
      const deck = extractJson(m.content);
      return (
        <div key={i} className="msg assistant">
          <div className="msg-inner">
            <div className="msg-role">Pearson AI</div>
            <div className="ai-content">
              {deck?.slides
                ? <SlidePreview deck={deck} onExport={() => exportPptx(deck)} />
                : <ReactMarkdown remarkPlugins={[remarkGfm]} className="markdown">{m.content}</ReactMarkdown>
              }
            </div>
          </div>
        </div>
      );
    }

    // visual
    if (m.mode === "visual") {
      const html = extractHtml(m.content);
      return (
        <div key={i} className="msg assistant">
          <div className="msg-inner">
            <div className="msg-role">Pearson AI</div>
            <div className="ai-content">
              {html
                ? <VisualFrame html={html} />
                : <ReactMarkdown remarkPlugins={[remarkGfm]} className="markdown">{m.content}</ReactMarkdown>
              }
            </div>
          </div>
        </div>
      );
    }

    // default legal response
    return (
      <div key={i} className="msg assistant">
        <div className="msg-inner">
          <div className="msg-role">Pearson AI</div>
          <div className="ai-content">
            <ReactMarkdown remarkPlugins={[remarkGfm]} className="markdown">{m.content}</ReactMarkdown>
            <div className="msg-actions">
              <button className="action-small" onClick={() => exportDocx(m.content, i)}>Export Word</button>
              <button className="action-small" onClick={() => navigator.clipboard.writeText(m.content)}>Copy</button>
            </div>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="app">
      {/* HEADER */}
      <header className="header">
        <button className="sidebar-toggle" onClick={() => setSidebarOpen((v) => !v)}>
          <svg width="15" height="15" viewBox="0 0 15 15" fill="currentColor">
            <rect y="1"   width="15" height="1.5" rx="0.75"/>
            <rect y="6.75" width="15" height="1.5" rx="0.75"/>
            <rect y="12.5" width="15" height="1.5" rx="0.75"/>
          </svg>
        </button>
        <div className="header-logo">Pearson<span> AI</span></div>
        <div className="header-tag">Legal Intelligence</div>
        <button className="new-chat-btn" onClick={() => setMessages([])}>New Chat</button>
      </header>

      <div className="body">
        {/* SIDEBAR */}
        <aside className={`sidebar ${sidebarOpen ? "open" : "closed"}`}>

          <div className="section-label">Documents</div>
          <label className={`upload-btn ${loadingFiles ? "loading" : ""}`}>
            {loadingFiles ? "Processing…" : "Upload Files"}
            <input type="file" accept=".txt,.docx,image/*" multiple onChange={handleFiles} style={{ display: "none" }} />
          </label>
          <div className="doc-list">
            {docs.length === 0 && <p className="doc-empty">No files yet.</p>}
            {docs.map((doc) => (
              <div key={doc.id} className="doc-item">
                <svg className="doc-icon-svg" viewBox="0 0 14 14" fill="currentColor">
                  <path d="M3 0h6l3 3v10a1 1 0 01-1 1H3a1 1 0 01-1-1V1a1 1 0 011-1z" opacity=".12"/>
                  <path d="M3 0h6l3 3v10a1 1 0 01-1 1H3a1 1 0 01-1-1V1a1 1 0 011-1zm0 1v12h8V4H8V1H3zm5 0v2.5h2.5L8 1z"/>
                </svg>
                <span className="doc-name" title={doc.name}>{doc.name}</span>
                <button className="doc-remove" onClick={() => removeDoc(doc.id)}>
                  <svg width="9" height="9" viewBox="0 0 9 9" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
                    <path d="M1 1l7 7M8 1l-7 7"/>
                  </svg>
                </button>
              </div>
            ))}
          </div>

          <div className="section-label">Quick Actions</div>
          <div className="quick-actions">
            {QUICK_ACTIONS.map((a) => (
              <button
                key={a.id}
                className="action-btn"
                style={{ "--action-color": a.color }}
                onClick={() => sendMessage(a.prompt, a.mode, a.display)}
                disabled={loading || docs.length === 0}
                title={docs.length === 0 ? "Upload a document first" : ""}
              >
                <span className="action-icon">{a.icon}</span>
                <span>{a.label}</span>
              </button>
            ))}
          </div>

          <div className="section-label">Suggested</div>
          <div className="suggestions">
            {["Summarize key obligations", "Who are the parties?", "List termination clauses", "Extract defined terms", "Flag unusual provisions"].map((s) => (
              <button key={s} className="suggestion-chip"
                onClick={() => { setQuestion(s); textareaRef.current?.focus(); }}>
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
                <div className="empty-icon-wrap">
                  <svg width="28" height="28" viewBox="0 0 28 28" fill="none" stroke="currentColor" strokeWidth="1.2">
                    <circle cx="14" cy="14" r="12"/>
                    <path d="M9 14h10M14 9v10" strokeLinecap="round"/>
                  </svg>
                </div>
                <div className="empty-title">Pearson AI</div>
                <p className="empty-sub">Upload documents and ask anything. Use Quick Actions for risk scanning, timelines, and slide decks.</p>
                <div className="empty-chips">
                  {["Review this contract", "What are my risks?", "Summarize key terms"].map((c) => (
                    <button key={c} className="empty-chip"
                      onClick={() => { setQuestion(c); textareaRef.current?.focus(); }}>{c}</button>
                  ))}
                </div>
              </div>
            ) : (
              messages.map((m, i) => renderMsg(m, i))
            )}
          </div>

          <div className="input-bar">
            <div className="input-wrap">
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
                placeholder="Ask Pearson AI about your documents…"
                disabled={loading}
                rows={1}
              />
              <button
                className="send-btn"
                onClick={() => question.trim() && sendMessage(question)}
                disabled={loading || !question.trim()}
              >
                {loading ? <span className="send-dots" /> : (
                  <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                    <path d="M1 7h12M8 3l5 4-5 4" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
                  </svg>
                )}
              </button>
            </div>
          </div>
        </main>
      </div>
    </div>
  );
}