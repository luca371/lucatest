import { useState, useRef } from "react";
import mammoth from "mammoth";
import "./App.css";

const extractText = async (file) => {
  if (file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
    const buf = await file.arrayBuffer();
    const { value } = await mammoth.extractRawText({ arrayBuffer: buf });
    return value.trim();
  }
  if (file.type.startsWith("text/")) return await file.text();
  return `[Image: ${file.name}]`;
};

export default function App() {
  const [docs, setDocs] = useState([]);
  const [question, setQuestion] = useState("");
  const [messages, setMessages] = useState([]);
  const [loading, setLoading] = useState(false);
  const [loadingFiles, setLoadingFiles] = useState(false);
  const chatRef = useRef(null);

  const handleFiles = async (e) => {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    setLoadingFiles(true);
    try {
      const newDocs = await Promise.all(
        files.map(async (file) => {
          let text = "";
          try { text = await extractText(file); }
          catch { text = "[Error reading file]"; }
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

  const ask = async () => {
    if (!question.trim() || loading || loadingFiles) return;

    const userMsg = { role: "user", content: question };
    const history = [...messages, userMsg];
    setMessages([...history, { role: "assistant", content: "..." }]);
    setQuestion("");
    setLoading(true);

    const docsContext = docs.length
      ? docs.map((d) => `=== ${d.name} ===\n${d.text}`).join("\n\n---\n\n")
      : "";

    try {
      const res = await fetch("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messages: history, docsContext }),
      });

      const text = await res.text();
      let reply;
      try {
        const data = JSON.parse(text);
        reply = data.reply || data.error || text;
      } catch {
        reply = text;
      }

      setMessages((prev) => {
        const filtered = prev.filter((m) => !(m.role === "assistant" && m.content === "..."));
        return [...filtered, { role: "assistant", content: reply }];
      });
    } catch (err) {
      setMessages((prev) => {
        const filtered = prev.filter((m) => !(m.role === "assistant" && m.content === "..."));
        return [...filtered, { role: "assistant", content: `❌ ${err.message}` }];
      });
    } finally {
      setLoading(false);
      setTimeout(() => chatRef.current?.scrollTo({ top: 99999, behavior: "smooth" }), 50);
    }
  };

  return (
    <div className="app">
      <aside className="sidebar">
        <div className="sidebar-logo">DOCAI</div>
        <p className="sidebar-sub">Drop documents, ask anything.</p>

        <label className={`upload-btn ${loadingFiles ? "loading" : ""}`}>
          {loadingFiles ? "Processing…" : "+ Add Files"}
          <input type="file" accept=".txt,.docx,image/*" multiple onChange={handleFiles} style={{ display: "none" }} />
        </label>

        <div className="doc-list">
          {docs.length === 0 && <p className="doc-empty">No documents yet.</p>}
          {docs.map((doc) => (
            <div key={doc.id} className="doc-item">
              <span className="doc-name" title={doc.name}>{doc.name}</span>
              <button className="doc-remove" onClick={() => removeDoc(doc.id)}>✕</button>
            </div>
          ))}
        </div>
      </aside>

      <main className="chat">
        <div className="chat-history" ref={chatRef}>
          {messages.length === 0 ? (
            <div className="chat-empty">
              <span>Upload a document and ask a question.</span>
            </div>
          ) : (
            messages.map((m, i) => (
              <div key={i} className={`msg ${m.role}`}>
                <div className="msg-role">{m.role === "user" ? "You" : "AI"}</div>
                <div className="msg-content">
                  {m.content === "..." ? <span className="dots" /> : m.content}
                </div>
              </div>
            ))
          )}
        </div>

        <div className="input-bar">
          <textarea
            value={question}
            onChange={(e) => setQuestion(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); ask(); } }}
            placeholder={docs.length ? "Ask about your documents…" : "Upload documents or ask anything…"}
            disabled={loading}
            rows={1}
          />
          <button onClick={ask} disabled={loading || !question.trim()}>
            {loading ? "…" : "→"}
          </button>
        </div>
      </main>
    </div>
  );
}