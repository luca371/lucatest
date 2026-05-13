/* eslint-disable no-unused-vars */
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
    prompt: "Perform a comprehensive legal risk analysis of the uploaded documents. Use this structure:\n\n## Executive Summary\n\n## High Risk\n\n## Medium Risk\n\n## Low Risk\n\n## Recommended Actions\n\nCite clause numbers or sections. Be specific and actionable.",
    mode: "risk",
  },
  {
    id: "timeline",
    icon: "≡",
    label: "Timeline",
    display: "Extract key dates as a visual timeline",
    color: "#2a4a7a",
    prompt: "Extract all dates, deadlines, obligations, and key events from the documents. Return ONLY a self-contained HTML block starting with <div — no markdown fences, no explanation. Style as a clean vertical timeline with inline styles. Background #f7f6f3, text #1a1916, accent #2d5016.",
    mode: "visual",
  },
  {
    id: "slides",
    icon: "▨",
    label: "Slide Deck",
    display: "Generate a rich slide deck from documents",
    color: "#2d5016",
    prompt: `Generate a rich professional legal presentation from the uploaded documents.
Return ONLY valid JSON — no markdown fences, no explanation, nothing else.
Use a MIX of these slide types: "cover", "content", "table", "stats", "two-column", "quote", "closing".
JSON format: {"title":"string","subtitle":"string","slides":[{"type":"cover","title":"string","subtitle":"string"},{"type":"content","title":"string","bullets":["string"]},{"type":"table","title":"string","headers":["Col1","Col2","Col3"],"rows":[["a","b","c"]]},{"type":"stats","title":"string","stats":[{"value":"string","label":"string","note":"string"}]},{"type":"two-column","title":"string","left":{"heading":"string","bullets":["string"]},"right":{"heading":"string","bullets":["string"]}},{"type":"quote","title":"string","quote":"string","source":"string"},{"type":"closing","title":"string","bullets":["string"]}]}
Rules: 7-9 slides. First must be "cover", last must be "closing". Use at least 4 different types. Include one "quote" with verbatim text.`,
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

// ── pptx ─────────────────────────────────────────────────────────────────────
const T = { DARK:"171f17",DARK2:"1c2b1c",CARD:"223322",GREEN:"4a7c2f",GREEN2:"2d5016",GREEN3:"6aab4f",WHITE:"ddeedd",WHITE2:"c0d4b8",GREY:"7a9a6a",GREY2:"4a6040" };

const slideHeader = (s, title, idx, total) => {
  s.addShape("rect", { x:0,y:0,w:"100%",h:0.055, fill:{color:T.GREEN} });
  if (title) {
    s.addText(title, { x:0.55,y:0.18,w:11.3,h:0.72, fontSize:19,bold:true,color:T.WHITE,fontFace:"Georgia",align:"left",valign:"middle" });
    s.addShape("line", { x:0.55,y:0.96,w:12.4,h:0, line:{color:T.GREEN2,width:0.6} });
  }
  if (idx !== undefined) s.addText(`${idx}/${total}`, { x:11.6,y:5.3,w:1.3,h:0.25, fontSize:7,color:T.GREY2,fontFace:"Calibri",align:"right" });
};

const exportPptx = (deck) => {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  (deck.slides||[]).forEach((slide,idx) => {
    const s = pptx.addSlide();
    s.background = {color:T.DARK};
    if (slide.type==="cover") {
      s.addShape("rect",{x:0,y:0,w:"100%",h:0.055,fill:{color:T.GREEN}});
      s.addShape("rect",{x:0,y:0.055,w:0.55,h:5.65,fill:{color:T.DARK2}});
      s.addShape("rect",{x:0,y:4.7,w:"100%",h:0.8,fill:{color:T.DARK2}});
      s.addText(deck.title||slide.title,{x:0.85,y:1.3,w:11.2,h:1.8,fontSize:38,bold:true,color:T.WHITE,fontFace:"Georgia",align:"left",valign:"middle"});
      s.addText(deck.subtitle||"",{x:0.85,y:3.1,w:9,h:0.6,fontSize:15,color:T.GREY,fontFace:"Calibri"});
      s.addShape("line",{x:0.85,y:3.85,w:3.5,h:0,line:{color:T.GREEN,width:1.5}});
      s.addText("Pearson AI · Legal Intelligence",{x:0.85,y:4.9,w:8,h:0.35,fontSize:9,color:T.GREY2,fontFace:"Calibri",italic:true});
    } else if (slide.type==="content"||slide.type==="closing") {
      slideHeader(s,slide.title,idx+1,(deck.slides||[]).length);
      (slide.bullets||[]).forEach((b,j)=>{
        s.addShape("rect",{x:0.55,y:1.12+j*0.72+0.12,w:0.045,h:0.38,fill:{color:T.GREEN}});
        s.addText(b,{x:0.75,y:1.12+j*0.72,w:12,h:0.58,fontSize:13,color:T.WHITE,fontFace:"Calibri",valign:"middle"});
      });
    } else if (slide.type==="table") {
      slideHeader(s,slide.title,idx+1,(deck.slides||[]).length);
      const colW=12.4/(slide.headers||[1]).length;
      s.addTable([
        (slide.headers||[]).map(h=>({text:h,options:{bold:true,color:T.WHITE,fill:{color:T.GREEN2},fontSize:11,fontFace:"Calibri",align:"center",valign:"middle"}})),
        ...(slide.rows||[]).map((row,ri)=>row.map(cell=>({text:String(cell),options:{color:T.WHITE2,fill:{color:ri%2===0?T.DARK2:T.CARD},fontSize:11,fontFace:"Calibri",valign:"middle"}})))
      ],{x:0.55,y:1.1,w:12.4,rowH:0.5,border:{pt:0.5,color:T.GREEN2},colW:(slide.headers||[]).map(()=>colW)});
    } else if (slide.type==="stats") {
      slideHeader(s,slide.title,idx+1,(deck.slides||[]).length);
      const stats=(slide.stats||[]).slice(0,4); const count=stats.length;
      const boxW=count===2?5.6:count===3?3.7:2.8; const gap=(13-0.55-count*boxW)/(count+1);
      stats.forEach((stat,j)=>{
        const x=0.55+gap+j*(boxW+gap);
        s.addShape("rect",{x,y:1.25,w:boxW,h:3.5,fill:{color:T.CARD},line:{color:T.GREEN2,pt:0.75}});
        s.addShape("rect",{x,y:1.25,w:boxW,h:0.06,fill:{color:T.GREEN}});
        s.addText(stat.value,{x,y:1.7,w:boxW,h:1.4,fontSize:count<=2?42:34,bold:true,color:T.GREEN3,fontFace:"Georgia",align:"center",valign:"middle"});
        s.addText(stat.label,{x,y:3.1,w:boxW,h:0.55,fontSize:12,bold:true,color:T.WHITE,fontFace:"Calibri",align:"center"});
        if(stat.note) s.addText(stat.note,{x,y:3.65,w:boxW,h:0.75,fontSize:9,color:T.GREY,fontFace:"Calibri",align:"center"});
      });
    } else if (slide.type==="two-column") {
      slideHeader(s,slide.title,idx+1,(deck.slides||[]).length);
      s.addShape("line",{x:6.5,y:1.1,w:0,h:4.3,line:{color:T.GREEN2,width:0.75}});
      [{col:slide.left,x:0.55},{col:slide.right,x:6.8}].forEach(({col,x})=>{
        if(!col)return;
        if(col.heading) s.addText(col.heading,{x,y:1.15,w:5.7,h:0.5,fontSize:13,bold:true,color:T.GREEN3,fontFace:"Georgia"});
        (col.bullets||[]).forEach((b,j)=>{
          s.addShape("rect",{x,y:1.85+j*0.65+0.1,w:0.04,h:0.33,fill:{color:T.GREEN}});
          s.addText(b,{x:x+0.2,y:1.82+j*0.65,w:5.5,h:0.52,fontSize:12,color:T.WHITE,fontFace:"Calibri",valign:"middle"});
        });
      });
    } else if (slide.type==="quote") {
      slideHeader(s,slide.title,idx+1,(deck.slides||[]).length);
      s.addShape("rect",{x:0.55,y:1.15,w:0.12,h:3.2,fill:{color:T.GREEN}});
      s.addText("\u201C",{x:0.8,y:1.0,w:1.5,h:1.2,fontSize:80,color:T.GREEN2,fontFace:"Georgia"});
      s.addText(slide.quote||"",{x:0.85,y:1.6,w:11.3,h:2.8,fontSize:16,color:T.WHITE,fontFace:"Georgia",italic:true,align:"left",valign:"top",lineSpacingMultiple:1.4});
      if(slide.source) s.addText(`— ${slide.source}`,{x:0.85,y:4.55,w:11.3,h:0.4,fontSize:10,color:T.GREY,fontFace:"Calibri",align:"right"});
    }
  });
  pptx.writeFile({fileName:"PearsonAI_Deck.pptx"});
};

// ── word export ───────────────────────────────────────────────────────────────
const exportDocx = async (content, index) => {
  const lines = content.split("\n").filter(Boolean);
  const children = lines.map((line) => {
    if (line.startsWith("# "))   return new Paragraph({ text: line.slice(2),  heading: HeadingLevel.HEADING_1 });
    if (line.startsWith("## "))  return new Paragraph({ text: line.slice(3),  heading: HeadingLevel.HEADING_2 });
    if (line.startsWith("### ")) return new Paragraph({ text: line.slice(4),  heading: HeadingLevel.HEADING_3 });
    return new Paragraph({ children: [new TextRun({ text: line.replace(/\*\*/g, "") })] });
  });
  const doc = new Document({ sections: [{ children }] });
  const blob = await Packer.toBlob(doc);
  saveAs(blob, `PearsonAI_${index + 1}.docx`);
};

// ── slide cards ───────────────────────────────────────────────────────────────
const SLIDE_TYPE_LABEL = { cover:"Cover",content:"Bullets",table:"Table",stats:"Stats","two-column":"Compare",quote:"Quote",closing:"Closing" };

const SlideCard = ({ slide, index, total }) => {
  const type = slide.type || "content";
  const renderBody = () => {
    if (type==="cover") return <div className="sc-cover"><div className="sc-cover-title">{slide.title}</div><div className="sc-cover-sub">{slide.subtitle||""}</div><div className="sc-cover-line"/></div>;
    if (type==="table") return <div className="sc-table-wrap"><table className="sc-table"><thead><tr>{(slide.headers||[]).map((h,i)=><th key={i}>{h}</th>)}</tr></thead><tbody>{(slide.rows||[]).slice(0,3).map((row,i)=><tr key={i}>{row.map((c,j)=><td key={j}>{c}</td>)}</tr>)}</tbody></table></div>;
    if (type==="stats") return <div className="sc-stats">{(slide.stats||[]).slice(0,4).map((s,i)=><div key={i} className="sc-stat-box"><div className="sc-stat-value">{s.value}</div><div className="sc-stat-label">{s.label}</div></div>)}</div>;
    if (type==="two-column") return <div className="sc-two-col"><div className="sc-col"><div className="sc-col-head">{slide.left?.heading}</div>{(slide.left?.bullets||[]).slice(0,2).map((b,i)=><div key={i} className="sc-col-item">{b}</div>)}</div><div className="sc-col-divider"/><div className="sc-col"><div className="sc-col-head">{slide.right?.heading}</div>{(slide.right?.bullets||[]).slice(0,2).map((b,i)=><div key={i} className="sc-col-item">{b}</div>)}</div></div>;
    if (type==="quote") return <div className="sc-quote"><div className="sc-quote-mark">"</div><div className="sc-quote-text">{(slide.quote||"").slice(0,90)}{slide.quote?.length>90?"…":""}</div>{slide.source&&<div className="sc-quote-src">— {slide.source}</div>}</div>;
    return <ul className="sc-bullets">{(slide.bullets||[]).slice(0,4).map((b,i)=><li key={i}>{b}</li>)}</ul>;
  };
  return (
    <div className={`slide-card sc-${type}`}>
      <div className="sc-top-bar"/>
      <div className="sc-header"><span className="sc-type-badge">{SLIDE_TYPE_LABEL[type]||type}</span><span className="sc-num">{String(index+1).padStart(2,"0")}/{total}</span></div>
      {type!=="cover"&&<div className="sc-title">{slide.title}</div>}
      <div className="sc-body">{renderBody()}</div>
      <div className="sc-footer">Pearson AI</div>
    </div>
  );
};

const SlidePreview = ({ deck, onExport }) => {
  const slides = deck.slides || [];
  return (
    <div className="slides-preview">
      <div className="slides-header"><span className="slides-deck-title">{deck.title}</span><span className="slides-count">{slides.length} slides</span></div>
      <div className="slides-grid">{slides.map((s,i)=><SlideCard key={i} slide={s} index={i} total={slides.length}/>)}</div>
      <button className="action-small dl-btn" onClick={onExport}>Download PPTX</button>
    </div>
  );
};

const VisualFrame = ({ html }) => (
  <div className="visual-frame">
    <iframe srcDoc={`<!DOCTYPE html><html><head><meta charset="utf-8"><style>*{box-sizing:border-box}body{background:#f7f6f3;color:#1a1916;font-family:Inter,sans-serif;padding:1.5rem;margin:0;font-size:13px;}</style></head><body>${html}</body></html>`} sandbox="allow-scripts" title="Visual" className="visual-iframe"/>
  </div>
);

// ═══════════════════════════════════════════════════════════════════════════════
// CONTRACT BATTLE ARENA
// ═══════════════════════════════════════════════════════════════════════════════

const LAWYER_A = { id:"a", name:"Counsel Alpha", side:"Party A", color:"#2d5016", accent:"#4a7c2f", bg:"#eef3ea" };
const LAWYER_B = { id:"b", name:"Counsel Beta",  side:"Party B", color:"#7a2020", accent:"#b84040", bg:"#f5eaea" };

const BattleArena = ({ docs, onClose }) => {
  const [phase, setPhase]         = useState("intro");    // intro | battle | verdict
  const [debate, setDebate]       = useState([]);
  const [userInput, setUserInput] = useState("");
  const [loadingTurn, setLoadingTurn] = useState(false);
  const [scoreA, setScoreA]       = useState(50);
  const [verdict, setVerdict]     = useState(null);
  const debateRef                 = useRef(null);
  const roundRef                  = useRef(0);

  const docsContext = docs.length
    ? docs.map((d) => `=== ${d.name} ===\n${d.text}`).join("\n\n---\n\n")
    : "";

  useEffect(() => {
    debateRef.current?.scrollTo({ top: 99999, behavior: "smooth" });
  }, [debate]);

  const callBattle = async (messages, side) => {
    const res = await fetch("/api/battle", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ messages, side, docsContext }),
    });
    const text = await res.text();
    try { const d = JSON.parse(text); return d.reply || text; }
    catch { return text; }
  };

  const startBattle = async () => {
    setPhase("battle");
    setLoadingTurn(true);
    roundRef.current = 0;

    const opening = { role:"user", content:"Begin the debate. Analyze the contract from your client's perspective. Make your opening argument. Be sharp, specific, and cite actual clauses." };

    try {
      // Alpha opens
      const replyA = await callBattle([opening], "A");
      const msgA   = { side:"A", text: replyA, round: 0 };
      setDebate([msgA]);
      setScoreA(s => Math.min(65, s + 8));

      // Beta responds
      const replyB = await callBattle([opening, { role:"assistant", content: replyA }, { role:"user", content:"Counter Alpha's argument. Defend Party B's interests. Be aggressive and specific." }], "B");
      const msgB   = { side:"B", text: replyB, round: 0 };
      setDebate([msgA, msgB]);
      setScoreA(s => Math.max(35, s - 8));
      roundRef.current = 1;
    } catch(e) {
      console.error(e);
    } finally {
      setLoadingTurn(false);
    }
  };

  const sendIntervention = async () => {
    if (!userInput.trim() || loadingTurn) return;
    const intervention = userInput.trim();
    setUserInput("");
    setLoadingTurn(true);

    const judgeMsg = { side:"judge", text: intervention, round: roundRef.current };
    setDebate((prev) => [...prev, judgeMsg]);

    const history = debate.map((m) =>
      m.side === "judge"
        ? { role:"user", content: `[Judge intervention]: ${m.text}` }
        : { role: m.side==="A" ? "assistant" : "user", content: m.text }
    );

    try {
      const ctxA = [...history, { role:"user", content:`[Judge says]: ${intervention} — Respond as Counsel Alpha, defending Party A.` }];
      const replyA = await callBattle(ctxA, "A");
      setDebate((prev) => [...prev, { side:"A", text: replyA, round: roundRef.current }]);
      setScoreA(s => Math.min(80, s + 5));

      const ctxB = [...history, { role:"assistant", content: replyA }, { role:"user", content:`Now counter as Counsel Beta, defending Party B.` }];
      const replyB = await callBattle(ctxB, "B");
      setDebate((prev) => [...prev, { side:"B", text: replyB, round: roundRef.current }]);
      setScoreA(s => Math.max(20, s - 5));
      roundRef.current += 1;
    } catch(e) { console.error(e); }
    finally { setLoadingTurn(false); }
  };

  const requestVerdict = async () => {
    setLoadingTurn(true);
    const summary = debate.map((m) => `[${m.side==="judge"?"Judge":m.side==="A"?"Counsel Alpha":"Counsel Beta"}]: ${m.text}`).join("\n\n");
    try {
      const res = await fetch("/api/battle", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messages: [{ role:"user", content:`Based on this contract debate:\n\n${summary}\n\nIssue a final legal verdict. Identify which party has stronger legal standing, what the key risks are, and recommend 3 concrete next steps.` }], side:"JUDGE", docsContext }),
      });
      const text = await res.text();
      let v;
      try { v = JSON.parse(text).reply || text; } catch { v = text; }
      setVerdict(v);
      setPhase("verdict");
    } catch(e) { console.error(e); }
    finally { setLoadingTurn(false); }
  };

  // ── INTRO SCREEN ───────────────────────────────────────────────────────────
  if (phase === "intro") return (
    <div className="arena-overlay">
      <div className="arena-intro">
        <button className="arena-close" onClick={onClose}>✕</button>
        <div className="arena-intro-badge">CONTRACT BATTLE</div>
        <h1 className="arena-intro-title">Two lawyers.<br/>One contract.<br/>No mercy.</h1>
        <p className="arena-intro-sub">Two AI counsel will argue the contract from opposing sides. You play judge — intervene, tip the scales, and issue the final verdict.</p>
        <div className="arena-lawyers-preview">
          <div className="arena-lawyer-card" style={{"--lc": LAWYER_A.accent}}>
            <div className="alc-name">{LAWYER_A.name}</div>
            <div className="alc-side">Represents {LAWYER_A.side}</div>
          </div>
          <div className="arena-vs">VS</div>
          <div className="arena-lawyer-card" style={{"--lc": LAWYER_B.accent}}>
            <div className="alc-name">{LAWYER_B.name}</div>
            <div className="alc-side">Represents {LAWYER_B.side}</div>
          </div>
        </div>
        {docs.length === 0
          ? <p className="arena-no-doc">Upload a document first to start the battle.</p>
          : <button className="arena-start-btn" onClick={startBattle}>Start Battle</button>
        }
      </div>
    </div>
  );

  // ── VERDICT ────────────────────────────────────────────────────────────────
  if (phase === "verdict") return (
    <div className="arena-overlay">
      <div className="arena-verdict">
        <button className="arena-close" onClick={onClose}>✕</button>
        <div className="arena-intro-badge">FINAL VERDICT</div>
        <div className="arena-score-bar">
          <span className="score-label-a">Party A — {scoreA}%</span>
          <div className="score-track">
            <div className="score-fill-a" style={{width:`${scoreA}%`}}/>
            <div className="score-fill-b" style={{width:`${100-scoreA}%`}}/>
          </div>
          <span className="score-label-b">{100-scoreA}% — Party B</span>
        </div>
        <div className="verdict-body">
          <ReactMarkdown remarkPlugins={[remarkGfm]} className="markdown">{verdict}</ReactMarkdown>
        </div>
        <div className="verdict-actions">
          <button className="arena-start-btn" onClick={() => exportDocx(verdict, 99)}>Export Verdict</button>
          <button className="action-small" onClick={onClose}>Close</button>
        </div>
      </div>
    </div>
  );

  // ── BATTLE ─────────────────────────────────────────────────────────────────
  return (
    <div className="arena-overlay">
      <div className="arena-battle">
        <div className="arena-battle-header">
          <div className="arena-header-left">
            <span className="arena-intro-badge" style={{fontSize:"0.55rem"}}>CONTRACT BATTLE</span>
          </div>
          <div className="arena-score-bar compact">
            <span className="score-label-a">A — {scoreA}%</span>
            <div className="score-track">
              <div className="score-fill-a" style={{width:`${scoreA}%`}}/>
              <div className="score-fill-b" style={{width:`${100-scoreA}%`}}/>
            </div>
            <span className="score-label-b">{100-scoreA}% — B</span>
          </div>
          <div className="arena-header-right">
            <button className="arena-verdict-btn" onClick={requestVerdict} disabled={loadingTurn}>Issue Verdict</button>
            <button className="arena-close-sm" onClick={onClose}>✕</button>
          </div>
        </div>

        <div className="arena-columns">
          {/* Counsel Alpha */}
          <div className="arena-col col-a">
            <div className="col-lawyer-name" style={{color: LAWYER_A.accent}}>{LAWYER_A.name}</div>
            <div className="col-lawyer-role">Party A · Prosecution</div>
          </div>
          {/* Debate feed */}
          <div className="arena-feed" ref={debateRef}>
            {loadingTurn && debate.length === 0 && (
              <div className="arena-loading">
                <div className="dots"><span/><span/><span/></div>
                <span>Opening arguments being prepared…</span>
              </div>
            )}
            {debate.map((msg, i) => (
              <div key={i} className={`debate-msg dm-${msg.side}`}>
                <div className="dm-author">
                  {msg.side==="A" ? LAWYER_A.name : msg.side==="B" ? LAWYER_B.name : "Judge"}
                </div>
                <div className="dm-text">
                  <ReactMarkdown remarkPlugins={[remarkGfm]}>{msg.text}</ReactMarkdown>
                </div>
              </div>
            ))}
            {loadingTurn && debate.length > 0 && (
              <div className="arena-loading"><div className="dots"><span/><span/><span/></div></div>
            )}
          </div>
          {/* Counsel Beta */}
          <div className="arena-col col-b">
            <div className="col-lawyer-name" style={{color: LAWYER_B.accent}}>{LAWYER_B.name}</div>
            <div className="col-lawyer-role">Party B · Defence</div>
          </div>
        </div>

        {/* Judge input */}
        <div className="arena-judge-bar">
          <div className="judge-label">Judge Intervention</div>
          <div className="judge-input-wrap">
            <input
              value={userInput}
              onChange={(e) => setUserInput(e.target.value)}
              onKeyDown={(e) => { if(e.key==="Enter") sendIntervention(); }}
              placeholder="Ask a specific question, challenge an argument, or direct counsel…"
              disabled={loadingTurn}
            />
            <button onClick={sendIntervention} disabled={loadingTurn || !userInput.trim()}>Submit</button>
          </div>
        </div>
      </div>
    </div>
  );
};

// ── main ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [docs, setDocs]                 = useState([]);
  const [question, setQuestion]         = useState("");
  const [messages, setMessages]         = useState([]);
  const [loading, setLoading]           = useState(false);
  const [loadingFiles, setLoadingFiles] = useState(false);
  const [sidebarOpen, setSidebarOpen]   = useState(true);
  const [showBattle, setShowBattle]     = useState(false);
  const chatRef     = useRef(null);
  const textareaRef = useRef(null);

  useEffect(() => { chatRef.current?.scrollTo({ top: 99999, behavior: "smooth" }); }, [messages]);
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
        let text = ""; try { text = await extractText(file); } catch { text = "[Error]"; }
        return { id: crypto.randomUUID(), name: file.name, text };
      }));
      setDocs((prev) => [...prev, ...newDocs]);
    } finally { setLoadingFiles(false); e.target.value = ""; }
  };

  const removeDoc = (id) => setDocs((prev) => prev.filter((d) => d.id !== id));

  const sendMessage = async (promptText, mode, displayText) => {
    const modeToUse   = mode || detectMode(promptText);
    const visibleText = displayText || promptText;
    const userMsg     = { role: "user", content: promptText, display: visibleText };
    const history     = [...messages, userMsg];
    setMessages([...history, { role: "assistant", content: "...", mode: modeToUse }]);
    setQuestion(""); setLoading(true);
    const docsContext = docs.length ? docs.map((d) => `=== ${d.name} ===\n${d.text}`).join("\n\n---\n\n") : "";
    try {
      const res = await fetch("/api/analyze", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messages: history.map((m) => ({ role: m.role, content: m.content })), docsContext, mode: modeToUse }),
      });
      const text = await res.text();
      let reply; try { const d = JSON.parse(text); reply = d.reply || d.error || text; } catch { reply = text; }
      setMessages((prev) => { const f=prev.filter((m)=>!(m.role==="assistant"&&m.content==="...")); return [...f,{role:"assistant",content:reply,mode:modeToUse}]; });
    } catch(err) {
      setMessages((prev) => { const f=prev.filter((m)=>!(m.role==="assistant"&&m.content==="...")); return [...f,{role:"assistant",content:`Error: ${err.message}`,mode:"legal"}]; });
    } finally { setLoading(false); }
  };

  const renderMsg = (m, i) => {
    if (m.role==="user") return <div key={i} className="msg user"><div className="msg-inner"><div className="msg-role">You</div><div className="user-content">{m.display||m.content}</div></div></div>;
    if (m.content==="...") return <div key={i} className="msg assistant"><div className="msg-inner"><div className="msg-role">Pearson AI</div><div className="ai-content"><div className="dots"><span/><span/><span/></div></div></div></div>;
    if (m.mode==="slides") { const deck=extractJson(m.content); return <div key={i} className="msg assistant"><div className="msg-inner"><div className="msg-role">Pearson AI</div><div className="ai-content">{deck?.slides?<SlidePreview deck={deck} onExport={()=>exportPptx(deck)}/>:<ReactMarkdown remarkPlugins={[remarkGfm]} className="markdown">{m.content}</ReactMarkdown>}</div></div></div>; }
    if (m.mode==="visual") { const html=extractHtml(m.content); return <div key={i} className="msg assistant"><div className="msg-inner"><div className="msg-role">Pearson AI</div><div className="ai-content">{html?<VisualFrame html={html}/>:<ReactMarkdown remarkPlugins={[remarkGfm]} className="markdown">{m.content}</ReactMarkdown>}</div></div></div>; }
    return <div key={i} className="msg assistant"><div className="msg-inner"><div className="msg-role">Pearson AI</div><div className="ai-content"><ReactMarkdown remarkPlugins={[remarkGfm]} className="markdown">{m.content}</ReactMarkdown><div className="msg-actions"><button className="action-small" onClick={()=>exportDocx(m.content,i)}>Export Word</button><button className="action-small" onClick={()=>navigator.clipboard.writeText(m.content)}>Copy</button></div></div></div></div>;
  };

  return (
    <div className="app">
      {showBattle && <BattleArena docs={docs} onClose={() => setShowBattle(false)} />}

      <header className="header">
        <button className="sidebar-toggle" onClick={() => setSidebarOpen((v) => !v)}>
          <svg width="15" height="12" viewBox="0 0 15 12" fill="currentColor"><rect y="0" width="15" height="1.5" rx="0.75"/><rect y="5.25" width="15" height="1.5" rx="0.75"/><rect y="10.5" width="15" height="1.5" rx="0.75"/></svg>
        </button>
        <div className="header-logo">Pearson<span> AI</span></div>
        <div className="header-tag">Legal Intelligence</div>
        <button className="battle-trigger-btn" onClick={() => setShowBattle(true)} title={docs.length===0?"Upload a document first":""}>
          ⚔ Contract Battle
        </button>
        <button className="new-chat-btn" onClick={() => setMessages([])}>New Chat</button>
      </header>

      <div className="body">
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
                <svg className="doc-icon-svg" viewBox="0 0 14 14" fill="currentColor"><path d="M3 0h6l3 3v10a1 1 0 01-1 1H3a1 1 0 01-1-1V1a1 1 0 011-1z" opacity=".12"/><path d="M3 0h6l3 3v10a1 1 0 01-1 1H3a1 1 0 01-1-1V1a1 1 0 011-1zm0 1v12h8V4H8V1H3zm5 0v2.5h2.5L8 1z"/></svg>
                <span className="doc-name" title={doc.name}>{doc.name}</span>
                <button className="doc-remove" onClick={() => removeDoc(doc.id)}><svg width="9" height="9" viewBox="0 0 9 9" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"><path d="M1 1l7 7M8 1l-7 7"/></svg></button>
              </div>
            ))}
          </div>
          <div className="section-label">Quick Actions</div>
          <div className="quick-actions">
            {QUICK_ACTIONS.map((a) => (
              <button key={a.id} className="action-btn" style={{"--action-color":a.color}}
                onClick={() => sendMessage(a.prompt, a.mode, a.display)}
                disabled={loading || docs.length === 0}
                title={docs.length===0?"Upload a document first":""}>
                <span className="action-icon">{a.icon}</span><span>{a.label}</span>
              </button>
            ))}
          </div>
          <div className="section-label">Suggested</div>
          <div className="suggestions">
            {["Summarize key obligations","Who are the parties?","List termination clauses","Extract defined terms","Flag unusual provisions"].map((s) => (
              <button key={s} className="suggestion-chip" onClick={() => { setQuestion(s); textareaRef.current?.focus(); }}>{s}</button>
            ))}
          </div>
        </aside>

        <main className="chat">
          <div className="chat-history" ref={chatRef}>
            {messages.length === 0 ? (
              <div className="chat-empty">
                <div className="empty-icon-wrap"><svg width="28" height="28" viewBox="0 0 28 28" fill="none" stroke="currentColor" strokeWidth="1.2"><circle cx="14" cy="14" r="12"/><path d="M9 14h10M14 9v10" strokeLinecap="round"/></svg></div>
                <div className="empty-title">Pearson AI</div>
                <p className="empty-sub">Upload documents and ask anything — or start a <strong>Contract Battle</strong> to watch two AI lawyers argue your contract.</p>
                <div className="empty-chips">
                  {["Review this contract","What are my risks?","Summarize key terms"].map((c) => (
                    <button key={c} className="empty-chip" onClick={() => { setQuestion(c); textareaRef.current?.focus(); }}>{c}</button>
                  ))}
                </div>
              </div>
            ) : messages.map((m, i) => renderMsg(m, i))}
          </div>
          <div className="input-bar">
            <div className="input-wrap">
              <textarea ref={textareaRef} value={question}
                onChange={(e) => setQuestion(e.target.value)}
                onKeyDown={(e) => { if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();if(question.trim())sendMessage(question);} }}
                placeholder="Ask Pearson AI about your documents…" disabled={loading} rows={1}/>
              <button className="send-btn" onClick={() => question.trim()&&sendMessage(question)} disabled={loading||!question.trim()}>
                {loading?<span className="send-dots"/>:<svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M1 7h12M8 3l5 4-5 4"/></svg>}
              </button>
            </div>
          </div>
        </main>
      </div>
    </div>
  );
}