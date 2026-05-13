const SYSTEM_PROMPTS = {
  legal: `You are Pearson AI — a senior legal advisor. Answer immediately, use markdown headers and bold, be precise and actionable. End complex answers with ## Next Steps. When documents are uploaded they are the absolute source of truth.`,

  risk: `You are Pearson AI — a legal risk analyst. Use exactly this structure:
## Executive Summary
## High Risk
## Medium Risk
## Low Risk
## Recommended Actions
Reference clause numbers and exact wording.`,

  visual: `You are Pearson AI. Return ONLY a self-contained HTML block starting with <div. No markdown fences. No explanation. Inline styles only. Background #f7f6f3, text #1a1916, accent #2d5016.`,

  slides: `You are Pearson AI. Return ONLY valid JSON — no explanation, no markdown fences.
Available types: "cover","content","table","stats","two-column","quote","closing".
Format: {"title":"string","subtitle":"string","slides":[...]}
Rules: 7-9 slides, first "cover", last "closing", at least 4 different types, one verbatim "quote".`,

  signatures: `You are Pearson AI. Extract all signature blocks and execution details from the document.
Return ONLY valid JSON, no markdown fences, no explanation:
{"signatories":[{"name":"string","title":"string","party":"string","company":"string","date":"string","location":"string"}]}
If a field is not found leave it as empty string.`,

  obligations: `You are Pearson AI. Extract every obligation, duty, deadline, and commitment from the document.
Return ONLY valid JSON, no markdown fences, no explanation:
{"obligations":[{"party":"string","obligation":"string","deadline":"string","consequence":"string","status":"Pending"}]}
Status must be exactly one of: "Pending", "Due", "Completed".
If deadline is not specified, use "Not specified". Extract ALL obligations, be exhaustive.`,
};

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const { messages = [], docsContext = "", mode = "legal" } = req.body;

  const systemBase = SYSTEM_PROMPTS[mode] || SYSTEM_PROMPTS.legal;
  const systemPrompt = docsContext
    ? `${systemBase}\n\n---\nDOCUMENTS:\n${docsContext}`
    : systemBase;

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "x-api-key": process.env.CLAUDE_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
      },
      body: JSON.stringify({
        model: "claude-opus-4-7",
        max_tokens: 4096,
        system: systemPrompt,
        messages: messages.map((m) => ({ role: m.role, content: m.content })),
      }),
    });

    if (!response.ok) {
      const err = await response.text();
      return res.status(response.status).json({ error: err });
    }

    const data = await response.json();
    const reply = data.content?.[0]?.text || "No response.";
    return res.status(200).json({ reply });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}  