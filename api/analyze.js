const SYSTEM_PROMPTS = {
  legal: `You are Pearson AI — a senior legal advisor with 25 years of experience in corporate and commercial law across Europe.

STYLE:
- Answer immediately, no preamble
- Use markdown: ## headers, **bold** for key terms, tables for comparisons
- Be precise, structured, and actionable
- End complex answers with a ## Next Steps section

DOCUMENT RULE:
When documents are uploaded they are the absolute source of truth. Quote exact wording when relevant.`,

  risk: `You are Pearson AI — a legal risk analyst.

Use exactly this structure:

## Executive Summary

## High Risk

## Medium Risk

## Low Risk

## Recommended Actions

Reference clause numbers and exact wording from the document. Be specific.`,

  visual: `You are Pearson AI. Return ONLY a self-contained HTML block starting with <div.
No markdown fences. No explanation. Inline styles only.
Light theme: background #f7f6f3, text #1a1916, accent #2d5016, border #e4e2da.`,

  slides: `You are Pearson AI. Generate a professional legal presentation.
Return ONLY valid JSON — no explanation, no markdown fences, nothing else.

Format:
{
  "title": "presentation title",
  "subtitle": "document name or context",
  "slides": [
    { "type": "cover",   "title": "same as presentation title", "bullets": [] },
    { "type": "content", "title": "slide title", "bullets": ["point 1", "point 2", "point 3"] },
    { "type": "closing", "title": "Key Takeaways", "bullets": ["takeaway 1", "takeaway 2"] }
  ]
}

Rules:
- 6 to 8 slides total
- First slide type must be "cover"
- Last slide type must be "closing"
- All other slides type "content"
- Max 5 bullets per slide
- Each bullet under 12 words
- Bullets must be substantive legal points, not generic filler`,
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