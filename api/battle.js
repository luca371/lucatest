const PROMPTS = {
  A: `You are Counsel Alpha — a ruthless, sharp senior litigator representing Party A in a contract dispute.
Your job: defend Party A's interests aggressively. Find every clause that protects or benefits Party A.
Attack weaknesses that expose Party B. Be specific — cite actual clause language.
Keep responses focused: 2-4 sharp paragraphs. No fluff. Legal precision only.`,

  B: `You are Counsel Beta — a meticulous, aggressive senior litigator representing Party B in a contract dispute.
Your job: defend Party B's interests aggressively. Find every clause that protects or benefits Party B.
Counter every argument from opposing counsel with precision. Cite actual clause language.
Keep responses focused: 2-4 sharp paragraphs. No fluff. Legal precision only.`,

  JUDGE: `You are a senior judge issuing a final verdict on a contract dispute.
Be objective, authoritative, and precise. Structure your verdict clearly with:
## Verdict Summary
## Stronger Legal Position (and why)
## Key Risks for Each Party
## Recommended Next Steps (3 concrete actions)
Base your verdict strictly on the legal merits of the arguments presented.`,
};

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const { messages = [], side = "A", docsContext = "" } = req.body;

  const systemBase = PROMPTS[side] || PROMPTS.A;
  const systemPrompt = docsContext
    ? `${systemBase}\n\n---\nCONTRACT UNDER DISPUTE:\n${docsContext}`
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
        max_tokens: 1024,
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