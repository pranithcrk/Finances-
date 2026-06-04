export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(503).json({ error: 'ANTHROPIC_API_KEY not set' });

  const payload = req.body;
  const prompt = `You are a sharp, practical financial advisor. Analyze this person's finances for ${payload.month} and give 3-5 actionable, specific suggestions. Be direct and use exact dollar amounts from the data.

DATA:
- Take-home: $${payload.income}/mo
- Fixed outflows: $${payload.fixed}/mo  
- Savings/investments: $${payload.savings}/mo
- Living budget: $${payload.living_budget}/mo, actual spent: $${payload.living_actual}/mo
- Living by category: ${JSON.stringify(payload.living_by_category)}
- Flex purchases: $${payload.flex_spent}/mo (includes $${payload.travel} travel)
- Leftover this month: $${payload.leftover}
- India loan balance: $${payload.india_loan} at 9.8%
- Subscriptions: ${JSON.stringify(payload.subscriptions)}
- Card spend: ${JSON.stringify(payload.card_spend)}
- Savings breakdown: ${JSON.stringify(payload.savings_breakdown)}

Be concise. No fluff. Use markdown bold for amounts. Max 300 words.`;

  try {
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({ model: 'claude-sonnet-4-20250514', max_tokens: 600, messages: [{ role: 'user', content: prompt }] }),
    });
    const data = await r.json();
    const text = data.content?.[0]?.text || 'No response.';
    res.status(200).json({ text });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
