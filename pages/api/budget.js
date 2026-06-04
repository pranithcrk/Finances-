import { kv } from '@vercel/kv';

const KV_KEY = 'fin_data';

export default async function handler(req, res) {
  try {
    if (req.method === 'GET') {
      const data = await kv.get(KV_KEY);
      if (!data) {
        return res.status(200).json({ data: null, firstRun: true });
      }
      return res.status(200).json({ data });
    }

    if (req.method === 'POST') {
      const { data } = req.body;
      await kv.set(KV_KEY, data);
      return res.status(200).json({ ok: true });
    }

    res.status(405).json({ error: 'Method not allowed' });
  } catch (err) {
    console.error('KV error:', err);
    // If KV isn't set up yet, return firstRun so the app still loads with defaults
    if (req.method === 'GET') {
      return res.status(200).json({ data: null, firstRun: true });
    }
    res.status(500).json({ error: err.message });
  }
}
