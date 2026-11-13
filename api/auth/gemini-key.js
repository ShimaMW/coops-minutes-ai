// ============================================================================
// Vercel Serverless Function: /api/auth/gemini-key
// ブラウザ直接通信（Resumable Upload & generateContent）向けに環境変数のAPIキーを返却します
// ============================================================================

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const apiKey = (process.env.GEMINI_API_KEY || process.env.NEXT_PUBLIC_GEMINI_API_KEY || '').replace(/[\s\r\n]+/g, '').trim();
  if (!apiKey) {
    return res.status(404).json({
      error: 'Vercel環境変数 GEMINI_API_KEY が未設定です。',
      code: 'NO_ENV_KEY'
    });
  }

  return res.status(200).json({ apiKey });
}
