// ============================================================================
// Vercel Serverless Function: /api/auth/gemini-key
// URL直叩き(GET)および外部ドメインからのアクセスを遮断し、
// 同一オリジンのアプリ本体(POST + 専用ヘッダー)からのみAPIキーを返却します
// ============================================================================

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.setHeader('Pragma', 'no-cache');

  // URL直叩き(GET)や他メソッドは拒否
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  // アプリ専用ヘッダーの検証
  const clientHeader = req.headers['x-carecore-client'];
  if (clientHeader !== '1') {
    return res.status(403).json({ error: 'Forbidden' });
  }

  // 同一オリジン（Origin / Referer と Host の一致）検証
  const host = (req.headers['x-forwarded-host'] || req.headers.host || '').split(',')[0].trim();
  const origin = req.headers.origin || '';
  const referer = req.headers.referer || '';

  let isSameOrigin = false;
  if (host) {
    try {
      if (origin && new URL(origin).host === host) isSameOrigin = true;
      if (referer && new URL(referer).host === host) isSameOrigin = true;
    } catch (e) {}
  }

  if (!isSameOrigin) {
    return res.status(403).json({ error: 'Forbidden: Invalid Origin' });
  }

  const apiKey = (process.env.GEMINI_API_KEY || process.env.NEXT_PUBLIC_GEMINI_API_KEY || '')
    .replace(/[\s\r\n]+/g, '')
    .trim();

  if (!apiKey) {
    return res.status(404).json({
      error: 'Vercel環境変数 GEMINI_API_KEY が未設定です。',
      code: 'NO_ENV_KEY'
    });
  }

  return res.status(200).json({ apiKey });
}
