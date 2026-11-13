// ============================================================================
// Vercel Serverless Function: /api/gemini
// ブラウザ直接通信用のAPIキー取得（getApiKey）および
// Resumable Upload 初期化・自動リトライ付き生成のフォールバックを提供します
// ============================================================================

const FALLBACK_MODELS = ['gemini-3.6-flash', 'gemini-3.5-flash', 'gemini-2.5-flash'];

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const apiKey = (process.env.GEMINI_API_KEY || process.env.NEXT_PUBLIC_GEMINI_API_KEY || '').replace(/[\s\r\n]+/g, '').trim();

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  if (!apiKey) {
    return res.status(400).json({
      error: 'Vercel環境変数 GEMINI_API_KEY が未設定です。画面右上の設定からAPIキーを設定するか、Vercelに環境変数を追加してください。',
      code: 'NO_ENV_KEY'
    });
  }

  try {
    const { action, payload = {} } = req.body || {};

    // 1. 接続確認（生のAPIキーは返却しない）
    if (action === 'ping') {
      return res.status(200).json({
        success: true,
        hasKey: true,
        message: 'Vercel Gemini API エンドポイントが有効です'
      });
    }

    // 2. Resumable Upload セッションURLの発行（大容量音声用）
    if (action === 'initUpload') {
      const { mimeType = 'audio/mp3', numBytes, displayName = 'meeting_audio', origin } = payload;
      const clientOrigin = req.headers.origin || origin || 'https://coops-minutes-ai.vercel.app';
      const startUploadRes = await fetch(
        `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${apiKey}`,
        {
          method: 'POST',
          headers: {
            'Origin': clientOrigin,
            'X-Goog-Upload-Protocol': 'resumable',
            'X-Goog-Upload-Command': 'start',
            'X-Goog-Upload-Header-Content-Length': String(numBytes),
            'X-Goog-Upload-Header-Content-Type': mimeType,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ file: { display_name: displayName } })
        }
      );

      if (!startUploadRes.ok) {
        const errText = await startUploadRes.text();
        throw new Error(`Upload初期化失敗: ${errText}`);
      }

      const uploadUrl = startUploadRes.headers.get('x-goog-upload-url');
      return res.status(200).json({ uploadUrl });
    }

    // 3. 議事録生成リクエスト（モデルフォールバック＆リトライ付き）
    if (action === 'generate') {
      const {
        prompt,
        fileUri,
        mimeType,
        modelName = 'gemini-3.6-flash',
        fileNameToCleanUp,
        responseSchema
      } = payload;

      const parts = [{ text: prompt }];
      if (fileUri) {
        parts.push({
          file_data: {
            mime_type: mimeType || 'audio/mp3',
            file_uri: fileUri
          }
        });
      }

      const models = Array.from(new Set([modelName, ...FALLBACK_MODELS].filter(Boolean)));
      let lastError = null;
      let resultData = null;

      try {
        for (const currentModel of models) {
          for (let attempt = 0; attempt <= 2; attempt++) {
            const geminiRes = await fetch(
              `https://generativelanguage.googleapis.com/v1beta/models/${currentModel}:generateContent?key=${apiKey}`,
              {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                  contents: [{ parts }],
                  generationConfig: {
                    responseMimeType: 'application/json',
                    temperature: 0.2,
                    responseSchema
                  }
                })
              }
            );

            const json = await geminiRes.json().catch(() => ({}));
            const errMsg = json?.error?.message || `HTTP ${geminiRes.status}`;

            if (!geminiRes.ok || json.error) {
              lastError = new Error(errMsg);
              if (geminiRes.status === 404 || errMsg.toLowerCase().includes('not found')) {
                break; // 次のモデルへ
              }
              if ([429, 500, 502, 503, 504].includes(geminiRes.status) && attempt < 2) {
                await sleep(Math.pow(2, attempt + 1) * 1000 + Math.floor(Math.random() * 500));
                continue;
              }
              break;
            }

            const rawText = json.candidates?.[0]?.content?.parts?.[0]?.text;
            if (!rawText) {
              lastError = new Error('AIからの有効なテキスト応答がありませんでした。');
              break;
            }

            let cleaned = rawText.trim();
            if (cleaned.startsWith('```')) {
              cleaned = cleaned.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();
            }
            resultData = JSON.parse(cleaned);
            break;
          }
          if (resultData) break;
        }
      } finally {
        if (fileNameToCleanUp) {
          fetch(`https://generativelanguage.googleapis.com/v1beta/${fileNameToCleanUp}?key=${apiKey}`, {
            method: 'DELETE'
          }).catch(() => {});
        }
      }

      if (!resultData) {
        return res.status(500).json({ error: `【AI解析エラー】: ${lastError ? lastError.message : '生成に失敗しました'}` });
      }

      return res.status(200).json(resultData);
    }

    return res.status(400).json({ error: `Unknown action: ${action}` });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
