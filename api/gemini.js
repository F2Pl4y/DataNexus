/**
 * Vercel Serverless Function: /api/gemini
 * Proxy seguro para consultar Google Gemini API desde el servidor en Vercel
 * sin exponer GEMINI_API_KEY en el navegador.
 */

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ error: { message: 'Método no permitido' } });
  }

  const apiKey = (process.env.GEMINI_API_KEY || '').trim();
  if (!apiKey) {
    return res.status(500).json({
      error: { message: 'GEMINI_API_KEY no está configurada en las variables de entorno de Vercel.' }
    });
  }

  try {
    const { model = 'gemini-2.5-flash-lite', version = 'v1beta', bodyPayload } = req.body || {};

    // Validar formato de nombre de modelo y versión para evitar SSRF / Path Traversal
    if (!/^(v1|v1beta)$/.test(version) || !/^[a-zA-Z0-9._-]+$/.test(model)) {
      return res.status(400).json({ error: { message: 'Parámetros de modelo no válidos.' } });
    }

    if (!bodyPayload || typeof bodyPayload !== 'object') {
      return res.status(400).json({ error: { message: 'Payload inválido.' } });
    }

    const url = `https://generativelanguage.googleapis.com/${version}/models/${model}:generateContent`;
    const upstream = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-goog-api-key': apiKey
      },
      body: JSON.stringify(bodyPayload)
    });

    const data = await upstream.json().catch(() => ({}));
    return res.status(upstream.status).json(data);
  } catch (err) {
    return res.status(500).json({
      error: { message: err.message || 'Error interno al conectar con Gemini API' }
    });
  }
};
