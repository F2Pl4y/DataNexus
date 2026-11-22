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

    if (!bodyPayload || typeof bodyPayload !== 'object' || !Array.isArray(bodyPayload.contents)) {
      return res.status(400).json({ error: { message: 'Payload inválido.' } });
    }

    if (bodyPayload.contents.length === 0 || bodyPayload.contents.length > 20) {
      return res.status(400).json({ error: { message: 'Cantidad de turnos fuera de rango.' } });
    }

    const serializedBody = JSON.stringify(bodyPayload);
    if (serializedBody.length > 150000) {
      return res.status(413).json({ error: { message: 'La consulta supera el tamaño máximo permitido.' } });
    }

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 20000);

    try {
      const url = `https://generativelanguage.googleapis.com/${version}/models/${model}:generateContent`;
      const upstream = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-goog-api-key': apiKey
        },
        body: serializedBody,
        signal: controller.signal
      });

      const data = await upstream.json().catch(() => ({}));
      return res.status(upstream.status).json(data);
    } finally {
      clearTimeout(timeoutId);
    }
  } catch (err) {
    const isAbort = err && err.name === 'AbortError';
    return res.status(isAbort ? 504 : 500).json({
      error: { message: isAbort ? 'Tiempo de espera agotado al conectar con Gemini API' : 'Error interno al conectar con Gemini API' }
    });
  }
};
