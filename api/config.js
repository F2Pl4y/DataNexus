/**
 * Vercel Serverless Function: /api/config
 * Expone únicamente los identificadores públicos de cliente (GOOGLE_CLIENT_ID y MAPBOX_TOKEN)
 * e informa si el proxy de Gemini (/api/gemini) está disponible en el servidor.
 * NUNCA expone GEMINI_API_KEY al navegador.
 */

module.exports = function handler(req, res) {
  if (req.method !== 'GET') {
    res.setHeader('Allow', 'GET');
    return res.status(405).json({ error: 'Método no permitido' });
  }

  res.setHeader('Cache-Control', 'no-store, max-age=0');
  res.setHeader('Content-Type', 'application/json; charset=utf-8');

  return res.status(200).json({
    clientId: (process.env.GOOGLE_CLIENT_ID || '').trim(),
    mapboxToken: (process.env.MAPBOX_TOKEN || '').trim(),
    hasGeminiProxy: Boolean(process.env.GEMINI_API_KEY && process.env.GEMINI_API_KEY.trim())
  });
};
