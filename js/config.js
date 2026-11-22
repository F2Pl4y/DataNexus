/**
 * Configuración global de DataNexus
 *
 * 1. Despliegue en GitHub Pages (Principal):
 *    Los valores se inyectan automáticamente mediante GitHub Actions (.github/workflows/deploy.yml)
 *    desde Settings > Secrets and variables > Actions > Repository secrets:
 *    - GOOGLE_CLIENT_ID
 *    - GEMINI_API_KEY
 *    - MAPBOX_TOKEN (opcional; si no se define, usa CartoDB Dark + OSRM gratuito)
 *
 * 2. Despliegue en Vercel (Alternativo):
 *    Configura las mismas variables en Vercel > Project Settings > Environment Variables.
 *    El frontend obtendrá las variables públicas desde /api/config y consultará Gemini
 *    de forma segura mediante el proxy serverless /api/gemini (sin exponer GEMINI_API_KEY).
 */

window.APP_CONFIG = {
  // Client ID público de Google Cloud Console (OAuth 2.0 Web Client):
  DEFAULT_GOOGLE_CLIENT_ID: 'TU_CLIENT_ID_AQUI',

  // Clave de Google Gemini API (En GitHub Pages restringir por HTTP Referrer en Google Cloud):
  GEMINI_API_KEY: 'TU_GEMINI_API_KEY_AQUI',

  // Token público de Mapbox (opcional; restringir por URL permitida en Mapbox):
  MAPBOX_TOKEN: 'TU_MAPBOX_TOKEN_AQUI',

  // Endpoints Serverless para despliegues en Vercel:
  VERCEL_CONFIG_ENDPOINT: '/api/config',
  VERCEL_GEMINI_ENDPOINT: '/api/gemini'
};
