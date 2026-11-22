/**
 * Configuración global de DataNexus
 *
 * 1. Despliegue en GitHub Pages (Principal):
 *    Los identificadores públicos se inyectan mediante GitHub Actions (.github/workflows/deploy.yml)
 *    desde Settings > Secrets and variables > Actions > Repository secrets:
 *    - GOOGLE_CLIENT_ID (público OAuth; restringir orígenes en Google Cloud)
 *    - MAPBOX_TOKEN (opcional, token público pk.*; restringir por URL permitida en Mapbox)
 *    NOTA DE SEGURIDAD: GEMINI_API_KEY NO se inyecta en GitHub Pages para evitar exponerla en JS público.
 *    En GitHub Pages, el usuario puede ingresarla localmente en su navegador (localStorage).
 *
 * 2. Despliegue en Vercel (Alternativo):
 *    Configura GOOGLE_CLIENT_ID, MAPBOX_TOKEN y GEMINI_API_KEY en Vercel > Environment Variables.
 *    El frontend obtendrá los IDs públicos desde /api/config y consultará Gemini
 *    de forma segura mediante el proxy serverless /api/gemini (sin exponer GEMINI_API_KEY).
 */

window.APP_CONFIG = {
  // Client ID público de Google Cloud Console (OAuth 2.0 Web Client):
  DEFAULT_GOOGLE_CLIENT_ID: 'TU_CLIENT_ID_AQUI',

  // Solo para pruebas locales explícitas; en producción usar proxy /api/gemini o localStorage del usuario:
  GEMINI_API_KEY: 'TU_GEMINI_API_KEY_AQUI',

  // Token público de Mapbox (opcional; restringir por URL permitida en Mapbox):
  MAPBOX_TOKEN: 'TU_MAPBOX_TOKEN_AQUI',

  // Endpoints Serverless para despliegues en Vercel:
  VERCEL_CONFIG_ENDPOINT: '/api/config',
  VERCEL_GEMINI_ENDPOINT: '/api/gemini'
};
