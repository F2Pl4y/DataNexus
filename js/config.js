/**
 * Configuración global de DataNexus
 *
 * Despliegue en GitHub Pages:
 * - Los identificadores públicos se inyectan mediante GitHub Actions (.github/workflows/deploy.yml)
 *   desde Settings > Secrets and variables > Actions > Repository secrets:
 *   - GOOGLE_CLIENT_ID (público OAuth; restringir orígenes en Google Cloud)
 *   - MAPBOX_TOKEN (opcional, token público pk.*; restringir por URL permitida en Mapbox)
 * - NOTA DE SEGURIDAD: GEMINI_API_KEY NO se inyecta en GitHub Pages para evitar exponerla en JS público.
 *   Cada usuario puede ingresar su clave personal localmente en su navegador (localStorage).
 */

window.APP_CONFIG = {
  // Client ID público de Google Cloud Console (OAuth 2.0 Web Client):
  DEFAULT_GOOGLE_CLIENT_ID: 'TU_CLIENT_ID_AQUI',

  // Solo para pruebas locales explícitas; en GitHub Pages se ingresa en el navegador del usuario:
  GEMINI_API_KEY: 'TU_GEMINI_API_KEY_AQUI',

  // Token público de Mapbox (opcional; restringir por URL permitida en Mapbox):
  MAPBOX_TOKEN: 'TU_MAPBOX_TOKEN_AQUI'
};
