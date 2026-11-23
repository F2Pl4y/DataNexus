const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const ROOT = path.resolve(__dirname, '..');

function loadBrowserScript(relPath, extraGlobals = {}) {
  const code = fs.readFileSync(path.join(ROOT, relPath), 'utf8');
  const sandbox = {
    window: {
      location: { origin: 'https://f2pl4y.github.io', pathname: '/DataNexus/', hash: '' }
    },
    localStorage: {
      _data: new Map(),
      getItem(k) { return this._data.has(k) ? this._data.get(k) : null; },
      setItem(k, v) { this._data.set(k, String(v)); },
      removeItem(k) { this._data.delete(k); }
    },
    history: { replaceState() {} },
    navigator: {},
    document: {},
    console,
    Date,
    Math,
    JSON,
    Set,
    Map,
    Array,
    Object,
    String,
    Number,
    Boolean,
    RegExp,
    Error,
    Promise,
    encodeURIComponent,
    decodeURIComponent,
    btoa: (s) => Buffer.from(s, 'binary').toString('base64'),
    atob: (s) => Buffer.from(s, 'base64').toString('binary'),
    escape: global.escape,
    unescape: global.unescape,
    ...extraGlobals
  };
  sandbox.window.localStorage = sandbox.localStorage;
  sandbox.window.history = sandbox.history;
  vm.runInNewContext(code, sandbox, { filename: relPath });
  return sandbox.window;
}

test('GoogleSheetsService: extrae IDs válidos y parsea matrices con cabeceras desplazadas', () => {
  const win = loadBrowserScript('js/google-sheets.js');
  const svc = win.GoogleSheetsService;

  assert.equal(
    svc.extractSpreadsheetId('https://docs.google.com/spreadsheets/d/1AbCDefGhIjKlMnOpQrStUvWxYz_123456/edit#gid=0'),
    '1AbCDefGhIjKlMnOpQrStUvWxYz_123456'
  );
  assert.equal(svc.extractSpreadsheetId('1AbCDefGhIjKlMnOpQrStUvWxYz_123456'), '1AbCDefGhIjKlMnOpQrStUvWxYz_123456');
  assert.equal(svc.extractSpreadsheetId('corto'), null);

  const parsed = svc.parseSheetMatrix([
    ['Tabla6'],
    ['', ''],
    ['DNI', 'NOMBRE', 'DISTRITO', 'COSTO'],
    ['12345678', 'Ana Perez', 'Ate', 310],
    ['', '', '', ''],
    ['87654321', 'Luis Soto', 'Santa Anita', 290]
  ]);

  assert.deepEqual([...parsed.headers], ['DNI', 'NOMBRE', 'DISTRITO', 'COSTO']);
  assert.equal(parsed.rows.length, 2);
  assert.equal(parsed.rows[0].NOMBRE, 'Ana Perez');
  assert.equal(parsed.rows[1].COSTO, 290);
});

test('PivotEngine: calcula pivote y neutraliza payloads XSS en HTML renderizado', () => {
  const win = loadBrowserScript('js/pivot-engine.js');
  const engine = win.PivotEngine;

  const rows = [
    { Area: '<img src=x onerror=alert(1)>', Zona: 'Norte', Cajas: 100 },
    { Area: '<img src=x onerror=alert(1)>', Zona: 'Sur', Cajas: 50 },
    { Area: 'Secos', Zona: 'Norte', Cajas: 25 }
  ];

  const pivotModel = engine.buildPivotData(rows, {
    rowFields: ['Area'],
    colFields: ['Zona'],
    valField: 'Cajas',
    aggFunc: 'SUM'
  });

  const html = engine.renderTableHtml({
    id: 'tbl"<script>',
    type: 'pivot',
    ...pivotModel
  });

  assert.ok(!html.includes('<img src=x'), 'No debe inyectar HTML crudo desde datos');
  assert.ok(!html.includes('<script>'), 'No debe inyectar etiquetas script en el ID');
  assert.ok(html.includes('&lt;img src=x onerror=alert(1)&gt;'));
  assert.ok(html.includes('150'));
});

test('ClipboardUtil: protege contra CSV/TSV Formula Injection sin dañar números negativos', () => {
  const win = loadBrowserScript('js/clipboard.js');
  const clip = win.ClipboardUtil;

  assert.equal(clip.sanitizeSpreadsheetCell('=CMD("calc")'), '\'=CMD("calc")');
  assert.equal(clip.sanitizeSpreadsheetCell('+1+1'), '\'+1+1');
  assert.equal(clip.sanitizeSpreadsheetCell('@SUM(A1:A2)'), '\'@SUM(A1:A2)');
  assert.equal(clip.sanitizeSpreadsheetCell('-cmd /c calc'), '\'-cmd /c calc');
  // Los números negativos legítimos deben conservarse como números
  assert.equal(clip.sanitizeSpreadsheetCell('-12.5'), '-12.5');
  assert.equal(clip.sanitizeSpreadsheetCell('-15%'), '-15%');
  assert.equal(clip.sanitizeSpreadsheetCell('Normal'), 'Normal');
});

test('UrlStateManager: codifica y decodifica estado comprimido en hash', () => {
  const win = loadBrowserScript('js/url-state.js');
  const mgr = win.UrlStateManager;
  const original = { sheetId: 'abc12345678901234', tables: [{ id: 't1', mode: 'flat' }] };

  const hash = mgr.encodeState(original);
  assert.ok(hash.startsWith('#state='));

  win.location.hash = hash;
  const restored = mgr.decodeState();
  assert.deepEqual(restored, original);

  // Rechaza payloads que no sean objetos o superen el límite seguro de tamaño
  win.location.hash = '#state=' + encodeURIComponent(Buffer.from('["array"]', 'utf8').toString('base64'));
  assert.equal(mgr.decodeState(), null);

  win.location.hash = '#state=' + 'a'.repeat(100005);
  assert.equal(mgr.decodeState(), null);
});

test('Seguridad en HTML y configuración pública: CSP, Referrer-Policy, SRI y sin claves privadas expuestas', () => {
  const htmlFiles = [
    'index.html',
    'pages/Throughput/index.html',
    'pages/ant-inv/index.html',
    'pages/dias a tienda/index.html',
    'pages/dotacion/index.html',
    'pages/movilidad/index.html',
    'pages/lector/lector.html'
  ];

  for (const file of htmlFiles) {
    const content = fs.readFileSync(path.join(ROOT, file), 'utf8');
    assert.match(content, /http-equiv="Content-Security-Policy"/i, `${file} debe tener meta CSP`);
    const scriptSrcMatch = content.match(/script-src\s+([^;"]+)/i);
    assert.ok(scriptSrcMatch, `${file} debe declarar script-src en su CSP`);
    assert.ok(!scriptSrcMatch[1].includes("'unsafe-inline'"), `${file} no debe permitir 'unsafe-inline' en script-src`);
    assert.doesNotMatch(content, /<script(?![^>]*\bsrc=)[^>]*>/i, `${file} no debe tener etiquetas <script> inline`);
    assert.doesNotMatch(content, /\bon(?:click|change|input|keydown|keyup|keypress|submit)\s*=/i, `${file} no debe tener handlers inline on*=`);
    assert.match(content, /name="referrer"\s+content="strict-origin-when-cross-origin"/i, `${file} debe tener Referrer-Policy`);
    assert.doesNotMatch(content, /<(?:script|link|img|iframe|a)\b[^>]+(?:src|href)=["']http:\/\//i, `${file} no debe cargar recursos http:// inseguros`);

    // Verificar que ningún CDN versionado (cdnjs, jsdelivr, sheetjs, unpkg) carezca de SRI integrity
    const externalTags = [...content.matchAll(/<(?:script|link)\b[^>]+(?:cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|cdn\.sheetjs\.com|unpkg\.com)[^>]*>/gi)];
    for (const match of externalTags) {
      assert.match(match[0], /\bintegrity="sha384-[A-Za-z0-9+/=]+"/, `Falta integrity SRI en ${file}: ${match[0]}`);
      assert.match(match[0], /\bcrossorigin="anonymous"/, `Falta crossorigin="anonymous" en ${file}: ${match[0]}`);
    }
  }

  const uiJsFiles = [
    'js/app.js',
    'js/pivot-engine.js',
    'pages/Throughput/app.js',
    'pages/ant-inv/app.js',
    'pages/dias a tienda/app.js',
    'pages/dotacion/js/dotacion.js',
    'pages/movilidad/js/movilidad.js',
    'pages/lector/lector.js'
  ];
  for (const jsFile of uiJsFiles) {
    const jsContent = fs.readFileSync(path.join(ROOT, jsFile), 'utf8');
    assert.doesNotMatch(jsContent, /<[a-z][^>]*\bon(?:click|change|input|keydown|keyup|submit)\s*=/i, `${jsFile} no debe generar atributos inline on*=`);
  }

  const movilidadJs = fs.readFileSync(path.join(ROOT, 'pages/movilidad/js/movilidad.js'), 'utf8');
  assert.ok(!movilidadJs.includes('TARIFARIO_BD_BUSES'), 'movilidad.js no debe exponer tarifas hardcodeadas');
  assert.ok(!movilidadJs.includes('HORARIOS_OFICIALES_RUTAS'), 'movilidad.js no debe exponer horarios hardcodeados');

  const configContent = fs.readFileSync(path.join(ROOT, 'js/config.js'), 'utf8');
  assert.match(configContent, /GEMINI_API_KEY:\s*'TU_GEMINI_API_KEY_AQUI'/);
  assert.doesNotMatch(configContent, /AIza[0-9A-Za-z_-]{20,}/, 'config.js nunca debe contener claves privadas de Google/Gemini');
  assert.match(configContent, /DEFAULT_GOOGLE_CLIENT_ID:\s*'TU_CLIENT_ID_AQUI'/);
  assert.match(configContent, /MAPBOX_TOKEN:\s*'TU_MAPBOX_TOKEN_AQUI'/);
});
