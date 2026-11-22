/**
 * DataNexus - Antigüedad de Inventario (Secos 655 y Frescos 676)
 * Motor Ultra-Optimizado de Procesamiento por Lotes y Captura 16:9 Nativa
 * 
 * 1. Procesamiento ultra-rápido por lotes (5,000-8,000 filas/chunk) sin congelar la UI.
 * 2. Suma aritmética exacta de LPNs en todos los niveles.
 * 3. Gráfico de Barras con valores y porcentajes ENCIMA de cada barra.
 * 4. Tabla de resumen de rangos (Table 1) en Slide 1 sincronizada con tbBD.
 * 5. Captura 16:9 panorámica (1920x1080) directa al portapapeles (Ctrl + V en PPT) sin descargas forzadas.
 */

console.log('[DataNexus] app.js v20261102_week44_fix cargado exitosamente.');

// Estado global
let currentWarehouse = '655'; // '655' (Secos) o '676' (Frescos)
let currentSlide = 1;
let currentWeekLabel = 'Semana 37';
let currentChartS1Mode = 'bultos'; // 'bultos' o 'costo'
let chartInstanceS1 = null;
let chartInstanceS3 = null;
let isCapturing = false;
let lastCapturedBlob = null;
let lastCapturedDataUrl = null;

// Base de datos activa
let ACTIVE_DATABASE = {
  '655': null,
  '676': null
};

// ══════════════════════════════════════════════════════════════════════════════
// 1. UTILIDADES Y FORMATEADORES
// ══════════════════════════════════════════════════════════════════════════════

function safeFixed(val, decimals = 1, suffix = '') {
  if (val === null || val === undefined || isNaN(Number(val))) {
    return '0.0' + suffix;
  }
  return Number(val).toFixed(decimals) + suffix;
}

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function parseNum(v) {
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  if (!v) return 0;
  let str = String(v).trim().replace(/\s+/g, '');
  if (str.includes(',') && str.includes('.')) {
    if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      str = str.replace(/,/g, '');
    }
  } else if (str.includes(',')) {
    const parts = str.split(',');
    if (parts.length === 2 && parts[1].length !== 3) {
      str = str.replace(',', '.');
    } else if (parts.length > 2) {
      str = str.replace(/,/g, '');
    } else {
      str = str.replace(/,/g, '');
    }
  }
  const clean = str.replace(/[^\d\.-]/g, '').trim();
  const n = parseFloat(clean);
  return isNaN(n) ? 0 : n;
}

function formatCurrency(val) {
  const n = parseNum(val);
  return 'S/ ' + Math.round(n).toLocaleString('en-US');
}

function formatNumber(val) {
  const n = parseNum(val);
  return Math.round(n).toLocaleString('en-US');
}

function formatCompact(val) {
  const n = parseNum(val);
  if (!n) return '0';
  if (n >= 1000000) return (n / 1000000).toFixed(1) + 'M';
  if (n >= 1000) return (n / 1000).toFixed(0) + 'K';
  return Math.round(n).toString();
}

function showToast(message, type = 'info', duration = 3500) {
  const toast = document.getElementById('toastMessage');
  const toastText = document.getElementById('toastText');
  if (!toast || !toastText) return;

  toastText.textContent = message;
  toast.style.background = type === 'success' ? '#065f46' : (type === 'danger' ? '#991b1b' : '#0f172a');
  toast.classList.add('show');
  setTimeout(() => toast.classList.remove('show'), duration);
}

function showLoading(text, subtext = '') {
  const overlay = document.getElementById('loadingOverlay');
  const txt = document.getElementById('loadingText');
  const sub = document.getElementById('loadingSubtext');
  if (txt) txt.textContent = text || 'Procesando archivo Excel...';
  if (sub) sub.textContent = subtext;
  if (overlay) overlay.style.display = 'flex';
}

function hideLoading() {
  const overlay = document.getElementById('loadingOverlay');
  if (overlay) overlay.style.display = 'none';
}

function yieldToUi() {
  return new Promise(resolve => setTimeout(resolve, 0));
}

function normalizeRange(str) {
  if (!str) return '0 a 10 Semanas';
  const s = String(str).toLowerCase().trim();
  if (s.includes('0 a 10') || s.includes('0-10') || s.includes('< 10') || s.includes('menor a 10')) return '0 a 10 Semanas';
  if (s.includes('10 a 25') || s.includes('10-25')) return '10 a 25 Semanas';
  if (s.includes('25 a 52') || s.includes('25-52')) return '25 a 52 Semanas';
  if (s.includes('52') || s.includes('año') || s.includes('ano') || s.includes('> 52') || s.includes('mayor')) return 'mayor a 52 Semanas';
  return '0 a 10 Semanas';
}

// ══════════════════════════════════════════════════════════════════════════════
// 2. DETECTOR Y MAPA DE ÍNDICES DE COLUMNAS (ALTO RENDIMIENTO O(1))
// ══════════════════════════════════════════════════════════════════════════════

function buildColumnIndexMap(headers) {
  const normHeaders = headers.map(h => String(h || '').toLowerCase().replace(/[\s_\-\.\:\/%]/g, '').trim());

  function findIdx(candidates) {
    // 1. Coincidencia exacta limpia
    for (const cand of candidates) {
      const cleanCand = cand.toLowerCase().replace(/[\s_\-\.\:\/%]/g, '').trim();
      for (let i = 0; i < normHeaders.length; i++) {
        if (normHeaders[i] === cleanCand) return i;
      }
    }
    // 2. Coincidencia parcial si no hubo exacta
    for (const cand of candidates) {
      const cleanCand = cand.toLowerCase().replace(/[\s_\-\.\:\/%]/g, '').trim();
      for (let i = 0; i < normHeaders.length; i++) {
        if (normHeaders[i].includes(cleanCand)) return i;
      }
    }
    return -1;
  }

  function findIdxExact(candidates) {
    for (const cand of candidates) {
      const cleanCand = cand.toLowerCase().replace(/[\s_\-\.\:\/%]/g, '').trim();
      for (let i = 0; i < normHeaders.length; i++) {
        if (normHeaders[i] === cleanCand) return i;
      }
    }
    return -1;
  }

  return {
    whse: findIdx(['whse', 'almacenfisico', 'almacen', 'almacen_fisico']),
    costos: findIdx(['costos', 'costo', 'costototal', 'costo_total']),
    bultos: findIdx(['bultos', 'bulto', 'cajas', 'cantidad']),
    lpn: findIdx(['lpn', 'nrolpn', 'lpns', 'cantidadlpn', 'cantlpn']),
    sku: findIdx(['sku', 'codigo', 'codsku']),
    desc: findIdx(['descripcion', 'desc', 'descripcionsku', 'producto']),
    div: findIdx(['division', 'div', 'dpto', 'departamento']),
    semanas: findIdx(['semanas', 'rangosemanas', 'antiguedadsemanas', 'antiguedad']),
    zona: findIdx(['zona', 'tipozona', 'tipo_zona']),
    ubic: findIdx(['ubicacion', 'posicion', 'slot']),
    rngFv: findIdx(['rngfv', 'fechavencim', 'vencimiento', 'fechavencimiento']),
    onHand: findIdx(['onhand', 'unidades', 'on_hand']),
    dias: findIdx(['dias', 'diasantiguedad', 'antiguedaddias']),
    semana: findIdxExact(['semana', 'nrosemana', 'semanareporte', 'semreporte', 'semanacierre'])
  };
}

function findHeaderRowInMatrix(matrix) {
  const maxScan = Math.min(matrix.length, 12);
  for (let r = 0; r < maxScan; r++) {
    const row = matrix[r] || [];
    const text = row.map(c => String(c || '').toUpperCase()).join(' ');
    if ((text.includes('WHSE') || text.includes('ALMACEN')) &&
        (text.includes('SKU') || text.includes('COSTO') || text.includes('LPN') || text.includes('BULT'))) {
      return r;
    }
  }
  return 0;
}

// ══════════════════════════════════════════════════════════════════════════════
// 3. CARGA DE ARCHIVO Y DRAG & DROP
// ══════════════════════════════════════════════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
  setupDropzone();
});

function setupDropzone() {
  const dropzone = document.getElementById('dropzoneBox');
  if (!dropzone) return;

  ['dragenter', 'dragover'].forEach(name => {
    dropzone.addEventListener(name, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dragover');
    }, false);
  });

  ['dragleave', 'drop'].forEach(name => {
    dropzone.addEventListener(name, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');
    }, false);
  });

  dropzone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files.length > 0) {
      processExcelFile(dt.files[0]);
    }
  }, false);
}

function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (file) {
    processExcelFile(file);
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// 4. MOTOR PRINCIPAL DE LECTURA DE EXCEL (SHEETJS POR LOTES ASÍNCRONOS)
// ══════════════════════════════════════════════════════════════════════════════

function processExcelFile(file) {
  const MAX_FILE_BYTES = 50 * 1024 * 1024;
  if (!file || !/\.(xlsx|xls|xlsm|xlsb|csv)$/i.test(file.name || '')) {
    showToast('Formato no válido. Selecciona un archivo Excel (.xlsx, .xlsm, .xlsb, .xls) o CSV.', 'danger', 4500);
    return;
  }
  if (file.size > MAX_FILE_BYTES) {
    showToast('El archivo supera el límite seguro de 50 MB para el navegador.', 'danger', 4500);
    return;
  }

  showLoading(`Abriendo ${file.name}...`, 'Leyendo archivo binario...');

  const reader = new FileReader();

  reader.onload = async function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      showLoading('Inspeccionando hojas del libro...', 'Modo de alto rendimiento activo');
      await yieldToUi();

      // Lectura rápida de SheetJS sin strings pesados
      const workbook = XLSX.read(data, {
        type: 'array',
        cellDates: false,
        cellNF: false,
        cellText: false
      });

      // 1. Detectar Hoja BD
      let wsBD = null;
      let sheetNameBD = '';
      for (const sName of workbook.SheetNames) {
        const clean = sName.trim().toUpperCase();
        if (clean === 'BD' || clean === 'TBBD' || clean.includes('BD') || clean.includes('BASE')) {
          wsBD = workbook.Sheets[sName];
          sheetNameBD = sName;
          break;
        }
      }
      if (!wsBD) {
        sheetNameBD = workbook.SheetNames[0];
        wsBD = workbook.Sheets[sheetNameBD];
      }

      showLoading(`Estructurando matriz de datos de "${sheetNameBD}"...`);
      await yieldToUi();

      // Matriz 2D rápida: O(1) de memoria y sin miles de objetos
      const rawMatrix = XLSX.utils.sheet_to_json(wsBD, {
        header: 1,
        defval: ''
      });

      if (!rawMatrix || rawMatrix.length === 0) {
        throw new Error(`La hoja "${sheetNameBD}" no contiene registros.`);
      }

      const headerRowIdx = findHeaderRowInMatrix(rawMatrix);
      const headers = (rawMatrix[headerRowIdx] || []).map(h => String(h || '').trim());
      const colMap = buildColumnIndexMap(headers);

      // 2. Detectar Hoja EVOLUTIVO
      let evolSecos = null;
      let evolFrescos = null;
      const sheetNameEvol = workbook.SheetNames.find(n => {
        const clean = n.trim().toUpperCase();
        return clean.includes('EVOL') || clean.includes('HIST');
      });

      if (sheetNameEvol && workbook.Sheets[sheetNameEvol]) {
        showLoading('Extrayendo evolución de semanas desde hoja EVOLUTIVO...');
        await yieldToUi();
        const wsEvol = workbook.Sheets[sheetNameEvol];
        const evolFound = parseEvolutivoSheetSmart(wsEvol);
        evolSecos = evolFound.secos;
        evolFrescos = evolFound.frescos;
      }

      // 3. Detectar semana actual de los datos con máxima fidelidad
      let activeWeekNum = '';

      // Prioridad 1: Última semana con datos de la hoja EVOLUTIVO (fuente principal y oficial)
      const evolWeeksList = (evolSecos && evolSecos.weeks && evolSecos.weeks.length > 0)
        ? evolSecos.weeks
        : (evolFrescos && evolFrescos.weeks && evolFrescos.weeks.length > 0 ? evolFrescos.weeks : []);

      if (evolWeeksList.length > 0) {
        const lastWStr = evolWeeksList[evolWeeksList.length - 1];
        const numMatch = String(lastWStr || '').match(/\d+/);
        if (numMatch) activeWeekNum = numMatch[0];
      }

      // Prioridad 2: Si el nombre del archivo contiene un número de semana (ej: S-44, Semana 44, Sem 44)
      if (file && file.name) {
        const fileMatch = file.name.match(/(?:semana|sem|sem\.|s)[\s\-_]*(\d{1,2})/i);
        if (fileMatch && fileMatch[1]) {
          const fileWeek = fileMatch[1];
          if (!activeWeekNum || Number(fileWeek) >= Number(activeWeekNum)) {
            activeWeekNum = fileWeek;
          }
        }
      }

      // Prioridad 3: Columna explícita en tbBD (SOLO si no es la columna de antigüedad/rango y es constante)
      if (!activeWeekNum && colMap.semana !== -1 && colMap.semana !== colMap.semanas) {
        const sampleRows = rawMatrix.slice(headerRowIdx + 1, Math.min(headerRowIdx + 30, rawMatrix.length));
        const valSet = new Set(sampleRows.map(r => String(r[colMap.semana] || '').replace(/[^\d]/g, '').trim()).filter(Boolean));
        if (valSet.size === 1) {
          activeWeekNum = [...valSet][0];
        }
      }

      // Fallback seguro
      if (!activeWeekNum) activeWeekNum = '44';

      currentWeekLabel = `Semana ${activeWeekNum}`;
      const activeWeekTag = 'S-' + activeWeekNum;

      // 4. Procesar filas por bloques asíncronos para evitar congelamientos
      showLoading('Procesando inventario en lotes optimizados...', 'Calculando Secos (655) y Frescos (676)...');
      await yieldToUi();

      const dataRows = rawMatrix.slice(headerRowIdx + 1);
      const compiledData = await compileDataOptimized(dataRows, colMap, evolSecos, evolFrescos, activeWeekTag);

      ACTIVE_DATABASE['655'] = compiledData.secos;
      ACTIVE_DATABASE['676'] = compiledData.frescos;

      // 5. Actualizar interfaz
      document.getElementById('dropzoneBox').style.display = 'none';
      document.getElementById('fileStatusBar').style.display = 'flex';
      document.getElementById('loadedFileName').textContent = `Archivo: ${file.name}`;
      document.getElementById('loadedFileMeta').textContent = `${dataRows.length.toLocaleString('en-US')} filas procesadas | CD Secos 655 y Frescos 676 actualizados`;

      updateHeaderWeekBadges();
      renderWarehouseData(currentWarehouse);
      hideLoading();
      showToast(`¡${dataRows.length.toLocaleString('en-US')} filas procesadas con éxito!`, 'success', 4000);

    } catch (err) {
      console.error('Error al procesar Excel:', err);
      hideLoading();
      alert(`Error al procesar el archivo Excel: ${err.message}`);
    }
  };

  reader.onerror = function () {
    hideLoading();
    showToast('Error al leer el archivo desde el disco', 'danger');
  };

  reader.readAsArrayBuffer(file);
}

// ══════════════════════════════════════════════════════════════════════════════
// 5. EXTRACCIÓN DE LA HOJA EVOLUTIVO
// ══════════════════════════════════════════════════════════════════════════════

function parseEvolutivoSheetSmart(ws) {
  const result = { secos: null, frescos: null };
  if (!ws || !ws['!ref']) return result;

  const range = XLSX.utils.decode_range(ws['!ref']);
  const matchingRows = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= Math.min(range.e.c, 6); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v !== undefined) {
        const txt = String(cell.v).toLowerCase();
        if (txt.includes('0 a 10') || txt.includes('0-10')) {
          matchingRows.push({ rowDataStart: r, colLabel: c });
          break;
        }
      }
    }
  }

  if (matchingRows.length > 0) {
    result.secos = extractEvolutivoBlock(ws, matchingRows[0].rowDataStart, range);
  } else {
    result.secos = extractEvolutivoBlock(ws, 5, range);
  }

  if (matchingRows.length > 1) {
    result.frescos = extractEvolutivoBlock(ws, matchingRows[1].rowDataStart, range);
  } else {
    result.frescos = extractEvolutivoBlock(ws, 14, range);
  }

  return result;
}

function extractEvolutivoBlock(ws, dataStartRow, range) {
  const headerRow = Math.max(0, dataStartRow - 1);
  const validCols = [];

  for (let c = 1; c <= range.e.c; c++) {
    const hCell = ws[XLSX.utils.encode_cell({ r: headerRow, c })];
    if (!hCell || hCell.v === undefined || String(hCell.v).trim() === '') continue;

    let hasData = false;
    for (let r = 0; r < 4; r++) {
      const dCell = ws[XLSX.utils.encode_cell({ r: dataStartRow + r, c })];
      if (dCell && dCell.v !== undefined && String(dCell.v).trim() !== '') {
        hasData = true;
        break;
      }
    }

    if (hasData) {
      validCols.push(c);
    }
  }

  if (validCols.length === 0) return null;

  const weeks = validCols.map(c => {
    const cell = ws[XLSX.utils.encode_cell({ r: headerRow, c })];
    let txt = String(cell ? cell.v : '').trim();
    if (!txt.toUpperCase().startsWith('S-') && !txt.toUpperCase().startsWith('SEM')) {
      txt = 'S-' + txt;
    }
    return txt;
  });

  const rangeDefs = [
    { label: '0 a 10 Semanas', color: '#10b981' },
    { label: '10 a 25 Semanas', color: '#f59e0b' },
    { label: '25 a 52 Semanas', color: '#f97316' },
    { label: 'Mayor a 52 Semanas', color: '#ef4444' }
  ];

  const evolution = rangeDefs.map((rDef, idx) => {
    const r = dataStartRow + idx;
    const values = validCols.map(c => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v === undefined) return 0;
      
      let val = 0;
      if (typeof cell.v === 'number') {
        val = cell.v <= 1.0 && cell.v > 0 ? cell.v * 100 : cell.v;
      } else if (cell.w && typeof cell.w === 'string' && cell.w.includes('%')) {
        val = parseFloat(cell.w.replace('%', '').replace(',', '.').trim()) || 0;
      } else {
        const clean = String(cell.v || '').replace('%', '').replace(',', '.').trim();
        val = parseFloat(clean) || 0;
        if (val <= 1.0 && val > 0) val = val * 100;
      }
      return isNaN(val) ? 0 : Number(val.toFixed(5));
    });
    return {
      label: rDef.label,
      values,
      color: rDef.color
    };
  });

  return { weeks, evolution };
}

// ══════════════════════════════════════════════════════════════════════════════
// 6. MOTOR POR LOTES ASÍNCRONOS PARA EVITAR CONGELAMIENTOS
// ══════════════════════════════════════════════════════════════════════════════

function initAccumulator(whseCode, whseLabel) {
  return {
    whseCode,
    whseLabel,
    totalCost: 0,
    totalBultos: 0,
    totalLpns: 0,
    rangeAgg: {
      '0 a 10 Semanas': { cost: 0, bultos: 0, lpns: 0, color: '#10b981', status: 'Saludable' },
      '10 a 25 Semanas': { cost: 0, bultos: 0, lpns: 0, color: '#f59e0b', status: 'En Alerta' },
      '25 a 52 Semanas': { cost: 0, bultos: 0, lpns: 0, color: '#f97316', status: 'Riesgo Medio' },
      'mayor a 52 Semanas': { cost: 0, bultos: 0, lpns: 0, color: '#ef4444', status: 'Crítico >1 año' }
    },
    divisionAgg: {},
    zoneTableAgg: {},
    skusOver52: {},
    locationsAgg: {}
  };
}

function accumulateRow(row, colMap, acc) {
  const cost = colMap.costos !== -1 ? parseNum(row[colMap.costos]) : 0;
  const bultos = colMap.bultos !== -1 ? parseNum(row[colMap.bultos]) : 0;
  const onHand = colMap.onHand !== -1 ? parseNum(row[colMap.onHand]) : 0;
  const dias = colMap.dias !== -1 ? parseNum(row[colMap.dias]) : 0;
  const lpnVal = 1; // Cada registro en tbBD representa exactamente 1 pallet / 1 LPN físico

  const sku = colMap.sku !== -1 ? String(row[colMap.sku] || '').trim() : '';
  const desc = colMap.desc !== -1 ? String(row[colMap.desc] || '').trim() : '';
  const div = colMap.div !== -1 ? (String(row[colMap.div] || 'OTROS').trim() || 'OTROS') : 'OTROS';
  const rng = colMap.semanas !== -1 ? normalizeRange(row[colMap.semanas]) : '0 a 10 Semanas';
  const zona = colMap.zona !== -1 ? String(row[colMap.zona] || '').toUpperCase().trim() : '';
  const ubic = colMap.ubic !== -1 ? String(row[colMap.ubic] || '').trim() : '';
  const rngFv = colMap.rngFv !== -1 ? (String(row[colMap.rngFv] || '-').trim() || '-') : '-';

  acc.totalCost += cost;
  acc.totalBultos += bultos;
  acc.totalLpns += lpnVal;

  // 1. Rangos
  if (acc.rangeAgg[rng]) {
    acc.rangeAgg[rng].cost += cost;
    acc.rangeAgg[rng].bultos += bultos;
    acc.rangeAgg[rng].lpns += lpnVal;
  }

  // 2. Divisiones
  if (!acc.divisionAgg[div]) {
    acc.divisionAgg[div] = { code: div, r010: 0, r1025: 0, r2552: 0, r52: 0, total: 0 };
  }
  acc.divisionAgg[div].total += cost;
  if (rng === '0 a 10 Semanas') acc.divisionAgg[div].r010 += cost;
  else if (rng === '10 a 25 Semanas') acc.divisionAgg[div].r1025 += cost;
  else if (rng === '25 a 52 Semanas') acc.divisionAgg[div].r2552 += cost;
  else if (rng === 'mayor a 52 Semanas') acc.divisionAgg[div].r52 += cost;

  // 3. Zonas (Matriz Completa para Pivot Table)
  let zoneCode = zona;
  if (!zoneCode) {
    zoneCode = 'OTROS';
  } else if (zoneCode.includes('RCK') || zoneCode.includes('RACK')) {
    zoneCode = 'RCK';
  } else if (zoneCode.includes('RHB') || zoneCode.includes('HIGHBAY')) {
    zoneCode = 'RHB';
  } else if (zoneCode.includes('TOL')) {
    zoneCode = 'TOL';
  }

  if (!acc.zoneTableAgg[zoneCode]) {
    acc.zoneTableAgg[zoneCode] = {
      name: zoneCode,
      r010: { cost: 0, lpns: 0, bultos: 0 },
      r1025: { cost: 0, lpns: 0, bultos: 0 },
      r2552: { cost: 0, lpns: 0, bultos: 0 },
      r52: { cost: 0, lpns: 0, bultos: 0 },
      totalCost: 0,
      totalLpns: 0,
      totalBultos: 0,
      skus52: {}
    };
  }

  const zItem = acc.zoneTableAgg[zoneCode];
  zItem.totalCost += cost;
  zItem.totalLpns += lpnVal;
  zItem.totalBultos += bultos;

  let rTarget = null;
  if (rng === '0 a 10 Semanas') rTarget = zItem.r010;
  else if (rng === '10 a 25 Semanas') rTarget = zItem.r1025;
  else if (rng === '25 a 52 Semanas') rTarget = zItem.r2552;
  else if (rng === 'mayor a 52 Semanas') rTarget = zItem.r52;

  if (rTarget) {
    rTarget.cost += cost;
    rTarget.lpns += lpnVal;
    rTarget.bultos += bultos;
  }

  if (rng === 'mayor a 52 Semanas' && sku) {
    if (!zItem.skus52[sku]) {
      zItem.skus52[sku] = { sku, desc, fv: rngFv, lpns: 0, cost: 0, bultos: 0 };
    }
    zItem.skus52[sku].cost += cost;
    zItem.skus52[sku].bultos += bultos;
    zItem.skus52[sku].lpns += lpnVal;
  }

  // 4. Top SKUs >52 Semanas
  if (rng === 'mayor a 52 Semanas' && sku) {
    if (!acc.skusOver52[sku]) {
      acc.skusOver52[sku] = { sku, desc, div, rngFv, lpns: 0, cost: 0, bultos: 0, onHand: 0, diasTotal: 0, diasCount: 0 };
    }
    acc.skusOver52[sku].cost += cost;
    acc.skusOver52[sku].bultos += bultos;
    acc.skusOver52[sku].onHand += onHand;
    acc.skusOver52[sku].lpns += lpnVal;
    acc.skusOver52[sku].diasTotal += dias;
    acc.skusOver52[sku].diasCount++;
  }

  // 5. Ubicaciones
  if (ubic) {
    if (!acc.locationsAgg[div]) {
      acc.locationsAgg[div] = { div, r010Ubic: new Set(), r1025Ubic: new Set(), r2552Ubic: new Set(), r52Ubic: new Set(), totalUbic: new Set() };
    }
    acc.locationsAgg[div].totalUbic.add(ubic);
    if (rng === '0 a 10 Semanas') acc.locationsAgg[div].r010Ubic.add(ubic);
    else if (rng === '10 a 25 Semanas') acc.locationsAgg[div].r1025Ubic.add(ubic);
    else if (rng === '25 a 52 Semanas') acc.locationsAgg[div].r2552Ubic.add(ubic);
    else if (rng === 'mayor a 52 Semanas') acc.locationsAgg[div].r52Ubic.add(ubic);
  }
}

async function compileDataOptimized(dataRows, colMap, evolSecos, evolFrescos, activeWeekTag) {
  const CHUNK_SIZE = 8000;
  const totalRows = dataRows.length;

  const acc655 = initAccumulator('655', 'CD Secos 655');
  const acc676 = initAccumulator('676', 'CD Frescos 676');

  for (let i = 0; i < totalRows; i += CHUNK_SIZE) {
    const end = Math.min(i + CHUNK_SIZE, totalRows);

    for (let r = i; r < end; r++) {
      const row = dataRows[r];
      if (!row || row.length === 0) continue;

      const rawWhse = colMap.whse !== -1 ? String(row[colMap.whse] || '').trim() : '';
      let targetAcc = null;
      if (rawWhse === '655' || rawWhse.toLowerCase().includes('seco')) targetAcc = acc655;
      else if (rawWhse === '676' || rawWhse.toLowerCase().includes('fresco')) targetAcc = acc676;
      else continue;

      accumulateRow(row, colMap, targetAcc);
    }

    const pct = Math.round((end / totalRows) * 100);
    showLoading(`Procesando filas por lotes optimizados...`, `Fila ${end.toLocaleString()} de ${totalRows.toLocaleString()} (${pct}%)`);
    await yieldToUi();
  }

  const secos = finalizeWarehouseData(acc655, evolSecos, activeWeekTag);
  const frescos = finalizeWarehouseData(acc676, evolFrescos, activeWeekTag);

  return { secos, frescos };
}

function finalizeWarehouseData(acc, evolObj, activeWeekTag) {
  const { whseCode, whseLabel, totalCost, totalBultos, totalLpns, rangeAgg, divisionAgg, zoneAgg, skusOver52, locationsAgg } = acc;

  if (totalCost === 0 && totalBultos === 0) {
    return createEmptyWarehouseData(whseCode, whseLabel, evolObj, activeWeekTag);
  }

  // Estructuración de Rangos con % BULTOS, % COSTO y % LPN con precisión matemática exacta (5 decimales)
  const ranges = Object.keys(rangeAgg).map(k => {
    const itm = rangeAgg[k];
    const bultosPct = totalBultos > 0 ? Number(((itm.bultos / totalBultos) * 100).toFixed(5)) : 0;
    const costPct = totalCost > 0 ? Number(((itm.cost / totalCost) * 100).toFixed(5)) : 0;
    const lpnsPct = totalLpns > 0 ? Number(((itm.lpns / totalLpns) * 100).toFixed(5)) : 0;

    return {
      label: k,
      cost: itm.cost,
      costPct: costPct,
      lpns: itm.lpns,
      lpnsPct: lpnsPct,
      bultos: itm.bultos,
      bultosPct: bultosPct,
      color: itm.color,
      status: itm.status
    };
  });

  // Divisiones
  const divisions = Object.values(divisionAgg)
    .sort((a, b) => b.total - a.total)
    .map(d => ({
      code: d.code,
      r010: d.r010,
      r1025: d.r1025,
      r2552: d.r2552,
      r52: d.r52,
      total: d.total,
      pct: totalCost > 0 ? parseFloat(((d.total / totalCost) * 100).toFixed(1)) : 0
    }));

  // Top 10 SKUs >52s con Suma de LPNs
  const top10Skus = Object.values(skusOver52)
    .sort((a, b) => b.cost - a.cost)
    .slice(0, 10)
    .map((s, idx) => ({
      sku: s.sku,
      desc: s.desc,
      div: s.div,
      rngFv: s.rngFv,
      lpns: s.lpns,
      cost: Math.round(s.cost),
      bultos: s.bultos,
      onHand: s.onHand,
      days: s.diasCount > 0 ? Math.round(s.diasTotal / s.diasCount) : 0,
      badge: idx === 0 ? 'CRÍTICO #1' : (idx === 1 ? 'CRÍTICO #2' : (idx < 5 ? 'PRIORIDAD' : 'SALDO'))
    }));

  // 1. Matriz Completa de Zonas para el Pivot de la Diapositiva 2
  const zoneList = Object.values(acc.zoneTableAgg || {}).sort((a, b) => {
    const order = { 'RCK': 1, 'RHB': 2, 'TOL': 3 };
    const ordA = order[a.name] || 99;
    const ordB = order[b.name] || 99;
    if (ordA !== ordB) return ordA - ordB;
    return b.totalCost - a.totalCost;
  });

  const zoneTable = zoneList.map(z => ({
    name: z.name,
    costos: [Math.round(z.r010.cost), Math.round(z.r1025.cost), Math.round(z.r2552.cost), Math.round(z.r52.cost), Math.round(z.totalCost)],
    lpn: [z.r010.lpns, z.r1025.lpns, z.r2552.lpns, z.r52.lpns, z.totalLpns],
    bultos: [Math.round(z.r010.bultos), Math.round(z.r1025.bultos), Math.round(z.r2552.bultos), Math.round(z.r52.bultos), Math.round(z.totalBultos)]
  }));

  const zoneTotals = {
    costos: [
      Math.round(zoneList.reduce((sum, z) => sum + z.r010.cost, 0)),
      Math.round(zoneList.reduce((sum, z) => sum + z.r1025.cost, 0)),
      Math.round(zoneList.reduce((sum, z) => sum + z.r2552.cost, 0)),
      Math.round(zoneList.reduce((sum, z) => sum + z.r52.cost, 0)),
      Math.round(zoneList.reduce((sum, z) => sum + z.totalCost, 0))
    ],
    lpn: [
      zoneList.reduce((sum, z) => sum + z.r010.lpns, 0),
      zoneList.reduce((sum, z) => sum + z.r1025.lpns, 0),
      zoneList.reduce((sum, z) => sum + z.r2552.lpns, 0),
      zoneList.reduce((sum, z) => sum + z.r52.lpns, 0),
      zoneList.reduce((sum, z) => sum + z.totalLpns, 0)
    ],
    bultos: [
      Math.round(zoneList.reduce((sum, z) => sum + z.r010.bultos, 0)),
      Math.round(zoneList.reduce((sum, z) => sum + z.r1025.bultos, 0)),
      Math.round(zoneList.reduce((sum, z) => sum + z.r2552.bultos, 0)),
      Math.round(zoneList.reduce((sum, z) => sum + z.r52.bultos, 0)),
      Math.round(zoneList.reduce((sum, z) => sum + z.totalBultos, 0))
    ]
  };

  // Top SKUs >52s organizados por zona dinámica (RCK, RHB, TOL o cualquier otra)
  const topSkusByZone = {};
  zoneList.forEach(z => {
    const skusList = Object.values(z.skus52 || {})
      .sort((a, b) => b.cost - a.cost)
      .slice(0, 10)
      .map(s => ({
        sku: s.sku,
        desc: s.desc,
        fv: s.fv,
        lpns: s.lpns,
        cost: Math.round(s.cost),
        bultos: s.bultos
      }));

    topSkusByZone[z.name] = {
      zone: z.name,
      skus: skusList,
      hasData: skusList.length > 0,
      totalCost: skusList.reduce((sum, s) => sum + s.cost, 0),
      totalLpns: skusList.reduce((sum, s) => sum + s.lpns, 0),
      totalBultos: skusList.reduce((sum, s) => sum + s.bultos, 0)
    };
  });

  const rckItem = acc.zoneTableAgg['RCK'] || { skus52: {}, totalCost: 0, totalLpns: 0, totalBultos: 0, r010: { cost: 0 }, r1025: { cost: 0 }, r2552: { cost: 0 }, r52: { cost: 0 } };
  const rhbItem = acc.zoneTableAgg['RHB'] || { skus52: {}, totalCost: 0, totalLpns: 0, totalBultos: 0, r010: { cost: 0 }, r1025: { cost: 0 }, r2552: { cost: 0 }, r52: { cost: 0 } };

  const topSkusRck = topSkusByZone['RCK'] ? topSkusByZone['RCK'].skus : [];
  const topSkusRhb = topSkusByZone['RHB'] ? topSkusByZone['RHB'].skus : [];

  const zones = {
    rck: {
      totalCost: rckItem.totalCost || 0,
      lpns: rckItem.totalLpns || 0,
      bultos: rckItem.totalBultos || 0,
      pct: totalCost > 0 ? parseFloat((((rckItem.totalCost || 0) / totalCost) * 100).toFixed(1)) : 0,
      r010: rckItem.r010?.cost || 0,
      r1025: rckItem.r1025?.cost || 0,
      r2552: rckItem.r2552?.cost || 0,
      r52: rckItem.r52?.cost || 0,
      lpns52: Object.values(rckItem.skus52 || {}).reduce((sum, item) => sum + item.lpns, 0),
      topSkus: topSkusRck
    },
    rhb: {
      totalCost: rhbItem.totalCost || 0,
      lpns: rhbItem.totalLpns || 0,
      bultos: rhbItem.totalBultos || 0,
      pct: totalCost > 0 ? parseFloat((((rhbItem.totalCost || 0) / totalCost) * 100).toFixed(1)) : 0,
      r010: rhbItem.r010?.cost || 0,
      r1025: rhbItem.r1025?.cost || 0,
      r2552: rhbItem.r2552?.cost || 0,
      r52: rhbItem.r52?.cost || 0,
      lpns52: Object.values(rhbItem.skus52 || {}).reduce((sum, item) => sum + item.lpns, 0),
      topSkus: topSkusRhb
    }
  };

  // Ubicaciones
  let locationsTotal = 0;
  Object.values(locationsAgg).forEach(l => locationsTotal += l.totalUbic.size);

  const locations = Object.values(locationsAgg)
    .sort((a, b) => b.totalUbic.size - a.totalUbic.size)
    .map(l => ({
      div: l.div,
      r010: l.r010Ubic.size,
      r1025: l.r1025Ubic.size,
      r2552: l.r2552Ubic.size,
      r52: l.r52Ubic.size,
      total: l.totalUbic.size,
      pct: locationsTotal > 0 ? parseFloat(((l.totalUbic.size / locationsTotal) * 100).toFixed(1)) : 0
    }));

  // Sincronización de 7 Semanas del Evolutivo
  const currentWeekBultosPcts = {
    '0 a 10 Semanas': ranges[0].bultosPct,
    '10 a 25 Semanas': ranges[1].bultosPct,
    '25 a 52 Semanas': ranges[2].bultosPct,
    'mayor a 52 Semanas': ranges[3].bultosPct
  };

  let finalWeeks = [];
  let finalEvolution = [];
  const defaultRangeColors = {
    '0 a 10 Semanas': '#10b981',
    '10 a 25 Semanas': '#f59e0b',
    '25 a 52 Semanas': '#f97316',
    'Mayor a 52 Semanas': '#ef4444'
  };

  if (evolObj && evolObj.weeks && evolObj.weeks.length > 0) {
    const evolWeeks = [...evolObj.weeks];
    const lastWeekInSheet = evolWeeks[evolWeeks.length - 1];

    const lastWNum = String(lastWeekInSheet || '').replace(/[^\d]/g, '').trim();
    const activeWNum = String(activeWeekTag || '').replace(/[^\d]/g, '').trim();

    // Si la última semana de la hoja es la activa O la hoja ya contiene la semana activa en su serie
    if (lastWNum === activeWNum || evolWeeks.some(w => String(w).replace(/[^\d]/g, '').trim() === activeWNum)) {
      finalWeeks = evolWeeks.slice(-12);
      finalEvolution = evolObj.evolution.map(e => {
        const vals = [...(e.values || [])].slice(-12);
        const normLabel = normalizeRange(e.label);
        if (currentWeekBultosPcts[normLabel] !== undefined && vals.length > 0) {
          vals[vals.length - 1] = currentWeekBultosPcts[normLabel];
        }
        return {
          label: e.label,
          values: vals,
          color: e.color || defaultRangeColors[e.label] || '#2563eb'
        };
      });
    } else {
      const hist11Weeks = evolWeeks.slice(-11);
      finalWeeks = [...hist11Weeks, activeWeekTag];

      finalEvolution = evolObj.evolution.map(e => {
        const hist11Vals = (e.values || []).slice(-11);
        const normLabel = normalizeRange(e.label);
        const currentPct = currentWeekBultosPcts[normLabel] !== undefined ? currentWeekBultosPcts[normLabel] : 0;
        return {
          label: e.label,
          values: [...hist11Vals, currentPct],
          color: e.color || defaultRangeColors[e.label] || '#2563eb'
        };
      });
    }
  } else {
    finalWeeks = ['S-26', 'S-27', 'S-28', 'S-29', 'S-30', 'S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36', activeWeekTag];
    finalEvolution = [
      { label: '0 a 10 Semanas', values: [90.10, 89.80, 89.50, 89.20, 88.90, 89.14, 87.07, 85.83, 86.07, 87.09, 87.96, currentWeekBultosPcts['0 a 10 Semanas']], color: '#10b981' },
      { label: '10 a 25 Semanas', values: [8.10, 8.30, 8.50, 8.60, 8.80, 8.73, 10.68, 11.83, 11.71, 10.74, 10.18, currentWeekBultosPcts['10 a 25 Semanas']], color: '#f59e0b' },
      { label: '25 a 52 Semanas', values: [1.50, 1.60, 1.70, 1.80, 1.90, 1.83, 1.95, 2.03, 1.94, 1.91, 1.60, currentWeekBultosPcts['25 a 52 Semanas']], color: '#f97316' },
      { label: 'Mayor a 52 Semanas', values: [0.30, 0.30, 0.30, 0.40, 0.40, 0.30, 0.31, 0.31, 0.28, 0.26, 0.27, currentWeekBultosPcts['mayor a 52 Semanas']], color: '#ef4444' }
    ];
  }

  return {
    name: whseLabel,
    whseCode,
    badgeText: whseCode === '655' ? 'CD SECOS 655' : 'CD FRESCOS 676',
    totalCost: Math.round(totalCost),
    totalLpns,
    totalBultos: Math.round(totalBultos),
    ranges,
    divisions,
    top10Skus,
    zoneTable,
    zoneTotals,
    topSkusByZone,
    topSkusRck,
    topSkusRhb,
    zones,
    weeks: finalWeeks,
    evolution: finalEvolution,
    locationsTotal,
    locations
  };
}

function createEmptyWarehouseData(whseTarget, whseLabel, evolObj, activeWeekTag = 'S-37') {
  const weeks = evolObj ? evolObj.weeks : ['S-26', 'S-27', 'S-28', 'S-29', 'S-30', 'S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36', activeWeekTag];
  const evolution = evolObj ? evolObj.evolution : [
    { label: '0 a 10 Semanas', values: Array(12).fill(0), color: '#10b981' },
    { label: '10 a 25 Semanas', values: Array(12).fill(0), color: '#f59e0b' },
    { label: '25 a 52 Semanas', values: Array(12).fill(0), color: '#f97316' },
    { label: 'Mayor a 52 Semanas', values: Array(12).fill(0), color: '#ef4444' }
  ];

  return {
    name: whseLabel,
    whseCode: whseTarget,
    badgeText: whseTarget === '655' ? 'CD SECOS 655' : 'CD FRESCOS 676',
    totalCost: 0,
    totalLpns: 0,
    totalBultos: 0,
    ranges: [
      { label: '0 a 10 Semanas', cost: 0, costPct: 0, lpns: 0, lpnsPct: 0, bultos: 0, bultosPct: 0, color: '#10b981', status: 'Saludable' },
      { label: '10 a 25 Semanas', cost: 0, costPct: 0, lpns: 0, lpnsPct: 0, bultos: 0, bultosPct: 0, color: '#f59e0b', status: 'En Alerta' },
      { label: '25 a 52 Semanas', cost: 0, costPct: 0, lpns: 0, lpnsPct: 0, bultos: 0, bultosPct: 0, color: '#f97316', status: 'Riesgo Medio' },
      { label: 'mayor a 52 Semanas', cost: 0, costPct: 0, lpns: 0, lpnsPct: 0, bultos: 0, bultosPct: 0, color: '#ef4444', status: 'Crítico >1 año' }
    ],
    divisions: [],
    top10Skus: [],
    zoneTable: [],
    zoneTotals: { costos: [0, 0, 0, 0, 0], lpn: [0, 0, 0, 0, 0], bultos: [0, 0, 0, 0, 0] },
    topSkusByZone: {},
    topSkusRck: [],
    topSkusRhb: [],
    zones: {
      rck: { totalCost: 0, lpns: 0, bultos: 0, pct: 0, r010: 0, r1025: 0, r2552: 0, r52: 0, lpns52: 0, topSkus: [] },
      rhb: { totalCost: 0, lpns: 0, bultos: 0, pct: 0, r010: 0, r1025: 0, r2552: 0, r52: 0, lpns52: 0, topSkus: [] }
    },
    weeks,
    evolution,
    locationsTotal: 0,
    locations: []
  };
}

// ══════════════════════════════════════════════════════════════════════════════
// 7. DATOS DEMOSTRATIVOS (SEMANA 37)
// ══════════════════════════════════════════════════════════════════════════════

function loadSampleData() {
  showLoading('Cargando datos demostrativos de prueba...');
  currentWeekLabel = 'Semana 37 (Demo)';

  // CD Secos 655 (Datos Sintéticos de Demostración)
  ACTIVE_DATABASE['655'] = {
    name: 'Centro de Distribución - Secos (Demo)',
    whseCode: '655',
    badgeText: 'CD SECOS 655',
    totalCost: 25000000,
    totalLpns: 12500,
    totalBultos: 450000,
    ranges: [
      { label: '0 a 10 Semanas', cost: 21500000, costPct: 86.00, lpns: 10500, lpnsPct: 84.00, bultos: 391500, bultosPct: 87.00, color: '#10b981', status: 'Saludable' },
      { label: '10 a 25 Semanas', cost: 2750000, costPct: 11.00, lpns: 1625, lpnsPct: 13.00, bultos: 47250, bultosPct: 10.50, color: '#f59e0b', status: 'En Alerta' },
      { label: '25 a 52 Semanas', cost: 600000, costPct: 2.40, lpns: 325, lpnsPct: 2.60, bultos: 9000, bultosPct: 2.00, color: '#f97316', status: 'Riesgo Medio' },
      { label: 'mayor a 52 Semanas', cost: 150000, costPct: 0.60, lpns: 50, lpnsPct: 0.40, bultos: 2250, bultosPct: 0.50, color: '#ef4444', status: 'Crítico >1 año' }
    ],
    divisions: [
      { code: 'J01-PGC COMESTIBLE', r010: 9500000, r1025: 800000, r2552: 150000, r52: 50000, total: 10500000, pct: 42.0 },
      { code: 'J08-VESTUARIO', r010: 3000000, r1025: 500000, r2552: 100000, r52: 0, total: 3600000, pct: 14.4 },
      { code: 'J09-HOGAR', r010: 2600000, r1025: 450000, r2552: 140000, r52: 60000, total: 3250000, pct: 13.0 },
      { code: 'J02-PGC NO COMESTIBLE', r010: 2500000, r1025: 420000, r2552: 70000, r52: 10000, total: 3000000, pct: 12.0 },
      { code: 'J11-ELECTROHOGAR', r010: 2400000, r1025: 300000, r2552: 50000, r52: 0, total: 2750000, pct: 11.0 },
      { code: 'J10-BAZAR', r010: 600000, r1025: 200000, r2552: 70000, r52: 15000, total: 885000, pct: 3.5 },
      { code: 'J06-PANADERIA Y PASTELERIA', r010: 350000, r1025: 40000, r2552: 10000, r52: 5000, total: 405000, pct: 1.6 },
      { code: 'J12-INSTITUCIONALES', r010: 250000, r1025: 20000, r2552: 5000, r52: 5000, total: 280000, pct: 1.1 },
      { code: 'J05-FLC', r010: 200000, r1025: 10000, r2552: 0, r52: 0, total: 210000, pct: 0.8 },
      { code: 'J07-PLATOS PREPARADOS', r010: 100000, r1025: 10000, r2552: 5000, r52: 5000, total: 120000, pct: 0.5 }
    ],
    top10Skus: [
      { sku: '10001001', desc: 'PRODUCTO DEMO CONSERVAS 400G (EJEMPLO)', div: 'J01-PGC COMESTIBLE', rngFv: '20/07/2027', lpns: 12, cost: 32000, bultos: 650, onHand: 7800, days: 580, badge: 'CRÍTICO #1' },
      { sku: '10001002', desc: 'BEBIDA DEMO SELECCION 750ML (EJEMPLO)', div: 'J01-PGC COMESTIBLE', rngFv: '10/10/2028', lpns: 5, cost: 24500, bultos: 280, onHand: 1680, days: 540, badge: 'CRÍTICO #2' },
      { sku: '10001003', desc: 'SET MENAJE COCINA DEMO X 4 UN', div: 'J09-HOGAR', rngFv: '-', lpns: 7, cost: 19800, bultos: 190, onHand: 190, days: 510, badge: 'PRIORIDAD' },
      { sku: '10001004', desc: 'SALSA CONDIMENTO DEMO X 400G', div: 'J01-PGC COMESTIBLE', rngFv: '30/09/2027', lpns: 6, cost: 16400, bultos: 420, onHand: 5040, days: 460, badge: 'PRIORIDAD' },
      { sku: '10001005', desc: 'ACCESORIO HOGAR TEXTIL DEMO 2025', div: 'J09-HOGAR', rngFv: '-', lpns: 5, cost: 14200, bultos: 110, onHand: 110, days: 495, badge: 'PRIORIDAD' },
      { sku: '10001006', desc: 'ADEREZO CLASICO DEMO X 425G', div: 'J01-PGC COMESTIBLE', rngFv: '19/05/2027', lpns: 3, cost: 11500, bultos: 180, onHand: 2160, days: 450, badge: 'SALDO' },
      { sku: '10001007', desc: 'ENVASE EMPAQUE DEMO 280MM X 60U', div: 'J07-PLATOS PREPARADOS', rngFv: '05/08/2026', lpns: 2, cost: 9200, bultos: 60, onHand: 60, days: 390, badge: 'SALDO' },
      { sku: '10001008', desc: 'ORGANIZADOR MESA DEMO X 2 UN', div: 'J09-HOGAR', rngFv: '-', lpns: 4, cost: 8400, bultos: 120, onHand: 1440, days: 480, badge: 'SALDO' },
      { sku: '10001009', desc: 'VAJILLA APILABLE DEMO COLECCION', div: 'J09-HOGAR', rngFv: '-', lpns: 3, cost: 7600, bultos: 80, onHand: 80, days: 520, badge: 'SALDO' },
      { sku: '10001010', desc: 'UTENSILIO BAZAR DEMO MULTIUSO', div: 'J10-BAZAR', rngFv: '-', lpns: 3, cost: 6400, bultos: 160, onHand: 960, days: 410, badge: 'SALDO' }
    ],
    zoneTable: [
      {
        name: 'RCK',
        costos: [7500000, 950000, 210000, 90000, 8750000],
        lpn: [3600, 550, 110, 28, 4288],
        bultos: [135000, 16000, 3100, 1350, 155450]
      },
      {
        name: 'RHB',
        costos: [14000000, 1800000, 390000, 60000, 16250000],
        lpn: [6900, 1075, 215, 22, 8212],
        bultos: [256500, 31250, 5900, 900, 294550]
      }
    ],
    zoneTotals: {
      costos: [21500000, 2750000, 600000, 150000, 25000000],
      lpn: [10500, 1625, 325, 50, 12500],
      bultos: [391500, 47250, 9000, 2250, 450000]
    },
    topSkusRck: [
      { sku: '10001001', desc: 'PRODUCTO DEMO CONSERVAS 400G (EJEMPLO)', fv: '20/07/2027', lpns: 12, cost: 32000, bultos: 650 },
      { sku: '10001002', desc: 'BEBIDA DEMO SELECCION 750ML (EJEMPLO)', fv: '10/10/2028', lpns: 5, cost: 24500, bultos: 280 },
      { sku: '10001004', desc: 'SALSA CONDIMENTO DEMO X 400G', fv: '30/09/2027', lpns: 6, cost: 16400, bultos: 420 },
      { sku: '10001007', desc: 'ENVASE EMPAQUE DEMO 280MM X 60U', fv: '05/08/2026', lpns: 2, cost: 9200, bultos: 60 },
      { sku: '10001010', desc: 'UTENSILIO BAZAR DEMO MULTIUSO', fv: '-', lpns: 3, cost: 7900, bultos: 160 }
    ],
    topSkusRhb: [
      { sku: '10001003', desc: 'SET MENAJE COCINA DEMO X 4 UN', fv: '-', lpns: 7, cost: 19800, bultos: 190 },
      { sku: '10001005', desc: 'ACCESORIO HOGAR TEXTIL DEMO 2025', fv: '-', lpns: 5, cost: 14200, bultos: 110 },
      { sku: '10001006', desc: 'ADEREZO CLASICO DEMO X 425G', fv: '19/05/2027', lpns: 3, cost: 11500, bultos: 180 },
      { sku: '10001008', desc: 'ORGANIZADOR MESA DEMO X 2 UN', fv: '-', lpns: 4, cost: 8400, bultos: 120 },
      { sku: '10001009', desc: 'VAJILLA APILABLE DEMO COLECCION', fv: '-', lpns: 3, cost: 6100, bultos: 80 }
    ],
    topSkusByZone: {
      'RCK': {
        zone: 'RCK',
        skus: [
          { sku: '10001001', desc: 'PRODUCTO DEMO CONSERVAS 400G (EJEMPLO)', fv: '20/07/2027', lpns: 12, cost: 32000, bultos: 650 },
          { sku: '10001002', desc: 'BEBIDA DEMO SELECCION 750ML (EJEMPLO)', fv: '10/10/2028', lpns: 5, cost: 24500, bultos: 280 },
          { sku: '10001004', desc: 'SALSA CONDIMENTO DEMO X 400G', fv: '30/09/2027', lpns: 6, cost: 16400, bultos: 420 },
          { sku: '10001007', desc: 'ENVASE EMPAQUE DEMO 280MM X 60U', fv: '05/08/2026', lpns: 2, cost: 9200, bultos: 60 },
          { sku: '10001010', desc: 'UTENSILIO BAZAR DEMO MULTIUSO', fv: '-', lpns: 3, cost: 7900, bultos: 160 }
        ],
        totalCost: 90000,
        totalLpns: 28,
        totalBultos: 1570
      },
      'RHB': {
        zone: 'RHB',
        skus: [
          { sku: '10001003', desc: 'SET MENAJE COCINA DEMO X 4 UN', fv: '-', lpns: 7, cost: 19800, bultos: 190 },
          { sku: '10001005', desc: 'ACCESORIO HOGAR TEXTIL DEMO 2025', fv: '-', lpns: 5, cost: 14200, bultos: 110 },
          { sku: '10001006', desc: 'ADEREZO CLASICO DEMO X 425G', fv: '19/05/2027', lpns: 3, cost: 11500, bultos: 180 },
          { sku: '10001008', desc: 'ORGANIZADOR MESA DEMO X 2 UN', fv: '-', lpns: 4, cost: 8400, bultos: 120 },
          { sku: '10001009', desc: 'VAJILLA APILABLE DEMO COLECCION', fv: '-', lpns: 3, cost: 6100, bultos: 80 }
        ],
        totalCost: 60000,
        totalLpns: 22,
        totalBultos: 680
      }
    },
    zones: {
      rck: {
        totalCost: 8750000,
        lpns: 4288,
        bultos: 155450,
        pct: 35.0,
        r010: 7500000,
        r1025: 950000,
        r2552: 210000,
        r52: 90000,
        lpns52: 28,
        topSkus: []
      },
      rhb: {
        totalCost: 16250000,
        lpns: 8212,
        bultos: 294550,
        pct: 65.0,
        r010: 14000000,
        r1025: 1800000,
        r2552: 390000,
        r52: 60000,
        lpns52: 22,
        topSkus: []
      }
    },
    weeks: ['S-26', 'S-27', 'S-28', 'S-29', 'S-30', 'S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36', 'S-37'],
    evolution: [
      { label: '0 a 10 Semanas', values: [88.00, 88.20, 88.50, 88.80, 88.40, 88.10, 87.50, 86.80, 86.90, 87.20, 87.00, 87.00], color: '#10b981' },
      { label: '10 a 25 Semanas', values: [9.50, 9.30, 9.10, 8.90, 9.20, 9.50, 10.00, 10.60, 10.50, 10.30, 10.50, 10.50], color: '#f59e0b' },
      { label: '25 a 52 Semanas', values: [2.00, 2.00, 1.90, 1.80, 1.90, 1.90, 2.00, 2.10, 2.10, 2.00, 2.00, 2.00], color: '#f97316' },
      { label: 'Mayor a 52 Semanas', values: [0.50, 0.50, 0.50, 0.50, 0.50, 0.50, 0.50, 0.50, 0.50, 0.50, 0.50, 0.50], color: '#ef4444' }
    ],
    locationsTotal: 10000,
    locations: [
      { div: 'J01-PGC COMESTIBLE', r010: 3400, r1025: 350, r2552: 40, r52: 10, total: 3800, pct: 38.0 },
      { div: 'J09-HOGAR', r010: 2400, r1025: 500, r2552: 80, r52: 20, total: 3000, pct: 30.0 },
      { div: 'J08-VESTUARIO', r010: 1100, r1025: 250, r2552: 50, r52: 0, total: 1400, pct: 14.0 },
      { div: 'J02-PGC NO COMESTIBLE', r010: 850, r1025: 130, r2552: 15, r52: 5, total: 1000, pct: 10.0 },
      { div: 'J10-BAZAR', r010: 400, r1025: 80, r2552: 15, r52: 5, total: 500, pct: 5.0 },
      { div: 'J11-ELECTROHOGAR', r010: 250, r1025: 40, r2552: 10, r52: 0, total: 300, pct: 3.0 }
    ]
  };

  // CD Frescos 676 (Datos Sintéticos de Demostración)
  ACTIVE_DATABASE['676'] = {
    name: 'Centro de Distribución - Frescos (Demo)',
    whseCode: '676',
    badgeText: 'CD FRESCOS 676',
    totalCost: 5000000,
    totalLpns: 1800,
    totalBultos: 90000,
    ranges: [
      { label: '0 a 10 Semanas', cost: 4725000, costPct: 94.50, lpns: 1710, lpnsPct: 95.00, bultos: 85500, bultosPct: 95.00, color: '#10b981', status: 'Saludable' },
      { label: '10 a 25 Semanas', cost: 225000, costPct: 4.50, lpns: 65, lpnsPct: 3.61, bultos: 3600, bultosPct: 4.00, color: '#f59e0b', status: 'En Alerta' },
      { label: '25 a 52 Semanas', cost: 40000, costPct: 0.80, lpns: 18, lpnsPct: 1.00, bultos: 720, bultosPct: 0.80, color: '#f97316', status: 'Riesgo Medio' },
      { label: 'mayor a 52 Semanas', cost: 10000, costPct: 0.20, lpns: 7, lpnsPct: 0.39, bultos: 180, bultosPct: 0.20, color: '#ef4444', status: 'Crítico >1 año' }
    ],
    divisions: [
      { code: 'J03-CARNES Y PESCADOS', r010: 1350000, r1025: 60000, r2552: 8000, r52: 2000, total: 1420000, pct: 28.4 },
      { code: 'J05-FLC', r010: 950000, r1025: 50000, r2552: 10000, r52: 0, total: 1010000, pct: 20.2 },
      { code: 'J06-PANADERIA Y PASTELERIA', r010: 650000, r1025: 25000, r2552: 5000, r52: 3000, total: 683000, pct: 13.7 },
      { code: 'J01-PGC COMESTIBLE', r010: 600000, r1025: 40000, r2552: 0, r52: 0, total: 640000, pct: 12.8 },
      { code: 'J04-FRUTAS Y VERDURAS', r010: 625000, r1025: 0, r2552: 0, r52: 0, total: 625000, pct: 12.5 },
      { code: 'J07-PLATOS PREPARADOS', r010: 550000, r1025: 50000, r2552: 17000, r52: 5000, total: 622000, pct: 12.4 }
    ],
    top10Skus: [
      { sku: '20002001', desc: 'INSUMO EMPAQUE REFRIGERADO DEMO A', div: 'J07-PLATOS PREPARADOS', rngFv: '-', lpns: 3, cost: 5000, bultos: 60, onHand: 6000, days: 520, badge: 'CRÍTICO #1' },
      { sku: '20002002', desc: 'ENVASE PANADERIA DEMO MULTIUSO B', div: 'J06-PANADERIA Y PASTELERIA', rngFv: '-', lpns: 2, cost: 3000, bultos: 40, onHand: 4000, days: 480, badge: 'CRÍTICO #2' },
      { sku: '20002003', desc: 'PRODUCTO CONGELADO DEMO X KG', div: 'J03-CARNES Y PESCADOS', rngFv: '15/06/2027', lpns: 2, cost: 2000, bultos: 80, onHand: 800, days: 420, badge: 'PRIORIDAD' }
    ],
    zoneTable: [
      {
        name: 'RCK',
        costos: [1890000, 90000, 16000, 5000, 2001000],
        lpn: [684, 26, 7, 3, 720],
        bultos: [34200, 1440, 288, 60, 35988]
      },
      {
        name: 'RHB',
        costos: [2835000, 135000, 24000, 5000, 2999000],
        lpn: [1026, 39, 11, 4, 1080],
        bultos: [51300, 2160, 432, 120, 54012]
      }
    ],
    zoneTotals: {
      costos: [4725000, 225000, 40000, 10000, 5000000],
      lpn: [1710, 65, 18, 7, 1800],
      bultos: [85500, 3600, 720, 180, 90000]
    },
    topSkusRck: [
      { sku: '20002001', desc: 'INSUMO EMPAQUE REFRIGERADO DEMO A', fv: '-', lpns: 3, cost: 5000, bultos: 60 }
    ],
    topSkusRhb: [
      { sku: '20002002', desc: 'ENVASE PANADERIA DEMO MULTIUSO B', fv: '-', lpns: 2, cost: 3000, bultos: 40 },
      { sku: '20002003', desc: 'PRODUCTO CONGELADO DEMO X KG', fv: '15/06/2027', lpns: 2, cost: 2000, bultos: 80 }
    ],
    topSkusByZone: {
      'RCK': {
        zone: 'RCK',
        skus: [
          { sku: '20002001', desc: 'INSUMO EMPAQUE REFRIGERADO DEMO A', fv: '-', lpns: 3, cost: 5000, bultos: 60 }
        ],
        totalCost: 5000,
        totalLpns: 3,
        totalBultos: 60
      },
      'RHB': {
        zone: 'RHB',
        skus: [
          { sku: '20002002', desc: 'ENVASE PANADERIA DEMO MULTIUSO B', fv: '-', lpns: 2, cost: 3000, bultos: 40 },
          { sku: '20002003', desc: 'PRODUCTO CONGELADO DEMO X KG', fv: '15/06/2027', lpns: 2, cost: 2000, bultos: 80 }
        ],
        totalCost: 5000,
        totalLpns: 4,
        totalBultos: 120
      }
    },
    zones: {
      rck: {
        totalCost: 2001000,
        lpns: 720,
        bultos: 35988,
        pct: 40.0,
        r010: 1890000,
        r1025: 90000,
        r2552: 16000,
        r52: 5000,
        lpns52: 3,
        topSkus: []
      },
      rhb: {
        totalCost: 2999000,
        lpns: 1080,
        bultos: 54012,
        pct: 60.0,
        r010: 2835000,
        r1025: 135000,
        r2552: 24000,
        r52: 5000,
        lpns52: 4,
        topSkus: []
      }
    },
    weeks: ['S-26', 'S-27', 'S-28', 'S-29', 'S-30', 'S-31', 'S-32', 'S-33', 'S-34', 'S-35', 'S-36', 'S-37'],
    evolution: [
      { label: '0 a 10 Semanas', values: [94.00, 94.20, 94.50, 94.80, 95.00, 95.20, 94.80, 94.10, 93.90, 94.20, 94.50, 95.00], color: '#10b981' },
      { label: '10 a 25 Semanas', values: [5.00, 4.80, 4.50, 4.20, 4.00, 3.80, 4.20, 4.90, 5.10, 4.80, 4.50, 4.00], color: '#f59e0b' },
      { label: '25 a 52 Semanas', values: [0.80, 0.80, 0.80, 0.80, 0.80, 0.80, 0.80, 0.80, 0.80, 0.80, 0.80, 0.80], color: '#f97316' },
      { label: 'Mayor a 52 Semanas', values: [0.20, 0.20, 0.20, 0.20, 0.20, 0.20, 0.20, 0.20, 0.20, 0.20, 0.20, 0.20], color: '#ef4444' }
    ],
    locationsTotal: 1800,
    locations: [
      { div: 'J03-CARNES Y PESCADOS', r010: 480, r1025: 15, r2552: 3, r52: 2, total: 500, pct: 27.8 },
      { div: 'J05-FLC', r010: 360, r1025: 15, r2552: 5, r52: 0, total: 380, pct: 21.1 },
      { div: 'J06-PANADERIA Y PASTELERIA', r010: 240, r1025: 6, r2552: 2, r52: 2, total: 250, pct: 13.9 },
      { div: 'J01-PGC COMESTIBLE', r010: 230, r1025: 10, r2552: 0, r52: 0, total: 240, pct: 13.3 },
      { div: 'J04-FRUTAS Y VERDURAS', r010: 230, r1025: 0, r2552: 0, r52: 0, total: 230, pct: 12.8 },
      { div: 'J07-PLATOS PREPARADOS', r010: 170, r1025: 19, r2552: 8, r52: 3, total: 200, pct: 11.1 }
    ]
  };

  document.getElementById('dropzoneBox').style.display = 'none';
  document.getElementById('fileStatusBar').style.display = 'flex';
  document.getElementById('loadedFileName').textContent = 'Datos Demostrativos (Modo Ejemplo)';
  document.getElementById('loadedFileMeta').textContent = 'Valores sintéticos de demostración para CD Secos 655 y Frescos 676';

  updateHeaderWeekBadges();
  renderWarehouseData(currentWarehouse);
  hideLoading();
  showToast('Datos de ejemplo cargados con éxito', 'success', 3000);
}

// ══════════════════════════════════════════════════════════════════════════════
// 8. RENDERIZADO VISUAL DEL DASHBOARD
// ══════════════════════════════════════════════════════════════════════════════

function updateHeaderWeekBadges() {
  ['slide1-week-badge', 'slide2-week-badge', 'slide3-week-badge'].forEach(id => {
    const badge = document.getElementById(id);
    if (badge) badge.innerHTML = `<i class="fa-regular fa-calendar"></i> ${currentWeekLabel}`;
  });
}

function switchWarehouse(whseCode) {
  currentWarehouse = whseCode;
  document.getElementById('btnWhseSecos').classList.toggle('active', whseCode === '655');
  document.getElementById('btnWhseFrescos').classList.toggle('active', whseCode === '676');
  renderWarehouseData(whseCode);
  showToast(`Mostrando ${whseCode === '655' ? 'CD Secos (655)' : 'CD Frescos (676)'}`, 'info');
}

function switchSlide(slideNum) {
  currentSlide = slideNum;
  [1, 2, 3].forEach(n => {
    document.getElementById(`tabBtnSlide${n}`).classList.toggle('active', n === slideNum);
    const container = document.getElementById(`slide${n}-container`);
    if (container) container.style.display = (n === slideNum) ? 'block' : 'none';
  });

  setTimeout(() => {
    if (slideNum === 1 && chartInstanceS1) chartInstanceS1.resize();
    if (slideNum === 3 && chartInstanceS3) chartInstanceS3.resize();
  }, 60);
}

function renderWarehouseData(whseCode) {
  const data = ACTIVE_DATABASE[whseCode];
  if (!data) return;

  ['slide1', 'slide2', 'slide3'].forEach(id => {
    const el = document.getElementById(`${id}-whse-badge`);
    if (el) el.textContent = data.badgeText;
  });

  renderSlide1(data);
  renderSlide2(data);
  renderSlide3(data);
}

function renderSlide1(data) {
  document.getElementById('s1-kpi-total').textContent = formatCurrency(data.totalCost);
  document.getElementById('s1-kpi-lpns').textContent = `${formatNumber(data.totalLpns)} LPNs`;
  document.getElementById('s1-kpi-bultos').textContent = `${formatNumber(data.totalBultos)} Bultos`;

  const r010 = (data.ranges && data.ranges[0]) ? data.ranges[0] : { costPct: 0, cost: 0, bultos: 0, bultosPct: 0 };
  const r1025 = (data.ranges && data.ranges[1]) ? data.ranges[1] : { costPct: 0, cost: 0, bultos: 0, bultosPct: 0, lpns: 0 };
  const r2552 = (data.ranges && data.ranges[2]) ? data.ranges[2] : { cost: 0, costPct: 0, bultos: 0, bultosPct: 0 };
  const r52 = (data.ranges && data.ranges[3]) ? data.ranges[3] : { cost: 0, costPct: 0, bultos: 0, bultosPct: 0, lpns: 0 };

  const critCost = (r2552.cost || 0) + (r52.cost || 0);
  const critPct = (r2552.costPct || 0) + (r52.costPct || 0);
  const critBultos = (r2552.bultos || 0) + (r52.bultos || 0);
  const critBultosPct = (r2552.bultosPct || 0) + (r52.bultosPct || 0);

  // Tarjeta 2: Saludable (0 a 10 Sem) - % Costo destacado (alineado con la tabla) + % Bultos
  document.getElementById('s1-kpi-healthy-pct').innerHTML = `${safeFixed(r010.costPct, 1, '%')} <span style="font-size:0.75rem; font-weight:800; color:#065f46; letter-spacing:0.3px;">COSTO</span>`;
  document.getElementById('s1-kpi-healthy-cost').textContent = `${formatCurrency(r010.cost)} inmovilizado`;
  document.getElementById('s1-kpi-healthy-bultos').innerHTML = `📦 <strong>${safeFixed(r010.bultosPct, 2, '%')}</strong> Bultos (${formatNumber(r010.bultos)})`;

  // Tarjeta 3: En Seguimiento (10 a 25 Sem) - 11.1% Costo + 10.18% Bultos
  document.getElementById('s1-kpi-warn-pct').innerHTML = `${safeFixed(r1025.costPct, 1, '%')} <span style="font-size:0.75rem; font-weight:800; color:#92400e; letter-spacing:0.3px;">COSTO</span>`;
  document.getElementById('s1-kpi-warn-cost').textContent = `${formatCurrency(r1025.cost)} inmovilizado`;
  document.getElementById('s1-kpi-warn-lpns').innerHTML = `📦 <strong>${safeFixed(r1025.bultosPct, 2, '%')}</strong> Bultos (${formatNumber(r1025.bultos)})`;

  // Tarjeta 4: Crítico (>25 Sem) - 2.1% Costo + 1.86% Bultos
  document.getElementById('s1-kpi-crit-cost').innerHTML = `${safeFixed(critPct, 1, '%')} <span style="font-size:0.75rem; font-weight:800; color:#991b1b; letter-spacing:0.3px;">COSTO</span>`;
  document.getElementById('s1-kpi-crit-pct').textContent = `${formatCurrency(critCost)} (${safeFixed(critBultosPct, 2, '%')} Bultos)`;
  document.getElementById('s1-kpi-over52').textContent = `>52s: ${formatCurrency(r52.cost)} (${formatNumber(r52.bultos)} Bultos)`;

  // 1. Gráfico S1 con valores encima de las barras
  renderChartS1(data.ranges);

  // 2. Tabla 1 S1: Resumen de Rangos (Table 1 de la presentación original)
  renderTableS1Ranges(data.ranges, data.totalCost, data.totalLpns, data.totalBultos);

  // 3. Tabla 2 S1: Matriz Divisiones (Detalle exacto de PPT y Tabla Dinámica)
  const tbodyDiv = document.getElementById('tbodyS1Divisions');
  tbodyDiv.innerHTML = '';
  if (data.divisions && data.divisions.length > 0) {
    data.divisions.forEach(d => {
      const p010 = r010.cost > 0 ? (d.r010 / r010.cost) * 100 : 0;
      const p1025 = r1025.cost > 0 ? (d.r1025 / r1025.cost) * 100 : 0;
      const p2552 = r2552.cost > 0 ? (d.r2552 / r2552.cost) * 100 : 0;
      const p52 = r52.cost > 0 ? (d.r52 / r52.cost) * 100 : 0;
      const pTot = data.totalCost > 0 ? (d.total / data.totalCost) * 100 : 0;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-weight:700; color:#0f172a; text-align:left;">${escapeHtml(d.code)}</td>
        <td>${d.r010 > 0 ? formatNumber(d.r010) : '-'}</td>
        <td style="color:#059669; font-weight:600;">${d.r010 > 0 ? (p010 < 1 ? safeFixed(p010, 1, '%') : safeFixed(p010, 0, '%')) : '0%'}</td>
        <td>${d.r1025 > 0 ? formatNumber(d.r1025) : '-'}</td>
        <td style="color:#d97706; font-weight:600;">${d.r1025 > 0 ? (p1025 < 1 ? safeFixed(p1025, 1, '%') : safeFixed(p1025, 0, '%')) : '0%'}</td>
        <td style="${d.r2552 > 100000 ? 'background:#fffbeb; font-weight:700; color:#b45309;' : ''}">${d.r2552 > 0 ? formatNumber(d.r2552) : '-'}</td>
        <td style="color:#ea580c; font-weight:600;">${d.r2552 > 0 ? (p2552 < 1 ? safeFixed(p2552, 1, '%') : safeFixed(p2552, 0, '%')) : '0%'}</td>
        <td style="${d.r52 > 10000 ? 'background:#fee2e2; font-weight:800; color:#dc2626;' : ''}">${d.r52 > 0 ? formatNumber(d.r52) : '-'}</td>
        <td style="color:#dc2626; font-weight:700;">${d.r52 > 0 ? (p52 < 1 ? safeFixed(p52, 1, '%') : safeFixed(p52, 0, '%')) : '0%'}</td>
        <td style="font-weight:800; color:#0f172a;">${formatNumber(d.total)}</td>
        <td style="color:#2563eb; font-weight:800;">${pTot < 1 ? safeFixed(pTot, 2, '%') : safeFixed(pTot, 0, '%')}</td>
      `;
      tbodyDiv.appendChild(tr);
    });
  } else {
    tbodyDiv.innerHTML = `<tr><td colspan="11" style="text-align:center; padding:12px; color:#64748b;">No hay divisiones registradas</td></tr>`;
  }

  document.getElementById('tfootS1Divisions').innerHTML = `
    <tr>
      <td style="text-align:left; font-weight:900;">Total general</td>
      <td style="font-weight:900;">${formatNumber(r010.cost)}</td>
      <td style="font-weight:900; color:#059669;">100%</td>
      <td style="font-weight:900;">${formatNumber(r1025.cost)}</td>
      <td style="font-weight:900; color:#d97706;">100%</td>
      <td style="font-weight:900; color:#ea580c;">${formatNumber(r2552.cost)}</td>
      <td style="font-weight:900; color:#ea580c;">100%</td>
      <td style="font-weight:900; color:#dc2626;">${formatNumber(r52.cost)}</td>
      <td style="font-weight:900; color:#dc2626;">100%</td>
      <td style="color:#0f172a; font-weight:900;">${formatNumber(data.totalCost)}</td>
      <td style="color:#2563eb; font-weight:900;">100%</td>
    </tr>
  `;

  // 4. Tabla 3 S1: Top 10 SKUs >52 Semanas (Suma de LPNs)
  const tbodyTop = document.getElementById('tbodyS1Top10');
  tbodyTop.innerHTML = '';
  let sumTopCost = 0, sumTopLpns = 0, sumTopBultos = 0, sumTopOnHand = 0;

  if (!data.top10Skus || data.top10Skus.length === 0) {
    tbodyTop.innerHTML = `<tr><td colspan="10" style="text-align:center; padding:16px; color:#64748b; font-weight:600;">No se registraron SKUs con más de 52 semanas de antigüedad</td></tr>`;
  } else {
    data.top10Skus.forEach((sku, idx) => {
      sumTopCost += (sku.cost || 0);
      sumTopLpns += (sku.lpns || 0);
      sumTopBultos += (sku.bultos || 0);
      sumTopOnHand += (sku.onHand || 0);

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-family:'JetBrains Mono',monospace; font-weight:800; color:#1e293b;">${escapeHtml(sku.sku)}</td>
        <td style="text-align:left; font-weight:700; color:#0f172a;">${escapeHtml(sku.desc)}</td>
        <td style="text-align:left; color:#475569; font-weight:600;">${escapeHtml(sku.div)}</td>
        <td style="text-align:center; color:#64748b; font-size:0.75rem;">${escapeHtml(sku.rngFv)}</td>
        <td style="font-weight:800; color:#2563eb;">${formatNumber(sku.lpns)}</td>
        <td style="font-weight:800; color:${sku.cost > 20000 ? '#dc2626' : '#0f172a'};">${formatNumber(sku.cost)}</td>
        <td style="font-weight:600;">${formatNumber(sku.bultos)}</td>
        <td style="color:#64748b;">${formatNumber(sku.onHand)}</td>
        <td style="font-weight:700; color:${sku.days > 500 ? '#b91c1c' : '#475569'};">${escapeHtml(sku.days)} d</td>
        <td style="text-align:center;">
          <span style="background:${idx < 2 ? '#fee2e2' : '#f1f5f9'}; color:${idx < 2 ? '#b91c1c' : '#475569'}; padding:2px 6px; border-radius:5px; font-weight:800; font-size:0.72rem;">
            ${escapeHtml(sku.badge)}
          </span>
        </td>
      `;
      tbodyTop.appendChild(tr);
    });
  }

  document.getElementById('tfootS1Top10').innerHTML = `
    <tr style="background:#f1f5f9; font-weight:900; border-top:2px solid #0f172a;">
      <td colspan="4" style="text-align:left; font-weight:900; color:#0f172a; padding:4px 6px;">Total Top 10 SKUs</td>
      <td style="color:#2563eb; font-weight:900; padding:4px 6px;">${formatNumber(sumTopLpns)}</td>
      <td style="color:#dc2626; font-size:0.90rem; font-weight:900; padding:4px 6px;">S/ ${formatNumber(sumTopCost)}</td>
      <td style="font-weight:900; color:#0f172a; padding:4px 6px;">${formatNumber(sumTopBultos)}</td>
      <td style="font-weight:800; color:#475569; padding:4px 6px;">${formatNumber(sumTopOnHand)}</td>
      <td colspan="2" style="text-align:center; color:#0f172a; font-weight:800; padding:4px 6px; background:#e2e8f0;">Concentración crítica</td>
    </tr>
  `;
}

function renderTableS1Ranges(ranges, totalCost, totalLpns, totalBultos) {
  const tbody = document.getElementById('tbodyS1Ranges');
  const tfoot = document.getElementById('tfootS1Ranges');
  if (!tbody || !tfoot) return;

  tbody.innerHTML = '';
  (ranges || []).forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align:left; font-weight:700; color:#0f172a;">
        <span style="display:inline-block; width:8px; height:8px; border-radius:50%; background:${escapeHtml(r.color)}; margin-right:6px;"></span>
        ${escapeHtml(r.label)}
      </td>
      <td style="font-weight:700; color:#2563eb;">${formatNumber(r.lpns)}</td>
      <td style="font-weight:600; color:#475569;">${safeFixed(r.lpnsPct, 2, '%')}</td>
      <td>${formatNumber(r.bultos)}</td>
      <td style="font-weight:800; color:${r.color};">${safeFixed(r.bultosPct, 2, '%')}</td>
      <td style="font-weight:700;">${formatCurrency(r.cost)}</td>
      <td style="font-weight:800;">${safeFixed(r.costPct, 1, '%')}</td>
    `;
    tbody.appendChild(tr);
  });

  tfoot.innerHTML = `
    <tr>
      <td style="text-align:left; font-weight:900;">Total general</td>
      <td style="color:#2563eb; font-weight:900;">${formatNumber(totalLpns)}</td>
      <td style="font-weight:900; color:#475569;">100.00%</td>
      <td style="font-weight:900;">${formatNumber(totalBultos)}</td>
      <td style="font-weight:900;">100%</td>
      <td style="color:#0f172a; font-weight:900;">${formatCurrency(totalCost)}</td>
      <td style="font-weight:900;">100%</td>
    </tr>
  `;
}

// ══════════════════════════════════════════════════════════════════════════════
// 9. GRÁFICOS CHART.JS CON DATALABELS VISIBLES ENCIMA DE LAS BARRAS
// ══════════════════════════════════════════════════════════════════════════════

function setChartS1Mode(mode) {
  currentChartS1Mode = mode;
  const btnBultos = document.getElementById('btnChartModeBultos');
  const btnCosto = document.getElementById('btnChartModeCosto');
  const title = document.getElementById('chartS1Title');

  if (btnBultos && btnCosto) {
    if (mode === 'bultos') {
      btnBultos.style.background = '#ffffff';
      btnBultos.style.color = '#1e3a8a';
      btnBultos.style.fontWeight = '800';
      btnBultos.style.boxShadow = '0 1px 2px rgba(0,0,0,0.08)';

      btnCosto.style.background = 'transparent';
      btnCosto.style.color = '#64748b';
      btnCosto.style.fontWeight = '700';
      btnCosto.style.boxShadow = 'none';

      if (title) title.textContent = 'Distribución de Bultos por Antigüedad';
    } else {
      btnCosto.style.background = '#ffffff';
      btnCosto.style.color = '#1e3a8a';
      btnCosto.style.fontWeight = '800';
      btnCosto.style.boxShadow = '0 1px 2px rgba(0,0,0,0.08)';

      btnBultos.style.background = 'transparent';
      btnBultos.style.color = '#64748b';
      btnBultos.style.fontWeight = '700';
      btnBultos.style.boxShadow = 'none';

      if (title) title.textContent = 'Distribución de Costo por Antigüedad';
    }
  }

  const data = ACTIVE_DATABASE[currentWarehouse];
  if (data && data.ranges) {
    renderChartS1(data.ranges);
  }
}
window.setChartS1Mode = setChartS1Mode;

function renderChartS1(ranges) {
  const ctx = document.getElementById('chartS1Ranges');
  if (!ctx) return;

  if (chartInstanceS1) {
    chartInstanceS1.destroy();
  }

  const isBultos = (currentChartS1Mode === 'bultos');
  const labels = ranges.map(r => r.label);
  const colors = ranges.map(r => r.color);

  // Datos según modo activo:
  const chartData = isBultos
    ? ranges.map(r => r.bultos)
    : ranges.map(r => {
        return r.cost >= 1000000 ? parseFloat((r.cost / 1000000).toFixed(2)) : parseFloat((r.cost / 1000000).toFixed(3));
      });

  // Plugin personalizado para pintar valores y porcentajes ENCIMA de cada barra
  const valueLabelsPlugin = {
    id: 'barValueLabelsOnTop',
    afterDatasetsDraw(chart) {
      const { ctx: c } = chart;
      c.save();
      const meta = chart.getDatasetMeta(0);
      if (!meta || !meta.data) return;

      meta.data.forEach((bar, idx) => {
        const rItem = ranges[idx];
        if (!bar || !rItem) return;

        c.textAlign = 'center';
        c.textBaseline = 'bottom';

        if (isBultos) {
          // Línea 1: Bultos (ej. 117,009 Bultos)
          const bultosText = `${formatNumber(rItem.bultos)} Bultos`;
          c.font = 'bold 11px Inter, sans-serif';
          c.fillStyle = '#0f172a';
          c.fillText(bultosText, bar.x, bar.y - 14);

          // Línea 2: Porcentaje de Bultos (ej. 95.50%)
          const pctText = safeFixed(rItem.bultosPct, 2, '%');
          c.font = '800 11px Inter, sans-serif';
          c.fillStyle = rItem.color || '#2563eb';
          c.fillText(`(${pctText})`, bar.x, bar.y - 2);
        } else {
          // Línea 1: Costo (S/ 7.5M o S/ 334K)
          const costText = formatCompact(rItem.cost);
          c.font = 'bold 12px Inter, sans-serif';
          c.fillStyle = '#0f172a';
          c.fillText(`S/ ${costText}`, bar.x, bar.y - 14);

          // Línea 2: Porcentaje de Costo (ej. 94.9%)
          const pctText = safeFixed(rItem.costPct, 1, '%');
          c.font = '800 11px Inter, sans-serif';
          c.fillStyle = rItem.color || '#2563eb';
          c.fillText(`(${pctText})`, bar.x, bar.y - 2);
        }
      });
      c.restore();
    }
  };

  chartInstanceS1 = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: isBultos ? 'Bultos' : 'Costo (Millones S/)',
        data: chartData,
        backgroundColor: colors,
        borderRadius: 8,
        borderWidth: 0
      }]
    },
    plugins: [valueLabelsPlugin],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: {
        padding: {
          top: 32 // Margen superior para que el texto encima de la barra más alta nunca se corte
        }
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (item) => {
              const r = ranges[item.dataIndex];
              if (!r) return '';
              if (isBultos) {
                return ` ${formatNumber(r.bultos)} Bultos (${safeFixed(r.bultosPct, 2, '%')})`;
              } else {
                return ` ${formatCurrency(r.cost)} (${safeFixed(r.costPct, 1, '%')})`;
              }
            }
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          grace: '30%', // 30% de altura adicional sobre la barra mayor para alojar los datalabels
          grid: { color: '#f1f5f9' },
          ticks: {
            callback: (v) => isBultos ? formatCompact(v) : `S/ ${v}M`,
            font: { size: 11, weight: '700' }
          }
        },
        x: {
          grid: { display: false },
          ticks: { font: { size: 11, weight: '700' } }
        }
      }
    }
  });
}

function renderSlide2(data) {
  // 1. Renderizar Tabla Principal de Antigüedad por Zonas (Pivot Completo)
  const tbodyZones = document.getElementById('tbodyS2Zones');
  const tfootZones = document.getElementById('tfootS2Zones');
  if (tbodyZones && tfootZones) {
    tbodyZones.innerHTML = '';
    tfootZones.innerHTML = '';

    const zTable = data.zoneTable || [];
    if (zTable.length === 0) {
      tbodyZones.innerHTML = `<tr><td colspan="7" style="text-align:center; padding:14px; color:#64748b;">No se encontraron registros de zonas operativas</td></tr>`;
    } else {
      zTable.forEach((z) => {
        // Fila 1: COSTOS (con celda ZONA con rowspan=3)
        const trCost = document.createElement('tr');
        trCost.innerHTML = `
          <td rowspan="3" style="text-align:center; vertical-align:middle; font-weight:900; background:#f0fdf4; color:#166534; border-right:2px solid #bbf7d0; font-size:0.92rem;">
            ${escapeHtml(z.name)}
          </td>
          <td style="font-weight:800; color:#0f172a; text-align:left; background:#f8fafc;">COSTOS</td>
          <td style="text-align:right; font-weight:700;">${z.costos[0] > 0 ? formatNumber(z.costos[0]) : '-'}</td>
          <td style="text-align:right; font-weight:700;">${z.costos[1] > 0 ? formatNumber(z.costos[1]) : '-'}</td>
          <td style="text-align:right; font-weight:700; ${z.costos[2] > 50000 ? 'background:#fffbeb; color:#b45309;' : ''}">${z.costos[2] > 0 ? formatNumber(z.costos[2]) : '-'}</td>
          <td style="text-align:right; font-weight:800; ${z.costos[3] > 10000 ? 'background:#fee2e2; color:#dc2626;' : ''}">${z.costos[3] > 0 ? formatNumber(z.costos[3]) : '-'}</td>
          <td style="text-align:right; font-weight:900; background:#f8fafc; color:#0f172a;">${formatNumber(z.costos[4])}</td>
        `;
        tbodyZones.appendChild(trCost);

        // Fila 2: LPN
        const trLpn = document.createElement('tr');
        trLpn.innerHTML = `
          <td style="font-weight:700; color:#2563eb; text-align:left; background:#f8fafc;">LPN</td>
          <td style="text-align:right; color:#2563eb; font-weight:600;">${z.lpn[0] > 0 ? formatNumber(z.lpn[0]) : '-'}</td>
          <td style="text-align:right; color:#2563eb; font-weight:600;">${z.lpn[1] > 0 ? formatNumber(z.lpn[1]) : '-'}</td>
          <td style="text-align:right; color:#2563eb; font-weight:600;">${z.lpn[2] > 0 ? formatNumber(z.lpn[2]) : '-'}</td>
          <td style="text-align:right; color:#dc2626; font-weight:800;">${z.lpn[3] > 0 ? formatNumber(z.lpn[3]) : '-'}</td>
          <td style="text-align:right; font-weight:900; color:#2563eb; background:#f8fafc;">${formatNumber(z.lpn[4])}</td>
        `;
        tbodyZones.appendChild(trLpn);

        // Fila 3: BULTOS (con borde inferior más marcado para separar zonas)
        const trBultos = document.createElement('tr');
        trBultos.style.borderBottom = '2px solid #cbd5e1';
        trBultos.innerHTML = `
          <td style="font-weight:700; color:#475569; text-align:left; background:#f8fafc;">BULTOS</td>
          <td style="text-align:right; font-weight:600;">${z.bultos[0] > 0 ? formatNumber(z.bultos[0]) : '-'}</td>
          <td style="text-align:right; font-weight:600;">${z.bultos[1] > 0 ? formatNumber(z.bultos[1]) : '-'}</td>
          <td style="text-align:right; font-weight:600;">${z.bultos[2] > 0 ? formatNumber(z.bultos[2]) : '-'}</td>
          <td style="text-align:right; font-weight:700; color:#dc2626;">${z.bultos[3] > 0 ? formatNumber(z.bultos[3]) : '-'}</td>
          <td style="text-align:right; font-weight:900; color:#0f172a; background:#f8fafc;">${formatNumber(z.bultos[4])}</td>
        `;
        tbodyZones.appendChild(trBultos);
      });
    }

    // Pie de la Tabla de Zonas (Totales generales exactos)
    const zTot = data.zoneTotals || {
      costos: [0, 0, 0, 0, 0],
      lpn: [0, 0, 0, 0, 0],
      bultos: [0, 0, 0, 0, 0]
    };

    tfootZones.innerHTML = `
      <tr style="background:#f1f5f9; font-weight:900; border-top:2px solid #1e4620; font-size:0.84rem;">
        <td colspan="2" style="text-align:left; font-weight:900; color:#0f172a; padding:5px 10px;">Total COSTOS</td>
        <td style="text-align:right; font-weight:900; padding:5px 10px;">${formatNumber(zTot.costos[0])}</td>
        <td style="text-align:right; font-weight:900; padding:5px 10px;">${formatNumber(zTot.costos[1])}</td>
        <td style="text-align:right; font-weight:900; color:#ea580c; padding:5px 10px;">${formatNumber(zTot.costos[2])}</td>
        <td style="text-align:right; font-weight:900; color:#dc2626; padding:5px 10px;">${formatNumber(zTot.costos[3])}</td>
        <td style="text-align:right; font-weight:900; color:#0f172a; background:#e2e8f0; padding:5px 10px;">${formatNumber(zTot.costos[4])}</td>
      </tr>
      <tr style="background:#f1f5f9; font-weight:900; font-size:0.84rem;">
        <td colspan="2" style="text-align:left; font-weight:900; color:#2563eb; padding:5px 10px;">Total LPN</td>
        <td style="text-align:right; font-weight:900; color:#2563eb; padding:5px 10px;">${formatNumber(zTot.lpn[0])}</td>
        <td style="text-align:right; font-weight:900; color:#2563eb; padding:5px 10px;">${formatNumber(zTot.lpn[1])}</td>
        <td style="text-align:right; font-weight:900; color:#2563eb; padding:5px 10px;">${formatNumber(zTot.lpn[2])}</td>
        <td style="text-align:right; font-weight:900; color:#dc2626; padding:5px 10px;">${formatNumber(zTot.lpn[3])}</td>
        <td style="text-align:right; font-weight:900; color:#2563eb; background:#e2e8f0; padding:5px 10px;">${formatNumber(zTot.lpn[4])}</td>
      </tr>
      <tr style="background:#f1f5f9; font-weight:900; font-size:0.84rem;">
        <td colspan="2" style="text-align:left; font-weight:900; color:#475569; padding:5px 10px;">Total BULTOS</td>
        <td style="text-align:right; font-weight:900; padding:5px 10px;">${formatNumber(zTot.bultos[0])}</td>
        <td style="text-align:right; font-weight:900; padding:5px 10px;">${formatNumber(zTot.bultos[1])}</td>
        <td style="text-align:right; font-weight:900; padding:5px 10px;">${formatNumber(zTot.bultos[2])}</td>
        <td style="text-align:right; font-weight:900; color:#dc2626; padding:5px 10px;">${formatNumber(zTot.bultos[3])}</td>
        <td style="text-align:right; font-weight:900; color:#0f172a; background:#e2e8f0; padding:5px 10px;">${formatNumber(zTot.bultos[4])}</td>
      </tr>
    `;
  }

  // 2. Tablas Dinámicas: Top 10 SKUs en Riesgo (>52 semanas) por Zona
  const containerTopZones = document.getElementById('containerS2TopZones');
  if (containerTopZones) {
    containerTopZones.innerHTML = '';

    const topByZone = data.topSkusByZone || {};
    let zonesWithData = Object.values(topByZone).filter(z => z.skus && z.skus.length > 0);

    // Fallback con topSkusRck / topSkusRhb por compatibilidad
    if (zonesWithData.length === 0) {
      if (data.topSkusRck && data.topSkusRck.length > 0) {
        zonesWithData.push({
          zone: 'RCK',
          skus: data.topSkusRck,
          totalCost: data.topSkusRck.reduce((s, x) => s + (x.cost || 0), 0),
          totalLpns: data.topSkusRck.reduce((s, x) => s + (x.lpns || 0), 0),
          totalBultos: data.topSkusRck.reduce((s, x) => s + (x.bultos || 0), 0)
        });
      }
      if (data.topSkusRhb && data.topSkusRhb.length > 0) {
        zonesWithData.push({
          zone: 'RHB',
          skus: data.topSkusRhb,
          totalCost: data.topSkusRhb.reduce((s, x) => s + (x.cost || 0), 0),
          totalLpns: data.topSkusRhb.reduce((s, x) => s + (x.lpns || 0), 0),
          totalBultos: data.topSkusRhb.reduce((s, x) => s + (x.bultos || 0), 0)
        });
      }
    }

    if (zonesWithData.length === 0) {
      containerTopZones.style.gridTemplateColumns = '1fr';
      containerTopZones.innerHTML = `
        <div style="background:#ffffff; border:1px solid #cbd5e1; border-radius:8px; padding:20px; text-align:center; color:#64748b;">
          <i class="fa-solid fa-circle-check text-emerald-500" style="font-size:1.5rem; margin-bottom:6px;"></i>
          <div style="font-weight:700;">No se registran SKUs con antigüedad mayor a 52 semanas en ninguna zona operativa</div>
        </div>
      `;
    } else {
      // Ajustar dinámicamente columnas según cuántas zonas tienen información:
      // Si 1 zona -> 1 columna
      // Si 2 zonas -> 2 columnas (1fr 1fr)
      // Si 3 zonas -> 3 columnas (1fr 1fr 1fr)
      if (zonesWithData.length === 1) {
        containerTopZones.style.gridTemplateColumns = '1fr';
      } else if (zonesWithData.length === 2) {
        containerTopZones.style.gridTemplateColumns = 'repeat(2, 1fr)';
      } else if (zonesWithData.length === 3) {
        containerTopZones.style.gridTemplateColumns = 'repeat(3, 1fr)';
      } else {
        containerTopZones.style.gridTemplateColumns = 'repeat(auto-fit, minmax(360px, 1fr))';
      }

      zonesWithData.forEach(zItem => {
        containerTopZones.innerHTML += renderZoneTopCardHtml(zItem);
      });
    }
  }
}

function renderZoneTopCardHtml(zData) {
  const zone = zData.zone || 'ZONA';
  const safeZone = escapeHtml(zone);
  const skus = zData.skus || [];

  let rowsHtml = '';
  let sumCost = 0, sumLpn = 0, sumBultos = 0;

  if (skus.length === 0) {
    rowsHtml = `<tr><td colspan="6" style="text-align:center; padding:12px; color:#64748b;">No hay SKUs &gt;52 semanas en ${safeZone}</td></tr>`;
  } else {
    skus.forEach((s, idx) => {
      sumCost += (s.cost || 0);
      sumLpn += (s.lpns || 0);
      sumBultos += (s.bultos || 0);

      const bg = (idx % 2 === 1) ? '#f4faf6' : '#ffffff';
      const safeDesc = escapeHtml(s.desc);
      rowsHtml += `
        <tr style="background: ${bg}; font-size: 0.78rem;">
          <td style="font-family:'JetBrains Mono',monospace; font-weight:800; color:#0f172a; text-align:left; padding:4.5px 6px;">${escapeHtml(s.sku)}</td>
          <td style="text-align:left; font-weight:700; color:#1e293b; max-width:210px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; padding:4.5px 6px;" title="${safeDesc}">${safeDesc}</td>
          <td style="text-align:center; color:#64748b; font-size:0.74rem; padding:4.5px 4px;">${escapeHtml(s.fv || '-')}</td>
          <td style="text-align:right; font-weight:800; color:#0f172a; padding:4.5px 6px;">${formatNumber(s.lpns)}</td>
          <td style="text-align:right; font-weight:800; color:${s.cost > 20000 ? '#dc2626' : '#0f172a'}; padding:4.5px 6px;">${formatNumber(s.cost)}</td>
          <td style="text-align:right; font-weight:600; padding:4.5px 6px;">${formatNumber(s.bultos)}</td>
        </tr>
      `;
    });
  }

  const finalLpns = zData.totalLpns !== undefined ? zData.totalLpns : sumLpn;
  const finalCost = zData.totalCost !== undefined ? zData.totalCost : sumCost;
  const finalBultos = zData.totalBultos !== undefined ? zData.totalBultos : sumBultos;

  return `
    <div style="background: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px; padding: 6px 8px; overflow-x: auto; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
      <div style="background: #2d6a4f; color: #ffffff; text-align: center; font-weight: 900; font-size: 0.90rem; padding: 5px; border-radius: 5px 5px 0 0; margin-bottom: 4px; letter-spacing: 0.5px;">
        ${safeZone}
      </div>
      <table class="exec-table exec-table-sm exec-table-green" style="width: 100%; border-collapse: collapse;">
        <thead>
          <tr style="background: #2d6a4f !important;">
            <th style="width: 70px; text-align: left; background: #2d6a4f !important; color: #ffffff !important; font-weight: 800; font-size: 0.78rem; padding: 5px 6px;">SKU</th>
            <th style="text-align: left; background: #2d6a4f !important; color: #ffffff !important; font-weight: 800; font-size: 0.78rem; padding: 5px 6px;">DESCRIPCION</th>
            <th style="text-align: center; font-size: 0.74rem; background: #2d6a4f !important; color: #ffffff !important; font-weight: 800; padding: 5px 4px;">RNG-FV</th>
            <th style="text-align: right; background: #2d6a4f !important; color: #ffffff !important; font-weight: 800; font-size: 0.78rem; padding: 5px 6px;">LPN's</th>
            <th style="text-align: right; background: #2d6a4f !important; color: #ffffff !important; font-weight: 800; font-size: 0.78rem; padding: 5px 6px;">COSTOS</th>
            <th style="text-align: right; background: #2d6a4f !important; color: #ffffff !important; font-weight: 800; font-size: 0.78rem; padding: 5px 6px;">BULTOS</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
        </tbody>
        <tfoot>
          <tr style="background: #f8fafc; font-weight: 900; border-top: 2px solid #cbd5e1; font-size: 0.82rem;">
            <td colspan="3" style="text-align: left; font-weight: 900; color: #0f172a; padding: 5px 6px;">Total general</td>
            <td style="text-align: right; color: #0f172a; font-weight: 900; padding: 5px 6px;">${formatNumber(finalLpns)}</td>
            <td style="text-align: right; color: #dc2626; font-weight: 900; padding: 5px 6px;">${formatNumber(finalCost)}</td>
            <td style="text-align: right; font-weight: 900; color: #0f172a; padding: 5px 6px;">${formatNumber(finalBultos)}</td>
          </tr>
        </tfoot>
      </table>
    </div>
  `;
}

function renderSlide3(data) {
  renderChartS3(data.weeks, data.evolution);

  const theadRow = document.getElementById('theadS3EvolRow');
  theadRow.innerHTML = `<th style="text-align:left; font-weight:900; padding:4px 6px; font-size:0.78rem;">Rango Semanas</th>`;
  (data.weeks || []).forEach((w, idx) => {
    const isCurrent = idx === data.weeks.length - 1;
    theadRow.innerHTML += `<th style="${isCurrent ? 'background:#eff6ff; color:#1d4ed8; font-weight:900; font-size:0.78rem; padding:4px 3px;' : 'font-weight:800; font-size:0.76rem; padding:4px 3px;'}">${escapeHtml(w)}</th>`;
  });
  theadRow.innerHTML += `<th style="font-weight:900; font-size:0.78rem; padding:4px 4px;">Variación</th>`;

  const tbodyEvol = document.getElementById('tbodyS3Evolution');
  tbodyEvol.innerHTML = '';
  (data.evolution || []).forEach(e => {
    const vals = e.values || [];
    const lastVal = vals[vals.length - 1] || 0;
    const prevVal = vals[vals.length - 2] || lastVal;
    const diff = lastVal - prevVal;
    const isNeutral = Math.abs(diff) < 0.005;
    const isPositiveGood = e.label.includes('0 a 10') ? diff > 0 : diff < 0;

    let trHtml = `
      <td style="text-align:left; font-weight:800; color:#1e293b; font-size:0.78rem; padding:4px 6px; white-space:nowrap;">
        <span style="display:inline-block; width:7px; height:7px; border-radius:50%; background:${escapeHtml(e.color)}; margin-right:5px;"></span>
        ${escapeHtml(e.label)}
      </td>
    `;
    vals.forEach((v, idx) => {
      const isCurrent = idx === vals.length - 1;
      trHtml += `<td style="${isCurrent ? 'background:#eff6ff; font-weight:900; font-size:0.80rem; color:' + e.color + '; padding:4px 3px;' : 'font-size:0.76rem; font-weight:600; padding:4px 3px;'}">${safeFixed(v, 2, '%')}</td>`;
    });

    trHtml += `
      <td style="font-weight:800; font-size:0.76rem; padding:4px 4px; white-space:nowrap; color:${isNeutral ? '#64748b' : (isPositiveGood ? '#059669' : '#dc2626')};">
        ${isNeutral ? '= 0.00%' : (diff > 0 ? '▲ +' : '▼ -') + safeFixed(Math.abs(diff), 2, '%')}
      </td>
    `;

    const tr = document.createElement('tr');
    tr.innerHTML = trHtml;
    tbodyEvol.appendChild(tr);
  });

  // Ubicaciones
  document.getElementById('s3-total-ubic').textContent = `${formatNumber(data.locationsTotal)} ubicaciones`;
  const tbodyLoc = document.getElementById('tbodyS3Locations');
  tbodyLoc.innerHTML = '';
  let sum010 = 0, sum1025 = 0, sum2552 = 0, sum52 = 0;

  (data.locations || []).forEach(loc => {
    sum010 += (loc.r010 || 0);
    sum1025 += (loc.r1025 || 0);
    sum2552 += (loc.r2552 || 0);
    sum52 += (loc.r52 || 0);

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="font-weight:700; color:#0f172a;">${escapeHtml(loc.div)}</td>
      <td style="font-weight:600;">${formatNumber(loc.r010)}</td>
      <td style="${loc.r1025 > 500 ? 'background:#fffbeb; font-weight:700; color:#b45309;' : 'font-weight:600;'}">${formatNumber(loc.r1025)}</td>
      <td style="${loc.r2552 > 100 ? 'background:#ffedd5; font-weight:800; color:#c2410c;' : 'font-weight:600;'}">${formatNumber(loc.r2552)}</td>
      <td style="${loc.r52 > 10 ? 'background:#fee2e2; font-weight:900; color:#dc2626;' : 'font-weight:600;'}">${loc.r52 > 0 ? formatNumber(loc.r52) : '-'}</td>
      <td style="font-weight:800; color:#0f172a;">${formatNumber(loc.total)}</td>
      <td style="color:#2563eb; font-weight:800;">${safeFixed(loc.pct, 1, '%')}</td>
    `;
    tbodyLoc.appendChild(tr);
  });

  document.getElementById('tfootS3Locations').innerHTML = `
    <tr>
      <td style="text-align:left; font-weight:900;">Total Ubicaciones</td>
      <td style="font-weight:900;">${formatNumber(sum010)}</td>
      <td style="font-weight:900;">${formatNumber(sum1025)}</td>
      <td style="color:#ea580c; font-weight:900;">${formatNumber(sum2552)}</td>
      <td style="color:#dc2626; font-weight:900;">${formatNumber(sum52)}</td>
      <td style="color:#0f172a; font-weight:900;">${formatNumber(data.locationsTotal)}</td>
      <td style="color:#2563eb; font-weight:900;">100%</td>
    </tr>
  `;
}

function renderChartS3(weeks, evolution) {
  const ctx = document.getElementById('chartS3Evolution');
  if (!ctx) return;

  if (chartInstanceS3) {
    chartInstanceS3.destroy();
  }

  const wList = weeks || [];
  const datasets = (evolution || []).map(e => ({
    label: e.label,
    data: e.values || [],
    borderColor: e.color,
    backgroundColor: e.color,
    borderWidth: e.label.includes('0 a 10') ? 2.5 : 2,
    tension: 0.25,
    pointRadius: 3,
    pointHoverRadius: 5
  }));

  chartInstanceS3 = new Chart(ctx, {
    type: 'line',
    data: {
      labels: wList,
      datasets: datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'top',
          labels: { boxWidth: 12, font: { size: 11, weight: '700' } }
        },
        tooltip: {
          callbacks: {
            label: (item) => ` ${item.dataset.label}: ${safeFixed(item.raw, 2, '%')} (${safeFixed(item.raw, 4, '%')})`
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          max: 100,
          ticks: {
            callback: (v) => `${v}%`,
            font: { size: 11, weight: '700' }
          },
          grid: { color: '#f1f5f9' }
        },
        x: {
          grid: { display: false },
          ticks: { font: { size: 11, weight: '700' } }
        }
      }
    }
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 10. MOTOR DE COPIA 16:9 NATIVO DIRECTO AL PORTAPAPELES (SIN MODAL NI DESCARGAS)
// ══════════════════════════════════════════════════════════════════════════════

async function generateCurrentSlideBlob() {
  const containerId = `slide${currentSlide}-container`;
  const container = document.getElementById(containerId);
  if (!container) throw new Error('Contenedor de diapositiva no encontrado');

  // 1. Aplicar clase que bloquea dimensiones estrictas 16:9 (1600x900)
  container.classList.add('capturing-16-9');
  if (currentSlide === 1 && chartInstanceS1) chartInstanceS1.resize();
  if (currentSlide === 3 && chartInstanceS3) chartInstanceS3.resize();
  await yieldToUi();

  let renderedCanvas;
  try {
    // 2. Renderizar con html2canvas en alta definición capturando todo el scrollHeight
    renderedCanvas = await window.html2canvas(container, {
      scale: 1.5, // Resolución nítida Retina
      backgroundColor: '#ffffff',
      useCORS: true,
      logging: false,
      scrollX: 0,
      scrollY: 0,
      windowWidth: 1600,
      height: container.scrollHeight
    });
  } finally {
    // Restaurar vista web responsiva inmediatamente
    container.classList.remove('capturing-16-9');
    if (currentSlide === 1 && chartInstanceS1) chartInstanceS1.resize();
    if (currentSlide === 3 && chartInstanceS3) chartInstanceS3.resize();
  }

  // 3. Crear canvas panorámico exacto 1920x1080 (16:9)
  const W = 1920;
  const H = 1080;
  const offscreen = document.createElement('canvas');
  offscreen.width = W;
  offscreen.height = H;
  const ctx = offscreen.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, W, H);

  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(renderedCanvas, 0, 0, W, H);

  const blob = await new Promise((resolve, reject) => {
    offscreen.toBlob((b) => {
      if (b) resolve(b);
      else reject(new Error('Error al crear imagen PNG de la diapositiva'));
    }, 'image/png');
  });

  lastCapturedBlob = blob;
  return blob;
}

async function copyCurrentSlideToClipboard() {
  if (isCapturing) return;
  isCapturing = true;

  const btn = document.getElementById('btnCapturePpt');
  const originalText = btn ? btn.innerHTML : '';
  if (btn) btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Copiando...';

  showToast(`📸 Copiando diapositiva 16:9 de Lámina ${currentSlide}...`, 'info', 2000);

  try {
    if (window.focus) window.focus();

    // Iniciar generación del Blob PNG inmediatamente
    const blobPromise = generateCurrentSlideBlob();
    let copied = false;

    if (navigator.clipboard && window.ClipboardItem) {
      try {
        // Registrar ClipboardItem síncronamente con la Promesa para preservar el gesto de usuario (User Activation)
        // aunque html2canvas tarde más de 1 segundo en renderizar la lámina
        const item = new ClipboardItem({ 'image/png': blobPromise });
        await navigator.clipboard.write([item]);
        copied = true;
      } catch (asyncClipErr) {
        // Respaldo para navegadores que requieren el Blob ya resuelto
        try {
          const blob = await blobPromise;
          if (window.focus) window.focus();
          const item = new ClipboardItem({ 'image/png': blob });
          await navigator.clipboard.write([item]);
          copied = true;
        } catch (syncClipErr) {
          console.warn('Error al escribir imagen en el portapapeles:', syncClipErr);
        }
      }
    }

    if (copied) {
      showToast(`✅ ¡Lámina ${currentSlide} copiada al portapapeles! Lista para pegar en PowerPoint (Ctrl + V)`, 'success', 4000);
      if (btn) {
        btn.innerHTML = '<i class="fa-solid fa-check"></i> ¡Copiada!';
        setTimeout(() => { if (btn) btn.innerHTML = originalText; }, 2200);
      }
    } else {
      if (!document.hasFocus()) {
        showToast('⚠️ La ventana perdió el foco antes de terminar el copiado. Mantén la ventana abierta 1 segundo al presionar Copiar.', 'warning', 5500);
      } else {
        showToast('No se pudo copiar la imagen al portapapeles. Vuelve a intentarlo.', 'danger', 4000);
      }
      if (btn) btn.innerHTML = originalText;
    }

  } catch (err) {
    console.error('Error al capturar diapositiva:', err);
    showToast('No se pudo generar la diapositiva', 'danger');
    if (btn) btn.innerHTML = originalText;
  } finally {
    setTimeout(() => { isCapturing = false; }, 300);
  }
}

function showCopySuccessModal() {
  // Deshabilitado permanentemente para copiar directo sin pasos extra (como Throughput)
}

function closeCopyModal() {
  const modal = document.getElementById('modalCopySuccess');
  if (modal) modal.style.display = 'none';
}

async function copyModalImgAgain() {
  if (!lastCapturedBlob) return;
  try {
    const item = new ClipboardItem({ 'image/png': lastCapturedBlob });
    await navigator.clipboard.write([item]);
    showToast('✅ ¡Copiado nuevamente al portapapeles!', 'success', 3000);
  } catch (err) {
    showToast('Por favor pulsa botón derecho sobre la imagen y selecciona "Copiar imagen"', 'info', 4500);
  }
}

function downloadModalImg() {
  if (!lastCapturedDataUrl) return;
  const link = document.createElement('a');
  link.download = `Antiguedad_Inventario_CD${currentWarehouse}_Lamina${currentSlide}_16x9.png`;
  link.href = lastCapturedDataUrl;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

async function downloadCurrentSlide() {
  const containerId = `slide${currentSlide}-container`;
  const container = document.getElementById(containerId);
  if (!container) return;

  showToast(`💾 Generando archivo PNG 16:9 (1920x1080)...`, 'info', 2000);

  container.classList.add('capturing-16-9');
  await yieldToUi();

  const renderedCanvas = await window.html2canvas(container, {
    scale: 1.5,
    backgroundColor: '#ffffff',
    useCORS: true,
    logging: false,
    scrollX: 0,
    scrollY: 0,
    windowWidth: 1600,
    height: container.scrollHeight
  });

  container.classList.remove('capturing-16-9');

  const W = 1920;
  const H = 1080;
  const offscreen = document.createElement('canvas');
  offscreen.width = W;
  offscreen.height = H;
  const ctx = offscreen.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, W, H);
  ctx.drawImage(renderedCanvas, 0, 0, W, H);

  const link = document.createElement('a');
  link.download = `Antiguedad_Inventario_CD${currentWarehouse}_Lamina${currentSlide}_16x9.png`;
  link.href = offscreen.toDataURL('image/png');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
