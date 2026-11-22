document.addEventListener('DOMContentLoaded', () => {
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input');
  const copyBtn = document.getElementById('copy-btn');

  if (dropZone && fileInput) {
    dropZone.addEventListener('click', () => {
      fileInput.click();
    });
  }

  if (fileInput) {
    fileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;

      const MAX_FILE_BYTES = 50 * 1024 * 1024;
      if (!/\.(xlsx|xls|xlsm|xlsb)$/i.test(file.name || '')) {
        alert('Formato no compatible. Selecciona un archivo Excel (.xlsx, .xls, .xlsm, .xlsb).');
        return;
      }
      if (file.size > MAX_FILE_BYTES) {
        alert('El archivo supera el límite seguro de 50 MB.');
        return;
      }

      if (dropZone) {
        dropZone.textContent = '⏳ Procesando archivo de forma completa...';
      }

      const reader = new FileReader();
      reader.onload = function (loadEvent) {
        try {
          const data = new Uint8Array(loadEvent.target.result);
          const workbook = XLSX.read(data, { type: 'array', cellFormula: true });

          let outputText = '=== RESUMEN DEL ARCHIVO: ' + file.name + ' ===\n\n';

          workbook.SheetNames.forEach(sheetName => {
            outputText += '========================================\n';
            outputText += 'HOJA: ' + sheetName + '\n';
            outputText += '========================================\n';

            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet['!ref']) {
              outputText += 'Hoja vacía.\n\n';
              return;
            }

            const rowsWithData = {};
            let totalFormulas = 0;

            for (const z in worksheet) {
              if (z[0] === '!') continue;

              const cell = worksheet[z];
              const decoded = XLSX.utils.decode_cell(z);

              if (!rowsWithData[decoded.r]) {
                rowsWithData[decoded.r] = [];
              }

              const colLetter = XLSX.utils.encode_col(decoded.c);
              let valStr = '';

              const formulaDict = {
                'IF': 'SI', 'IFERROR': 'SI.ERROR', 'VLOOKUP': 'BUSCARV',
                'SUM': 'SUMA', 'SUMIF': 'SUMAR.SI', 'SUMIFS': 'SUMAR.SI.CONJUNTO',
                'COUNT': 'CONTAR', 'COUNTA': 'CONTARA', 'COUNTIF': 'CONTAR.SI', 'COUNTIFS': 'CONTAR.SI.CONJUNTO',
                'AVERAGE': 'PROMEDIO', 'AVERAGEIF': 'PROMEDIO.SI', 'AVERAGEIFS': 'PROMEDIO.SI.CONJUNTO',
                'INDEX': 'INDICE', 'MATCH': 'COINCIDIR', 'SMALL': 'K.ESIMO.MENOR', 'LARGE': 'K.ESIMO.MAYOR',
                'ROUND': 'REDONDEAR', 'TEXT': 'TEXTO', 'AND': 'Y', 'OR': 'O', 'TODAY': 'HOY'
              };

              if (cell.f) {
                let f_es = cell.f;
                for (const [eng, spa] of Object.entries(formulaDict)) {
                  const regex = new RegExp('\\b' + eng + '\\b', 'gi');
                  f_es = f_es.replace(regex, spa);
                }
                f_es = f_es.replace(/,/g, ';');

                valStr = `[FÓRMULA: =${f_es}] (Valor: ${cell.v !== undefined ? cell.v : ''})`;
                totalFormulas++;
              } else {
                valStr = cell.v !== undefined ? cell.v : '';
              }

              rowsWithData[decoded.r].push(`${colLetter}: ${valStr}`);
            }

            const sortedRows = Object.keys(rowsWithData).map(Number).sort((a, b) => a - b);

            if (sortedRows.length === 0) {
              outputText += 'No se encontró información legible en esta hoja.\n\n';
              return;
            }

            outputText += `Total de filas con información: ${sortedRows.length}\n`;
            outputText += `Fórmulas detectadas: ${totalFormulas}\n\n`;
            outputText += '--- MAPA DE DATOS Y FÓRMULAS ---\n';

            const limit = Math.min(sortedRows.length, 3000);

            for (let i = 0; i < limit; i++) {
              const r = sortedRows[i];
              const rowNumber = r + 1;

              if (i > 0 && r - sortedRows[i - 1] > 1) {
                const salto = (r - sortedRows[i - 1]) - 1;
                outputText += `\n   ... [${salto} fila(s) en blanco] ...\n\n`;
              }

              outputText += `Fila ${rowNumber} | ` + rowsWithData[r].join(' | ') + '\n';
            }

            if (sortedRows.length > 3000) {
              outputText += `\n... Y ${sortedRows.length - 3000} filas adicionales omitidas por límite de seguridad.\n`;
            }

            outputText += '\n';
          });

          const outputEl = document.getElementById('output');
          if (outputEl) outputEl.value = outputText;

          if (copyBtn) copyBtn.style.display = 'inline-block';
          if (dropZone) {
            dropZone.textContent = '✅ ¡Escaneo Completo finalizado! Haz clic en Copiar.';
            dropZone.style.background = '#d1fae5';
            dropZone.style.borderColor = '#10b981';
            dropZone.style.color = '#047857';
          }
        } catch (err) {
          console.error('Error al leer Excel:', err);
          if (dropZone) dropZone.textContent = '❌ Error al leer el archivo Excel.';
          alert('No se pudo procesar el archivo Excel: ' + (err.message || 'archivo corrupto'));
        }
      };

      reader.readAsArrayBuffer(file);
    });
  }

  if (copyBtn) {
    copyBtn.addEventListener('click', async () => {
      const output = document.getElementById('output');
      if (!output) return;
      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(output.value);
        } else {
          output.select();
          document.execCommand('copy');
        }
        alert('¡Texto copiado!\n\nPégalo en el chat de la IA.');
      } catch (err) {
        output.select();
        document.execCommand('copy');
        alert('¡Texto copiado!\n\nPégalo en el chat de la IA.');
      }
    });
  }
});
