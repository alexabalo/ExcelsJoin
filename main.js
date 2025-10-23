// === selectores DOM ===
const filesInput = document.getElementById('files');
const mergeBtn = document.getElementById('mergeBtn');
const previewBtn = document.getElementById('previewBtn');
const alerts = document.getElementById('alerts');
const preview = document.getElementById('preview');

// === util: mostrar alertas ===
function showAlert(msg, type = 'danger') {
  alerts.innerHTML = `<div class="alert alert-${type}" role="alert">${msg}</div>`;
  setTimeout(() => { if (alerts.firstChild) alerts.removeChild(alerts.firstChild); }, 8000);
}

// === util: leer archivo como ArrayBuffer ===
function readFileAsArrayBuffer(file) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => res(e.target.result);
    reader.onerror = e => rej(e);
    reader.readAsArrayBuffer(file);
  });
}

// === parse workbook a json (header:1 para obtener filas como arrays) ===
async function parseWorkbook(file) {
  const ab = await readFileAsArrayBuffer(file);
  const wb = XLSX.read(ab, { type: 'array' });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return { fileName: file.name, sheetName, json };
}

// -------------- util: normalizar celda de encabezado ---------------
function normalizeHeaderCell(v) {
  if (v === null || typeof v === 'undefined') return '';
  let s = String(v).replace(/^\uFEFF/, '').trim();
  // colapsar espacios y quitar acentos y caracteres combinantes
  s = s.replace(/\s+/g, ' ').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  // eliminar caracteres no imprimibles extremos
  s = s.replace(/[^\x20-\x7EÁÉÍÓÚáéíóúÑñüÜ]/g, '');
  return s.toLowerCase();
}

// -------------- convertir primera fila a encabezados (más robusto) ---------------
function arrayFirstRowToHeader(row) {
  // row: array de valores crudos
  if (!Array.isArray(row)) return [];
  // 1) convertir a strings normalizados
  const tmp = row.map(h => normalizeHeaderCell(h));
  // 2) propagar el último valor no vacío hacia adelante (para merged cells)
  let last = '';
  for (let i = 0; i < tmp.length; i++) {
    if (tmp[i] && tmp[i].length > 0) {
      last = tmp[i];
    } else {
      tmp[i] = last || '';
    }
  }
  // 3) recortar columnas vacías al final
  let lastNonEmpty = -1;
  for (let i = 0; i < tmp.length; i++) if (tmp[i] && tmp[i].length) lastNonEmpty = i;
  const clean = lastNonEmpty === -1 ? [] : tmp.slice(0, lastNonEmpty + 1);

  // 4) si quedan nombres vacíos en medio, asignar etiquetas únicas col_n
  return clean.map((h, idx) => (h && h.length ? h : `col_${idx + 1}`));
}

// -------------- heurística de fila de encabezado mejorada ---------------
function findHeaderRowIndex(jsonRows, referenceHeaderNormalized) {
  if (!Array.isArray(jsonRows) || jsonRows.length === 0) return 0;
  // revisar primeras 8 filas en busca de la fila con mayoría de celdas no vacías
  for (let i = 0; i < Math.min(8, jsonRows.length); i++) {
    const row = jsonRows[i];
    if (!row) continue;
    const nonEmptyCount = row.reduce((acc, c) => acc + (c !== null && c !== undefined && String(c).trim() !== '' ? 1 : 0), 0);
    if (nonEmptyCount >= 2) return i;
  }
  // si no encontramos y hay referencia, buscar fila con mayor coincidencia de nombres
  if (referenceHeaderNormalized && Array.isArray(referenceHeaderNormalized)) {
    for (let i = 0; i < Math.min(12, jsonRows.length); i++) {
      const row = jsonRows[i] || [];
      const norm = row.map(r => normalizeHeaderCell(r));
      const matches = norm.reduce((acc, v) => acc + (referenceHeaderNormalized.includes(v) ? 1 : 0), 0);
      if (matches >= Math.max(1, Math.floor(referenceHeaderNormalized.length * 0.4))) return i;
    }
  }
  return 0;
}

// === Preview: muestra info básica por archivo ===
previewBtn.addEventListener('click', async () => {
  preview.innerHTML = '';
  const files = [...filesInput.files];
  if (!files.length) { showAlert('No seleccionaste archivos.', 'warning'); return; }
  preview.innerHTML = `<div class="spinner-border" role="status"><span class="visually-hidden">Cargando...</span></div> Cargando...`;
  try {
    const parsed = await Promise.all(files.map(f => parseWorkbook(f)));
    preview.innerHTML = '';
    parsed.forEach(p => {
      const headerRowIdx = findHeaderRowIndex(p.json);
      const headerRow = p.json[headerRowIdx] || [];
      const headers = headerRow.length ? arrayFirstRowToHeader(headerRow) : [];
      const table = document.createElement('div');
      table.innerHTML = `
            <div class="card mb-2">
              <div class="card-body">
                <h6>${p.fileName} — Hoja: ${p.sheetName}</h6>
                <p>Fila de encabezado detectada: ${headerRowIdx + 1}</p>
                <p>Encabezados (normalizados): <code>${headers.join(' | ')}</code></p>
                <p>Filas totales (incluye encabezados y posibles filas vacías): ${p.json.length}</p>
              </div>
            </div>`;
      preview.appendChild(table);

      // Logueo adicional en consola para debug
      console.log(`Preview: ${p.fileName}`, {
        headerRowIdx,
        headerRowPreview: headerRow.slice(0, 30),
        normalized: headers,
        firstRows: p.json.slice(0, 8)
      });
    });
  } catch (err) {
    console.error(err);
    showAlert('Error al leer archivos. Revisá que sean archivos Excel válidos.');
    preview.innerHTML = '';
  }
});

// -------------- merge tolerante: reemplaza la sección que valida y construye outRows ---------------
mergeBtn.addEventListener('click', async () => {
  const files = [...filesInput.files];
  if (!files.length) { showAlert('No seleccionaste archivos.', 'warning'); return; }

  mergeBtn.disabled = true;
  mergeBtn.innerText = 'Procesando...';
  try {
    const parsed = await Promise.all(files.map(f => parseWorkbook(f)));

    // referencia: primer archivo
    const first = parsed[0];
    const firstHeaderRowIdx = findHeaderRowIndex(first.json);
    const firstHeaderRow = first.json[firstHeaderRowIdx] || [];
    const headerRef = arrayFirstRowToHeader(firstHeaderRow); // normalizado y limpio (referencia)
    const headerOriginal = headerRef.slice(); // usaremos estos nombres limpios en el XLSX final

    if (!headerRef || headerRef.length === 0) {
      showAlert('El primer archivo no tiene encabezados detectables.', 'danger');
      mergeBtn.disabled = false; mergeBtn.innerText = 'Unir y descargar';
      return;
    }

    // Validar archivos: permitir distinto orden y columnas extras/faltantes.
    for (let i = 0; i < parsed.length; i++) {
      const p = parsed[i];
      const headerRowIdx = findHeaderRowIndex(p.json, headerRef);
      const headerRow = p.json[headerRowIdx] || [];
      const header = arrayFirstRowToHeader(headerRow);

      // calcular faltantes y extras respecto a headerRef
      const missing = headerRef.filter(h => !header.includes(h));
      const extra = header.filter(h => !headerRef.includes(h));

      if (missing.length && (missing.length / headerRef.length) > 0.5) {
        console.error('Archivo con muchas columnas faltantes:', p.fileName, { missing, extra, headerRow, header });
        showAlert(`Estructura muy diferente en "${p.fileName}". Muchas columnas faltan.`, 'danger');
        mergeBtn.disabled = false; mergeBtn.innerText = 'Unir y descargar';
        return;
      }
      if (extra.length) {
        console.warn(`Archivo "${p.fileName}" tiene columnas extra:`, extra);
      }
    }

    // Preparar salida: encabezado referencia + source_file
    const outHeader = [...headerOriginal, 'source_file'];
    const outRows = [outHeader];

    // Por cada archivo, crear mapa nombre->índice y rellenar filas con orden de headerRef
    parsed.forEach(p => {
      const headerRowIdx = findHeaderRowIndex(p.json, headerRef);
      const headerRow = p.json[headerRowIdx] || [];
      const header = arrayFirstRowToHeader(headerRow);

      const idxByName = {};
      header.forEach((h, idx) => { if (h && h.length) idxByName[h] = idx; });

      const dataRows = p.json.slice(headerRowIdx + 1);
      dataRows.forEach(r => {
        const nonEmpty = r && r.some(c => c !== null && c !== undefined && String(c).trim() !== '');
        if (!nonEmpty) return;
        const rowFixed = headerRef.map(h => {
          if (idxByName.hasOwnProperty(h)) {
            return typeof r[idxByName[h]] !== 'undefined' ? r[idxByName[h]] : '';
          } else {
            return '';
          }
        });
        rowFixed.push(p.fileName);
        outRows.push(rowFixed);
      });
    });

    // Crear libro y descargar
    const ws = XLSX.utils.aoa_to_sheet(outRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Unificado');
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    const a = document.createElement('a');
    const url = URL.createObjectURL(blob);
    a.href = url;
    a.download = 'unificado.xlsx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    showAlert('Unificación completada. Archivo descargado.', 'success');
  } catch (err) {
    console.error(err);
    showAlert('Error durante el proceso: ' + (err && err.message ? err.message : err));
  } finally {
    mergeBtn.disabled = false;
    mergeBtn.innerText = 'Unir y descargar';
  }
});
