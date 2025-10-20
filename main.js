


    const filesInput = document.getElementById('files');
    const mergeBtn = document.getElementById('mergeBtn');
    const previewBtn = document.getElementById('previewBtn');
    const alerts = document.getElementById('alerts');
    const preview = document.getElementById('preview');

    function showAlert(msg, type='danger') {
      alerts.innerHTML = `<div class="alert alert-${type}" role="alert">${msg}</div>`;
      setTimeout(()=> { if (alerts.firstChild) alerts.removeChild(alerts.firstChild); }, 8000);
    }

    function readFileAsArrayBuffer(file) {
      return new Promise((res, rej) => {
        const reader = new FileReader();
        reader.onload = e => res(e.target.result);
        reader.onerror = e => rej(e);
        reader.readAsArrayBuffer(file);
      });
    }

    async function parseWorkbook(file) {
      const ab = await readFileAsArrayBuffer(file);
      const wb = XLSX.read(ab, {type:'array'});
      // Usamos la primera hoja
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      // Convertir a JSON con encabezados
      const json = XLSX.utils.sheet_to_json(ws, {header:1, defval: ''});
      return {fileName: file.name, sheetName, json};
    }

    function arrayFirstRowToHeader(row) {
      return row.map(h => String(h).trim());
    }

    function arraysEqual(a,b) {
      if (!a || !b) return false;
      if (a.length !== b.length) return false;
      for (let i=0;i<a.length;i++){
        if (String(a[i]) !== String(b[i])) return false;
      }
      return true;
    }

    previewBtn.addEventListener('click', async () => {
      preview.innerHTML = '';
      const files = [...filesInput.files];
      if (!files.length) { showAlert('No seleccionaste archivos.', 'warning'); return; }
      preview.innerHTML = `<div class="spinner-border" role="status"><span class="visually-hidden">Cargando...</span></div> Cargando...`;
      try {
        const parsed = await Promise.all(files.map(f => parseWorkbook(f)));
        preview.innerHTML = '';
        parsed.forEach(p => {
          const headers = p.json[0] ? arrayFirstRowToHeader(p.json[0]) : [];
          const table = document.createElement('div');
          table.innerHTML = `
            <div class="card mb-2">
              <div class="card-body">
                <h6>${p.fileName} — Hoja: ${p.sheetName}</h6>
                <p>Encabezados: <code>${headers.join(' | ')}</code></p>
                <p>Filas (incluye encabezado): ${p.json.length}</p>
              </div>
            </div>`;
          preview.appendChild(table);
        });
      } catch(err) {
        console.error(err);
        showAlert('Error al leer archivos. Revisá que sean archivos Excel válidos.');
        preview.innerHTML = '';
      }
    });

    mergeBtn.addEventListener('click', async () => {
      const files = [...filesInput.files];
      if (!files.length) { showAlert('No seleccionaste archivos.', 'warning'); return; }

      mergeBtn.disabled = true;
      mergeBtn.innerText = 'Procesando...';

      try {
        const parsed = await Promise.all(files.map(f => parseWorkbook(f)));
        // obtener encabezado de referencia (primer archivo)
        const firstJson = parsed[0].json;
        if (!firstJson || firstJson.length === 0) {
          showAlert('El primer archivo está vacío.', 'danger');
          mergeBtn.disabled = false; mergeBtn.innerText = 'Unir y descargar';
          return;
        }
        const headerRef = arrayFirstRowToHeader(firstJson[0]);

        // verificar que todos tengan el mismo header
        for (let i=0;i<parsed.length;i++){
          const p = parsed[i];
          const header = p.json[0] ? arrayFirstRowToHeader(p.json[0]) : [];
          if (!arraysEqual(headerRef, header)) {
            showAlert(`Estructura diferente en "${p.fileName}". Encabezados no coinciden.`, 'danger');
            mergeBtn.disabled = false; mergeBtn.innerText = 'Unir y descargar';
            return;
          }
        }

        // Construir matriz resultante: encabezado + columna source_file
        const outHeader = [...headerRef, 'source_file'];
        const outRows = [outHeader];

        // Para cada archivo, iterar filas (saltamos la fila 0 que es encabezado)
        parsed.forEach(p => {
          const rows = p.json.slice(1); // filas sin encabezado
          rows.forEach(r => {
            // Asegurar que r tenga la misma longitud que headerRef (completar si hace falta)
            const rowFixed = headerRef.map((_, idx) => (typeof r[idx] !== 'undefined' ? r[idx] : ''));
            rowFixed.push(p.fileName);
            outRows.push(rowFixed);
          });
        });

        // Crear libro y hoja
        const ws = XLSX.utils.aoa_to_sheet(outRows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Unificado');

        // Generar archivo para descarga
        const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
        const blob = new Blob([wbout], {type: 'application/octet-stream'});

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