(function () {
  "use strict";

  const MAX_FILES = 5;
  const MIN_FILES = 2;
  let nextId = 3;

  const state = {
    mode: 'merge',      // 'merge' (Consolidar y sumar) | 'side' (Comparar en paralelo)
    keyField: null,     // Campo Clave — elegido una sola vez, común a todos los archivos
    valueField: null,   // Campo a Calcular — solo aplica en modo 'merge'
    files: [
      { id: 1, label: 'Archivo 1' },
      { id: 2, label: 'Archivo 2' }
    ]
  };

  let lastResult = null;   // { columns, data }
  let xlsxLoadPromise = null;

  const els = {
    modeCards: document.querySelectorAll('.mode-card'),
    fileSlots: document.getElementById('file-slots'),
    addFileBtn: document.getElementById('add-file-btn'),
    keyFields: document.getElementById('key-fields'),
    processBtn: document.getElementById('process-btn'),
    validationMsg: document.getElementById('validation-msg'),
    stepResults: document.getElementById('step-results'),
    stats: document.getElementById('stats'),
    resultTable: document.getElementById('result-table'),
    exportXlsx: document.getElementById('export-xlsx'),
    exportCsv: document.getElementById('export-csv'),
    copyBtn: document.getElementById('copy-btn'),
    sortField1: document.getElementById('sort-field-1'),
    sortDir1: document.getElementById('sort-dir-1'),
    sortField2: document.getElementById('sort-field-2'),
    sortDir2: document.getElementById('sort-dir-2'),
  };

  /* ================= Carga diferida de SheetJS ================= */
  // La librería (~250 KB gzip) solo se descarga cuando el usuario realmente
  // carga un archivo o pega datos, no durante la carga inicial de la página.
  function ensureXlsx() {
    if (window.XLSX) return Promise.resolve();
    if (xlsxLoadPromise) return xlsxLoadPromise;
    xlsxLoadPromise = new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = '/assets/vendor/xlsx.full.min.js';
      script.onload = () => resolve();
      script.onerror = () => reject(new Error('No se pudo cargar el motor de lectura de Excel.'));
      document.head.appendChild(script);
    });
    return xlsxLoadPromise;
  }

  /* ================= Parsing helpers ================= */

  function splitDelim(line, delim) {
    const result = []; let cur = ''; let q = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') { q = !q; continue; }
      if (c === delim && !q) { result.push(cur); cur = ''; }
      else cur += c;
    }
    result.push(cur);
    return result.map(s => s.trim());
  }

  function parseDelimitedText(text) {
    const lines = text.split(/\r\n|\n|\r/).filter(l => l.length > 0);
    if (lines.length === 0) return [];
    const delim = lines[0].includes('\t') ? '\t' : (lines[0].includes(';') && !lines[0].includes(',') ? ';' : ',');
    return lines.map(l => splitDelim(l, delim));
  }

  function aoaToHeadersRows(aoa) {
    aoa = aoa.filter(r => r.some(c => String(c).trim() !== ''));
    if (aoa.length === 0) return { headers: [], rows: [] };
    const headers = aoa[0].map((h, i) => String(h).trim() || ('Columna ' + (i + 1)));
    const rows = aoa.slice(1).map(r => {
      const o = {};
      headers.forEach((h, i) => { o[h] = r[i] !== undefined ? r[i] : ''; });
      return o;
    });
    return { headers, rows };
  }

  async function parseFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    let aoa;
    if (ext === 'csv' || ext === 'txt' || ext === 'tsv') {
      const text = await file.text();
      aoa = parseDelimitedText(text);
    } else {
      await ensureXlsx();
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    }
    return aoaToHeadersRows(aoa);
  }

  function parseNumber(v) {
    if (v === undefined || v === null || v === '') return 0;
    if (typeof v === 'number') return v;
    let s = String(v).trim();
    if (s === '') return 0;
    const hasComma = s.includes(','), hasDot = s.includes('.');
    if (hasComma && hasDot) {
      if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
      else s = s.replace(/,/g, '');
    } else if (hasComma) {
      s = s.replace(',', '.');
    }
    s = s.replace(/[^0-9.\-]/g, '');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  function normHeader(h) {
    return String(h == null ? '' : h).trim().toLowerCase();
  }

  function compareValues(a, b) {
    const na = parseFloat(a), nb = parseFloat(b);
    const aNum = a !== '' && a != null && !isNaN(na) && isFinite(na);
    const bNum = b !== '' && b != null && !isNaN(nb) && isFinite(nb);
    if (aNum && bNum) return na - nb;
    return String(a == null ? '' : a).localeCompare(String(b == null ? '' : b), 'es', { numeric: true, sensitivity: 'base' });
  }

  function sortByKeyAsc(arr, getter) {
    return arr.slice().sort((a, b) => compareValues(getter(a), getter(b)));
  }

  function roundSmart(n) {
    return Math.round(n * 1e6) / 1e6;
  }

  function escapeHtml(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function escapeAttr(v) {
    return escapeHtml(v).replace(/"/g, '&quot;');
  }

  /* ================= Mode cards ================= */

  els.modeCards.forEach(card => {
    card.addEventListener('click', () => {
      state.mode = card.dataset.mode;
      els.modeCards.forEach(c => c.classList.toggle('active', c === card));
      renderFileSlots();
      updateProcessButtonState();
    });
  });

  /* ================= File slots ================= */

  function slotTemplate(f) {
    const canRemove = state.files.length > MIN_FILES;
    const loaded = !!f.headers;
    let inner = '';
    if (!loaded) {
      inner = `
        <div class="dropzone" data-id="${f.id}" tabindex="0" aria-label="Zona de carga para ${f.label}">
          <input type="file" class="file-input" data-id="${f.id}" accept=".xlsx,.xls,.csv" />
          <span>Arrastra un archivo aquí, haz clic para elegirlo, o pega con Ctrl+V</span>
        </div>
        <button class="btn-outline paste-btn" data-id="${f.id}" type="button">Pegar desde portapapeles</button>
      `;
    } else {
      const previewRows = f.rows.slice(0, 4);
      const previewHead = f.headers.map(h => `<th>${escapeHtml(h)}</th>`).join('');
      const previewBody = previewRows.map(r =>
        `<tr>${f.headers.map(h => `<td>${escapeHtml(r[h])}</td>`).join('')}</tr>`
      ).join('');
      inner = `
        <div class="loaded-info">
          <span>${escapeHtml(f.fileName)} · ${f.rows.length} filas</span>
          <button class="clear-slot" data-id="${f.id}" type="button">Cambiar archivo</button>
        </div>
        <div class="preview-wrap">
          <table class="preview-table">
            <thead><tr>${previewHead}</tr></thead>
            <tbody>${previewBody}</tbody>
          </table>
        </div>
      `;
    }
    return `
      <div class="file-slot" data-id="${f.id}">
        <div class="slot-head">
          <input class="slot-label" data-id="${f.id}" value="${escapeAttr(f.label)}" aria-label="Nombre del archivo" />
          ${canRemove ? `<button class="remove-slot" data-id="${f.id}" type="button">Quitar</button>` : ''}
        </div>
        ${inner}
      </div>
    `;
  }

  function renderFileSlots() {
    els.fileSlots.innerHTML = state.files.map(slotTemplate).join('');
    els.addFileBtn.style.display = state.files.length >= MAX_FILES ? 'none' : 'inline-block';
  }

  els.fileSlots.addEventListener('click', (e) => {
    const rm = e.target.closest('.remove-slot');
    if (rm) {
      state.files = state.files.filter(f => f.id !== Number(rm.dataset.id));
      renderFileSlots(); updateProcessButtonState(); return;
    }
    const clr = e.target.closest('.clear-slot');
    if (clr) {
      const f = state.files.find(x => x.id === Number(clr.dataset.id));
      delete f.headers; delete f.rows; delete f.fileName;
      renderFileSlots(); updateProcessButtonState(); return;
    }
    const pasteBtn = e.target.closest('.paste-btn');
    if (pasteBtn) { handlePasteButton(Number(pasteBtn.dataset.id)); return; }
  });

  els.fileSlots.addEventListener('change', (e) => {
    if (e.target.matches('.file-input')) {
      const id = Number(e.target.dataset.id);
      const file = e.target.files[0];
      if (file) handleFileInput(id, file);
    }
  });

  els.fileSlots.addEventListener('input', (e) => {
    if (e.target.matches('.slot-label')) {
      const f = state.files.find(x => x.id === Number(e.target.dataset.id));
      f.label = e.target.value || f.label;
    }
  });

  els.fileSlots.addEventListener('dragover', (e) => {
    const dz = e.target.closest('.dropzone');
    if (dz) { e.preventDefault(); dz.classList.add('drag-over'); }
  });
  els.fileSlots.addEventListener('dragleave', (e) => {
    const dz = e.target.closest('.dropzone');
    if (dz) dz.classList.remove('drag-over');
  });
  els.fileSlots.addEventListener('drop', (e) => {
    const dz = e.target.closest('.dropzone');
    if (dz) {
      e.preventDefault(); dz.classList.remove('drag-over');
      const file = e.dataTransfer.files[0];
      if (file) handleFileInput(Number(dz.dataset.id), file);
    }
  });
  els.fileSlots.addEventListener('paste', (e) => {
    const dz = e.target.closest('.dropzone');
    if (!dz) return;
    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (!text) return;
    e.preventDefault();
    applyPastedText(Number(dz.dataset.id), text);
  });

  async function handleFileInput(id, file) {
    try {
      const { headers, rows } = await parseFile(file);
      if (headers.length === 0) { alert('El archivo no contiene datos legibles.'); return; }
      applyParsedData(id, file.name, headers, rows);
    } catch (err) {
      console.error(err);
      alert('No se pudo leer el archivo. Verifica que sea un .xlsx, .xls o .csv válido.');
    }
  }

  async function handlePasteButton(id) {
    try {
      const text = await navigator.clipboard.readText();
      if (!text || !text.trim()) { alert('El portapapeles está vacío.'); return; }
      applyPastedText(id, text);
    } catch (err) {
      alert('El navegador bloqueó la lectura automática del portapapeles. Haz clic dentro del recuadro de carga y pulsa Ctrl+V en su lugar.');
    }
  }

  function applyPastedText(id, text) {
    const aoa = parseDelimitedText(text);
    const { headers, rows } = aoaToHeadersRows(aoa);
    if (headers.length === 0) { alert('No se pudo interpretar el contenido pegado.'); return; }
    applyParsedData(id, 'Pegado desde portapapeles', headers, rows);
  }

  function applyParsedData(id, fileName, headers, rows) {
    const f = state.files.find(x => x.id === id);
    f.fileName = `${fileName} (${rows.length} filas)`;
    f.headers = headers;
    f.rows = rows;
    renderFileSlots();
    updateProcessButtonState();
  }

  els.addFileBtn.addEventListener('click', () => {
    if (state.files.length >= MAX_FILES) return;
    state.files.push({ id: nextId, label: `Archivo ${state.files.length + 1}` });
    nextId++;
    renderFileSlots(); updateProcessButtonState();
  });

  /* ================= Campo Clave / Campo a Calcular (elegidos una sola vez) ================= */

  function computeCommonHeaderDisplays(files) {
    if (files.length === 0) return [];
    const headerSets = files.map(f => new Set(f.headers.map(normHeader)));
    const commonNorms = [...headerSets[0]].filter(hn => headerSets.every(s => s.has(hn)));
    return commonNorms.map(hn => {
      for (const f of files) {
        const found = f.headers.find(h => normHeader(h) === hn);
        if (found !== undefined) return found;
      }
      return hn;
    });
  }

  function guessValueColFromCommon(common, files, keyDisplay) {
    const candidates = common.filter(h => h !== keyDisplay);
    if (candidates.length === 0) return null;
    const sample = files[0].rows.slice(0, 6);
    for (const h of candidates) {
      const numericCount = sample.filter(r => {
        const v = String(r[h]).trim();
        return v !== '' && !isNaN(parseFloat(v.replace(',', '.')));
      }).length;
      if (numericCount >= Math.max(1, Math.ceil(sample.length * 0.6))) return h;
    }
    return candidates[0];
  }

  function renderKeyFieldsBlock() {
    const container = els.keyFields;
    const loaded = state.files.filter(f => f.headers);
    if (state.files.length < MIN_FILES || loaded.length !== state.files.length) {
      container.innerHTML = `<p class="hint">Carga todos los archivos para elegir ${state.mode === 'merge' ? 'el Campo Clave y el Campo a Calcular' : 'el Campo Clave'}.</p>`;
      state.keyField = null; state.valueField = null;
      return;
    }

    const common = computeCommonHeaderDisplays(state.files);
    if (common.length === 0) {
      container.innerHTML = `<div class="warning-box">⚠ No es posible continuar: los archivos cargados no comparten ninguna columna con el mismo nombre. Revisa que las cabeceras coincidan exactamente en todos los archivos (por ejemplo, "Producto" en todos, no "Producto" y "Artículo").</div>`;
      state.keyField = null; state.valueField = null;
      return;
    }

    if (!common.includes(state.keyField)) state.keyField = common[0];

    let valueBlockHtml = '';
    if (state.mode === 'merge') {
      if (!common.includes(state.valueField) || state.valueField === state.keyField) {
        state.valueField = guessValueColFromCommon(common, state.files, state.keyField);
      }
      const valOptions = common.map(h =>
        `<option value="${escapeAttr(h)}" ${h === state.valueField ? 'selected' : ''}>${escapeHtml(h)}</option>`
      ).join('');
      valueBlockHtml = `
        <label>Campo a Calcular (común a todos los archivos)
          <select id="key-value-select">${valOptions}</select>
        </label>`;
    } else {
      state.valueField = null;
    }

    const keyOptions = common.map(h =>
      `<option value="${escapeAttr(h)}" ${h === state.keyField ? 'selected' : ''}>${escapeHtml(h)}</option>`
    ).join('');

    container.innerHTML = `
      <div class="col-selects">
        <label>Campo Clave (común a todos los archivos)
          <select id="key-key-select">${keyOptions}</select>
        </label>
        ${valueBlockHtml}
      </div>
      <p class="hint">Columnas comunes detectadas: ${common.map(escapeHtml).join(', ')}</p>
    `;
  }

  els.keyFields.addEventListener('change', (e) => {
    if (e.target.id === 'key-key-select') { state.keyField = e.target.value; updateProcessButtonState(); }
    else if (e.target.id === 'key-value-select') { state.valueField = e.target.value; updateProcessButtonState(); }
  });

  /* ================= Validación ================= */

  function updateProcessButtonState() {
    renderKeyFieldsBlock();
    const loaded = state.files.filter(f => f.headers);
    const allLoaded = state.files.length >= MIN_FILES && loaded.length === state.files.length;
    let ok = allLoaded;
    let msg = '';

    if (!allLoaded) {
      msg = 'Carga los archivos que falten antes de procesar.';
    } else {
      const common = computeCommonHeaderDisplays(state.files);
      if (common.length === 0) {
        ok = false;
        msg = 'No es posible continuar: faltan columnas con el mismo nombre en todos los archivos.';
      } else if (!state.keyField) {
        ok = false;
        msg = 'Elige el Campo Clave.';
      } else if (state.mode === 'merge' && !state.valueField) {
        ok = false;
        msg = 'Elige el Campo a Calcular.';
      } else if (state.mode === 'merge' && state.keyField === state.valueField) {
        ok = false;
        msg = 'El Campo Clave y el Campo a Calcular deben ser columnas distintas.';
      }
    }
    els.processBtn.disabled = !ok;
    els.validationMsg.textContent = ok ? '' : msg;
  }

  /* ================= Procesamiento: Consolidar y sumar ================= */

  function processMergeAggregate(files, keyLabel, valueLabel) {
    const filesWithCols = files.map(f => ({
      ...f,
      keyCol: f.headers.find(h => normHeader(h) === normHeader(keyLabel)),
      valueCol: f.headers.find(h => normHeader(h) === normHeader(valueLabel))
    }));

    const headerSets = filesWithCols.map(f => new Set(f.headers.map(normHeader)));
    const commonNorms = [...headerSets[0]].filter(hn =>
      headerSets.every(s => s.has(hn)) && hn !== normHeader(keyLabel) && hn !== normHeader(valueLabel)
    );
    const otherCommon = commonNorms.map(hn => {
      for (const f of filesWithCols) {
        const found = f.headers.find(h => normHeader(h) === hn);
        if (found !== undefined) return found;
      }
      return hn;
    });

    const sumLabel = `Suma de ${valueLabel}`;
    const map = new Map();
    const order = [];
    filesWithCols.forEach(f => {
      f.rows.forEach(row => {
        const rawKey = row[f.keyCol];
        if (rawKey === undefined || rawKey === null || String(rawKey).trim() === '') return;
        const norm = String(rawKey).trim().toLowerCase();
        const num = parseNumber(row[f.valueCol]);
        if (!map.has(norm)) {
          map.set(norm, { displayKey: String(rawKey).trim(), sum: 0, files: new Set(), extra: {} });
          order.push(norm);
        }
        const entry = map.get(norm);
        entry.sum += num;
        entry.files.add(f.id);
        otherCommon.forEach(displayHeader => {
          if (entry.extra[displayHeader] !== undefined && entry.extra[displayHeader] !== '') return;
          const matchHeader = f.headers.find(h => normHeader(h) === normHeader(displayHeader));
          if (matchHeader !== undefined) {
            const val = row[matchHeader];
            if (val !== undefined && val !== '') entry.extra[displayHeader] = val;
          }
        });
      });
    });

    let out = order.map(n => map.get(n));
    out = sortByKeyAsc(out, x => x.displayKey);
    const matched = out.filter(x => x.files.size > 1).length;
    const uniqueOnly = out.length - matched;

    const columns = [keyLabel, ...otherCommon, sumLabel];
    const data = out.map(x => {
      const o = {};
      o[keyLabel] = x.displayKey;
      otherCommon.forEach(h => { o[h] = x.extra[h] !== undefined ? x.extra[h] : ''; });
      o[sumLabel] = roundSmart(x.sum);
      return o;
    });

    const commonCount = otherCommon.length + 1;
    const stats = `<strong>${out.length}</strong> claves en el resultado &nbsp;·&nbsp; <strong>${matched}</strong> combinadas por coincidencia &nbsp;·&nbsp; <strong>${uniqueOnly}</strong> añadidas sin coincidencia &nbsp;·&nbsp; <strong>${commonCount}</strong> columnas comunes conservadas`;
    return { columns, data, stats };
  }

  /* ================= Procesamiento: Comparar en paralelo ================= */

  function processSideBySide(files, keyLabel) {
    const filesWithKey = files.map(f => ({
      ...f,
      keyCol: f.headers.find(h => normHeader(h) === normHeader(keyLabel))
    }));

    const groups = filesWithKey.map(f => {
      const g = new Map();
      f.rows.forEach(row => {
        const rawKey = row[f.keyCol];
        if (rawKey === undefined || rawKey === null || String(rawKey).trim() === '') return;
        const norm = String(rawKey).trim().toLowerCase();
        if (!g.has(norm)) g.set(norm, []);
        g.get(norm).push(row);
      });
      return g;
    });

    const allKeys = new Set();
    groups.forEach(g => { for (const k of g.keys()) allKeys.add(k); });

    const displayKeyFor = {};
    filesWithKey.forEach(f => {
      f.rows.forEach(row => {
        const rawKey = row[f.keyCol];
        if (rawKey === undefined || rawKey === null || String(rawKey).trim() === '') return;
        const norm = String(rawKey).trim().toLowerCase();
        if (!(norm in displayKeyFor)) displayKeyFor[norm] = String(rawKey).trim();
      });
    });

    // Columnas en orden: Archivo 1 completo, luego Archivo 2, y sucesivos.
    const columns = [];
    filesWithKey.forEach(f => f.headers.forEach(h => columns.push(`${f.label}: ${h}`)));

    let outputRows = [];
    allKeys.forEach(norm => {
      const counts = groups.map(g => (g.get(norm) || []).length);
      const maxCount = Math.max(...counts);
      for (let idx = 0; idx < maxCount; idx++) {
        const rowObj = { __key: displayKeyFor[norm], __hits: 0 };
        filesWithKey.forEach((f, i) => {
          const list = groups[i].get(norm) || [];
          const rowData = list[idx];
          if (rowData) rowObj.__hits++;
          f.headers.forEach(h => {
            rowObj[`${f.label}: ${h}`] = rowData ? (rowData[h] !== undefined ? rowData[h] : '') : '';
          });
        });
        outputRows.push(rowObj);
      }
    });

    outputRows = sortByKeyAsc(outputRows, r => r.__key);
    const total = outputRows.length;
    const full = outputRows.filter(r => r.__hits === filesWithKey.length).length;
    const partial = total - full;
    const stats = `<strong>${total}</strong> filas en el resultado &nbsp;·&nbsp; <strong>${full}</strong> con coincidencia en todos los archivos &nbsp;·&nbsp; <strong>${partial}</strong> con datos parciales`;
    return { columns, data: outputRows, stats };
  }

  /* ================= Botón Procesar ================= */

  els.processBtn.addEventListener('click', () => {
    const files = state.files.map(f => ({ id: f.id, label: f.label, headers: f.headers, rows: f.rows }));
    const result = state.mode === 'merge'
      ? processMergeAggregate(files, state.keyField, state.valueField)
      : processSideBySide(files, state.keyField);

    lastResult = { columns: result.columns, data: result.data };
    setupSortControls(result.columns, result.columns[0]);
    renderResults(result.stats);
  });

  /* ================= Orden multicriterio ================= */

  function setupSortControls(columns, defaultField) {
    const opts = columns.map(c => `<option value="${escapeAttr(c)}">${escapeHtml(c)}</option>`).join('');
    els.sortField1.innerHTML = opts;
    els.sortField1.value = defaultField;
    els.sortField2.innerHTML = `<option value="">— Ninguno —</option>${opts}`;
    els.sortField2.value = '';
    els.sortDir1.value = 'asc';
    els.sortDir2.value = 'asc';
  }

  function applySort() {
    if (!lastResult) return;
    const f1 = els.sortField1.value, d1 = els.sortDir1.value;
    const f2 = els.sortField2.value, d2 = els.sortDir2.value;
    lastResult.data.sort((ra, rb) => {
      let c = compareValues(ra[f1], rb[f1]);
      if (d1 === 'desc') c = -c;
      if (c !== 0) return c;
      if (f2) {
        let c2 = compareValues(ra[f2], rb[f2]);
        if (d2 === 'desc') c2 = -c2;
        return c2;
      }
      return 0;
    });
    renderTableOnly();
  }

  [els.sortField1, els.sortDir1, els.sortField2, els.sortDir2].forEach(el =>
    el.addEventListener('change', applySort)
  );

  /* ================= Render de resultados ================= */

  function renderResults(statsHtml) {
    els.stats.innerHTML = statsHtml;
    applySort(); // aplica el orden por defecto (Campo Clave, ascendente) y pinta la tabla
    els.stepResults.hidden = false;
    els.stepResults.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function renderTableOnly() {
    const thead = els.resultTable.querySelector('thead');
    const tbody = els.resultTable.querySelector('tbody');
    thead.innerHTML = ''; tbody.innerHTML = '';

    const trHead = document.createElement('tr');
    lastResult.columns.forEach(c => {
      const th = document.createElement('th');
      th.textContent = c;
      trHead.appendChild(th);
    });
    thead.appendChild(trHead);

    lastResult.data.forEach(row => {
      const tr = document.createElement('tr');
      lastResult.columns.forEach(c => {
        const td = document.createElement('td');
        const v = row[c];
        td.textContent = (v === undefined || v === null) ? '' : v;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
  }

  /* ================= Exportación ================= */

  els.exportXlsx.addEventListener('click', async () => {
    if (!lastResult) return;
    await ensureXlsx();
    const ws = XLSX.utils.json_to_sheet(lastResult.data, { header: lastResult.columns });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Resultado');
    XLSX.writeFile(wb, 'excel-lists-process-resultado.xlsx');
  });

  els.exportCsv.addEventListener('click', async () => {
    if (!lastResult) return;
    await ensureXlsx();
    const ws = XLSX.utils.json_to_sheet(lastResult.data, { header: lastResult.columns });
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'excel-lists-process-resultado.csv';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  els.copyBtn.addEventListener('click', async () => {
    if (!lastResult) return;
    const lines = [lastResult.columns.join('\t')];
    lastResult.data.forEach(row => {
      lines.push(lastResult.columns.map(c => (row[c] === undefined || row[c] === null) ? '' : row[c]).join('\t'));
    });
    try {
      await navigator.clipboard.writeText(lines.join('\n'));
      const original = els.copyBtn.textContent;
      els.copyBtn.textContent = 'Copiado ✓';
      setTimeout(() => { els.copyBtn.textContent = original; }, 1600);
    } catch (err) {
      alert('No se pudo copiar automáticamente. Selecciona la tabla manualmente y copia con Ctrl+C.');
    }
  });

  /* ================= Init ================= */

  renderFileSlots();
  updateProcessButtonState();
})();
