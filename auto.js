// ============================================================
// ImageForge Auto — Pipeline Script
// ============================================================

// ---- STATE ----
const state = {
  sourceType: 'excel',       // 'excel' | 'archive'
  imageItems: [],            // [{url, filename, row, col, status}]
  selectedRows: new Set(),
  overlayImg: null,          // HTMLImageElement
  overlayMeta: {},           // {name, w, h}
  previewIdx: 0,
  isRunning: false,
  outputUrls: [],            // [{filename, url}] after upload
};

// ---- DOM REFS ----
const $ = id => document.getElementById(id);

// ============================================================
// SOURCE TAB SWITCHING
// ============================================================
function switchSourceTab(tab) {
  state.sourceType = tab;
  document.querySelectorAll('.src-tab').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.toggle('active', c.id === 'tab-' + tab));
  checkReady();
}

// ============================================================
// EXCEL / SHEETS FETCH
// ============================================================
async function fetchExcel() {
  const url = $('excel-url').value.trim();
  if (!url) return;

  const btn = $('btn-fetch-excel');
  btn.textContent = 'Fetching...';
  btn.disabled = true;

  try {
    let fetchUrl = url;

    // Google Sheets: convert to CSV export URL
    if (url.includes('docs.google.com/spreadsheets')) {
      // Support both /edit and /pub URLs
      const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (match) {
        const sheetId = match[1];
        fetchUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=0`;
      }
    }

    // Try direct fetch first, then CORS proxy
    let data;
    try {
      const resp = await fetch(fetchUrl);
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      const buf = await resp.arrayBuffer();
      data = buf;
    } catch (e) {
      clog(`[WARN] Direct fetch blocked, trying CORS proxy...`, 'warn');
      const proxy = `https://corsproxy.io/?${encodeURIComponent(fetchUrl)}`;
      const resp = await fetch(proxy);
      if (!resp.ok) throw new Error('Proxy fetch failed: HTTP ' + resp.status);
      data = await resp.arrayBuffer();
    }

    parseExcelBuffer(data, url);
  } catch (err) {
    clog(`[ERR] Fetch gagal: ${err.message}`, 'err');
    alert('Gagal fetch: ' + err.message);
  }

  btn.textContent = 'Fetch';
  btn.disabled = false;
}

// Local Excel file
$('excel-file').addEventListener('change', e => {
  if (e.target.files[0]) handleLocalExcel(e.target.files[0]);
});
setupDrop('excel-mini-drop', 'excel-file', f => handleLocalExcel(f));

function handleLocalExcel(file) {
  const reader = new FileReader();
  reader.onload = e => parseExcelBuffer(e.target.result, file.name, file);
  reader.readAsArrayBuffer(file);
}

function parseExcelBuffer(buffer, srcName, file = null) {
  try {
    const wb = XLSX.read(buffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!rows.length) { alert('File kosong.'); return; }

    const headers = rows[0].map((h, i) => h || `Col${i+1}`);

    // Auto-detect URL columns
    const urlColCounts = headers.map((_, ci) => {
      let c = 0;
      for (let r = 1; r < Math.min(rows.length, 40); r++) {
        const v = String(rows[r][ci] || '').trim();
        if (v.startsWith('http') || v.startsWith('//')) c++;
      }
      return c;
    });

    // Render column chips
    const chipsEl = $('col-chips');
    chipsEl.innerHTML = '';
    const autoSelected = new Set();

    headers.forEach((h, i) => {
      const cnt = urlColCounts[i];
      if (cnt > 0) autoSelected.add(i);
      const chip = document.createElement('button');
      chip.className = 'chip' + (cnt > 0 ? ' active' : '');
      chip.dataset.colIdx = i;
      chip.innerHTML = `${h}${cnt > 0 ? ` <small style="opacity:.6">(${cnt})</small>` : ''}`;
      chip.onclick = () => {
        if (autoSelected.has(i)) { autoSelected.delete(i); chip.classList.remove('active'); }
        else { autoSelected.add(i); chip.classList.add('active'); }
        buildImageListFromExcel(rows, headers, autoSelected);
      };
      chipsEl.appendChild(chip);
    });

    $('col-picker').style.display = 'block';

    // Update drop label
    const label = file ? `${file.name} (${rows.length - 1} baris)` : `Loaded — ${rows.length - 1} baris`;
    const dropEl = $('excel-mini-drop');
    dropEl.querySelector('span').textContent = label;
    dropEl.classList.add('loaded');

    $('excel-stat').style.display = 'flex';
    $('excel-stat').innerHTML = `
      <div class="stat-pill"><div class="stat-val">${rows.length - 1}</div><div class="stat-lbl">Baris</div></div>
      <div class="stat-pill"><div class="stat-val">${headers.length}</div><div class="stat-lbl">Kolom</div></div>
    `;

    buildImageListFromExcel(rows, headers, autoSelected);
    setPipeActive(1);
  } catch (err) {
    alert('Gagal parse Excel: ' + err.message);
  }
}

function buildImageListFromExcel(rows, headers, selectedCols) {
  state.imageItems = [];
  for (let r = 1; r < rows.length; r++) {
    for (const ci of selectedCols) {
      const v = String(rows[r][ci] || '').trim();
      if (v && (v.startsWith('http') || v.startsWith('//'))) {
        state.imageItems.push({ url: v, filename: null, row: r, col: headers[ci], status: 'pending' });
      }
    }
  }
  state.selectedRows = new Set(state.imageItems.map((_, i) => i));
  renderQueue();
  checkReady();
}

// ============================================================
// ARCHIVE FETCH
// ============================================================
async function fetchArchive() {
  const url = $('archive-url').value.trim();
  if (!url) return;

  const btn = $('btn-fetch-archive');
  btn.textContent = 'Fetching...';
  btn.disabled = true;

  clog(`[INFO] Fetching archive dari: ${url}`, 'info');
  showConsole();

  try {
    let resp;
    try {
      resp = await fetch(url);
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
    } catch (e) {
      clog(`[WARN] Proxy fallback...`, 'warn');
      resp = await fetch(`https://corsproxy.io/?${encodeURIComponent(url)}`);
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
    }

    const buf = await resp.arrayBuffer();
    const filename = url.split('/').pop() || 'archive.zip';
    const blob = new Blob([buf]);
    const file = new File([blob], filename);
    await extractArchive(file, buf);
  } catch (err) {
    clog(`[ERR] Fetch archive gagal: ${err.message}`, 'err');
    alert('Gagal: ' + err.message);
  }

  btn.textContent = 'Fetch';
  btn.disabled = false;
}

// Local archive
$('archive-file').addEventListener('change', e => {
  if (e.target.files[0]) handleLocalArchive(e.target.files[0]);
});
setupDrop('archive-mini-drop', 'archive-file', f => handleLocalArchive(f));

async function handleLocalArchive(file) {
  const buf = await file.arrayBuffer();
  await extractArchive(file, buf);
}

async function extractArchive(file, arrayBuffer) {
  showConsole();
  clog(`[INFO] Mengekstrak ${file.name}...`, 'info');

  try {
    const zip = new JSZip();
    const content = await zip.loadAsync(arrayBuffer);
    const imgPaths = Object.keys(content.files).filter(p => {
      const ext = p.split('.').pop().toLowerCase();
      return ['png','jpg','jpeg','webp'].includes(ext) && !content.files[p].dir;
    });

    if (!imgPaths.length) throw new Error('Tidak ada gambar valid di archive.');

    // Revoke old object URLs
    state.imageItems.forEach(it => { if (it.url?.startsWith('blob:')) URL.revokeObjectURL(it.url); });

    state.imageItems = [];
    for (let i = 0; i < imgPaths.length; i++) {
      const path = imgPaths[i];
      const blob = await content.files[path].async('blob');
      const url = URL.createObjectURL(blob);
      const fname = path.split('/').pop();
      state.imageItems.push({ url, filename: fname, row: i + 1, col: 'Archive', status: 'pending', _blob: blob });
    }

    state.selectedRows = new Set(state.imageItems.map((_, i) => i));

    const dropEl = $('archive-mini-drop');
    dropEl.querySelector('span').textContent = `${file.name} — ${imgPaths.length} gambar`;
    dropEl.classList.add('loaded');

    $('archive-stat').style.display = 'flex';
    $('archive-stat').innerHTML = `
      <div class="stat-pill"><div class="stat-val">${imgPaths.length}</div><div class="stat-lbl">Gambar</div></div>
      <div class="stat-pill"><div class="stat-val">${(file.size/1024/1024).toFixed(1)} MB</div><div class="stat-lbl">Ukuran</div></div>
    `;

    clog(`[OK] Diekstrak ${imgPaths.length} gambar dari ${file.name}`, 'ok');
    renderQueue();
    checkReady();
    setPipeActive(1);
  } catch (err) {
    clog(`[ERR] Gagal ekstrak: ${err.message}`, 'err');
    alert('Gagal: ' + err.message);
  }
}

// ============================================================
// OVERLAY
// ============================================================
async function fetchOverlay() {
  const url = $('overlay-url').value.trim();
  if (!url) return;

  const btn = $('btn-fetch-overlay');
  btn.textContent = 'Loading...';
  btn.disabled = true;

  try {
    const img = await loadImage(url);
    setOverlayImage(img, url.split('/').pop() || 'overlay', url);
  } catch (err) {
    clog(`[ERR] Gagal load overlay: ${err.message}`, 'err');
    alert('Gagal load overlay: ' + err.message);
  }

  btn.textContent = 'Load';
  btn.disabled = false;
}

$('overlay-file').addEventListener('change', e => {
  if (e.target.files[0]) handleOverlayFile(e.target.files[0]);
});
setupDrop('overlay-mini-drop', 'overlay-file', f => handleOverlayFile(f));

function handleOverlayFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    const img = new Image();
    img.onload = () => setOverlayImage(img, file.name, e.target.result);
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

function setOverlayImage(img, name, src) {
  state.overlayImg = img;
  state.overlayMeta = { name, w: img.naturalWidth, h: img.naturalHeight };

  const wrap = $('overlay-preview-wrap');
  $('overlay-preview-img').src = src;
  $('overlay-meta').innerHTML = `
    <strong>${name}</strong><br>
    ${img.naturalWidth} × ${img.naturalHeight}px
  `;
  wrap.style.display = 'flex';

  const dropEl = $('overlay-mini-drop');
  dropEl.querySelector('span').textContent = `${name} (${img.naturalWidth}×${img.naturalHeight})`;
  dropEl.classList.add('loaded');

  checkReady();
  setPipeActive(2);
  generatePreview();
}

// ============================================================
// QUEUE RENDER
// ============================================================
function renderQueue() {
  const list = $('queue-list');
  const countLbl = $('queue-count-lbl');
  const card = $('card-queue');

  if (!state.imageItems.length) { card.style.display = 'none'; return; }
  card.style.display = 'block';

  countLbl.textContent = `${state.imageItems.length} gambar`;

  list.innerHTML = '';
  state.imageItems.slice(0, 200).forEach((item, i) => {
    const display = item.filename || item.url;
    const div = document.createElement('div');
    div.className = 'queue-item';
    div.id = `qi-${i}`;
    div.innerHTML = `
      <input type="checkbox" ${state.selectedRows.has(i) ? 'checked' : ''} onchange="toggleQueueRow(${i}, this.checked)" style="accent-color:var(--accent)">
      <span class="q-url" title="${item.url}">${display}</span>
      <span class="q-status pending" id="qs-${i}"></span>
    `;
    list.appendChild(div);
  });

  if (state.imageItems.length > 200) {
    list.innerHTML += `<div style="text-align:center; padding:8px; font-size:11px; color:var(--text-muted);">... dan ${state.imageItems.length - 200} lainnya</div>`;
  }

  // Render preview chips
  generatePreview(0);
}

function toggleQueueRow(i, checked) {
  if (checked) state.selectedRows.add(i);
  else state.selectedRows.delete(i);
}

function toggleAllRows(checked) {
  state.imageItems.forEach((_, i) => {
    if (checked) state.selectedRows.add(i);
    else state.selectedRows.delete(i);
    const cb = document.querySelector(`#qi-${i} input[type=checkbox]`);
    if (cb) cb.checked = checked;
  });
}

function updateQueueStatus(i, status) {
  const dot = $(`qs-${i}`);
  if (dot) dot.className = 'q-status ' + status;
}

// ============================================================
// PREVIEW
// ============================================================
let previewIdx = 0;

async function generatePreview(idx = null) {
  if (idx !== null) previewIdx = idx;
  if (!state.overlayImg || !state.imageItems.length) return;

  const item = state.imageItems[previewIdx];
  if (!item) return;

  const canvas = $('preview-canvas');
  const empty = $('preview-empty');
  const label = $('preview-label');

  label.textContent = `${previewIdx + 1} / ${state.imageItems.length}`;

  try {
    const base = await loadImage(item.url);
    const result = composite(base, state.overlayImg);
    canvas.width = result.width;
    canvas.height = result.height;
    canvas.getContext('2d').drawImage(result, 0, 0);
    canvas.style.display = 'block';
    empty.style.display = 'none';
    setPipeActive(3);
  } catch (e) {
    empty.querySelector('p').textContent = 'Gagal load preview: ' + e.message;
    canvas.style.display = 'none';
    empty.style.display = 'block';
  }
}

function previewNav(dir) {
  const newIdx = previewIdx + dir;
  if (newIdx < 0 || newIdx >= state.imageItems.length) return;
  previewIdx = newIdx;
  generatePreview(previewIdx);
}

// ============================================================
// COMPOSITE ENGINE
// ============================================================
function composite(baseImg, overlayImg) {
  const gap    = parseInt($('s-gap').value) || 0;
  const zoom   = parseFloat($('s-zoom').value) || 100;
  const mode   = $('s-mode').value;
  const ratio  = parseFloat($('s-ratio').value) || 1;

  const origBw = baseImg.naturalWidth || baseImg.width;
  const origBh = baseImg.naturalHeight || baseImg.height;
  const bw = Math.round(origBw * (zoom / 100));
  const bh = Math.round(origBh * (zoom / 100));
  const oNW = overlayImg.naturalWidth;
  const oNH = overlayImg.naturalHeight;
  const overlayAR    = oNW / oNH;
  const overlayARinv = oNH / oNW;

  let canvasW, canvasH, baseX, baseY, ovX, ovY, ovW, ovH;

  if (mode === 'append-right') {
    let baseH = Math.max(bh, Math.ceil((bw + gap) / ratio));
    let reqH = baseH;
    if (ratio > overlayAR) reqH = Math.max(baseH, Math.ceil((bw + gap) / (ratio - overlayAR)));
    canvasH = reqH;
    canvasW = Math.round(canvasH * ratio);
    ovH = canvasH;
    ovW = Math.round(overlayAR * canvasH);
    const maxOW = canvasW - bw - gap;
    if (ovW > maxOW && maxOW > 0) { ovW = maxOW; ovH = canvasH; }
    baseX = 0; baseY = Math.round((canvasH - bh) / 2);
    ovX = canvasW - ovW; ovY = 0;

  } else if (mode === 'append-left') {
    let baseH = Math.max(bh, Math.ceil((bw + gap) / ratio));
    let reqH = baseH;
    if (ratio > overlayAR) reqH = Math.max(baseH, Math.ceil((bw + gap) / (ratio - overlayAR)));
    canvasH = reqH;
    canvasW = Math.round(canvasH * ratio);
    ovH = canvasH;
    ovW = Math.round(overlayAR * canvasH);
    const maxOW2 = canvasW - bw - gap;
    if (ovW > maxOW2 && maxOW2 > 0) { ovW = maxOW2; ovH = canvasH; }
    ovX = 0; ovY = 0;
    baseX = canvasW - bw; baseY = Math.round((canvasH - bh) / 2);

  } else if (mode === 'append-top') {
    const tarARinv = 1 / ratio;
    let baseW = Math.max(bw, Math.ceil((bh + gap) / tarARinv));
    let reqW = baseW;
    if (tarARinv > overlayARinv) reqW = Math.max(baseW, Math.ceil((bh + gap) / (tarARinv - overlayARinv)));
    canvasW = reqW; canvasH = Math.round(canvasW * tarARinv);
    ovW = canvasW; ovH = Math.round(overlayARinv * canvasW);
    const maxOH = canvasH - bh - gap;
    if (ovH > maxOH && maxOH > 0) { ovH = maxOH; ovW = canvasW; }
    ovX = 0; ovY = 0;
    baseX = Math.round((canvasW - bw) / 2); baseY = canvasH - bh;

  } else if (mode === 'append-bottom') {
    const tarARinv2 = 1 / ratio;
    let baseW2 = Math.max(bw, Math.ceil((bh + gap) / tarARinv2));
    let reqW2 = baseW2;
    if (tarARinv2 > overlayARinv) reqW2 = Math.max(baseW2, Math.ceil((bh + gap) / (tarARinv2 - overlayARinv)));
    canvasW = reqW2; canvasH = Math.round(canvasW * tarARinv2);
    ovW = canvasW; ovH = Math.round(overlayARinv * canvasW);
    const maxOH2 = canvasH - bh - gap;
    if (ovH > maxOH2 && maxOH2 > 0) { ovH = maxOH2; ovW = canvasW; }
    baseX = Math.round((canvasW - bw) / 2); baseY = 0;
    ovX = 0; ovY = canvasH - ovH;

  } else {
    // center
    if (bw / bh > ratio) { canvasW = bw; canvasH = Math.round(bw / ratio); }
    else { canvasH = bh; canvasW = Math.round(bh * ratio); }
    baseX = Math.round((canvasW - bw) / 2);
    baseY = Math.round((canvasH - bh) / 2);
    const scale = Math.min((bw * 0.8) / oNW, (bh * 0.8) / oNH, 1);
    ovW = Math.round(oNW * scale);
    ovH = Math.round(oNH * scale);
    ovX = Math.round((canvasW - ovW) / 2);
    ovY = Math.round((canvasH - ovH) / 2);
  }

  const c = document.createElement('canvas');
  c.width = canvasW;
  c.height = canvasH;
  const ctx = c.getContext('2d');
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, canvasW, canvasH);
  ctx.drawImage(baseImg, baseX, baseY, bw, bh);
  ctx.drawImage(overlayImg, ovX, ovY, ovW, ovH);
  return c;
}

// ============================================================
// LOAD IMAGE HELPER
// ============================================================
function loadImage(url) {
  return new Promise((res, rej) => {
    const img = new Image();
    const isBlobOrData = url.startsWith('blob:') || url.startsWith('data:');
    if (!isBlobOrData) img.crossOrigin = 'anonymous';

    img.onload = () => res(img);
    img.onerror = () => {
      if (isBlobOrData) { rej(new Error('Cannot load: ' + url)); return; }
      // fallback CORS proxy
      const img2 = new Image();
      img2.crossOrigin = 'anonymous';
      img2.onload = () => res(img2);
      img2.onerror = () => rej(new Error('Cannot load (proxy also failed): ' + url));
      img2.src = `https://corsproxy.io/?${encodeURIComponent(url)}`;
    };
    img.src = url;
  });
}

// ============================================================
// PIPELINE RUNNER
// ============================================================
async function startPipeline() {
  if (!state.imageItems.length || !state.overlayImg || state.isRunning) return;

  state.isRunning = true;
  state.outputUrls = [];

  const btn = $('btn-run');
  btn.textContent = '⏳ Running...';
  btn.classList.add('running');
  btn.disabled = true;

  const runDot = $('run-status-dot');
  runDot.className = 'run-status-dot running';
  $('run-status-text').textContent = 'Pipeline berjalan...';

  showConsole();
  $('card-output').style.display = 'none';

  const fmt      = $('s-fmt').value;
  const quality  = (parseInt($('s-quality').value) || 92) / 100;
  const mimeType = fmt === 'jpeg' ? 'image/jpeg' : fmt === 'webp' ? 'image/webp' : 'image/png';

  const uploadMode    = document.querySelector('input[name="upload-mode"]:checked').value;
  const doUpload      = uploadMode === 'upload';
  const uploadEndpoint = doUpload ? $('upload-endpoint').value.trim() : '';
  const uploadMethod  = doUpload ? $('upload-method').value : 'POST';
  const uploadField   = doUpload ? ($('upload-field').value.trim() || 'file') : 'file';
  const uploadAuth    = doUpload ? $('upload-auth').value.trim() : '';
  const uploadBaseUrl = doUpload ? $('upload-base-url').value.trim() : '';

  if (doUpload && !uploadEndpoint) {
    alert('Upload endpoint tidak boleh kosong untuk mode "ZIP + Upload ke CDN".');
    resetRunState();
    return;
  }

  const zip = new JSZip();
  const usedNames = new Set();
  const selectedList = state.imageItems.filter((_, i) => state.selectedRows.has(i));

  clog(`[INFO] Mulai pipeline: ${selectedList.length} gambar, mode=${fmt}, upload=${doUpload}`, 'info');
  setPipeActive(3);

  let done = 0, errors = 0;
  const CONCURRENCY = 3;
  let idx = 0;

  async function processOne() {
    while (idx < selectedList.length) {
      const i = idx++;
      const item = selectedList[i];
      const globalIdx = state.imageItems.indexOf(item);

      item.status = 'loading';
      updateQueueStatus(globalIdx, 'loading');

      try {
        const base = await loadImage(item.url);
        const canvas = composite(base, state.overlayImg);
        const dataUrl = canvas.toDataURL(mimeType, quality);
        const base64 = dataUrl.split(',')[1];

        // Determine output filename
        const origName = item.filename || (new URL(item.url, 'https://x.com').pathname.split('/').pop()) || `image_${i+1}`;
        let baseName = origName.replace(/\.[^/.]+$/, '').replace(/[^a-zA-Z0-9_\-]/g, '_').replace(/_+/g, '_') || `img_${i+1}`;
        let outName = `${baseName}.${fmt}`;
        let cnt = 2;
        while (usedNames.has(outName)) { outName = `${baseName}_${cnt++}.${fmt}`; }
        usedNames.add(outName);

        // Add to ZIP always
        zip.file(outName, base64, { base64: true });

        // Upload if configured
        if (doUpload) {
          try {
            const blob = dataUrlToBlob(dataUrl);
            const uploadedUrl = await uploadToServer(blob, outName, uploadEndpoint, uploadMethod, uploadField, uploadAuth, uploadBaseUrl);
            state.outputUrls.push({ filename: outName, url: uploadedUrl });
            clog(`[OK] ${outName} → ${uploadedUrl}`, 'ok');
          } catch (upErr) {
            state.outputUrls.push({ filename: outName, url: null, error: upErr.message });
            clog(`[WARN] Upload gagal untuk ${outName}: ${upErr.message}`, 'warn');
          }
        } else {
          state.outputUrls.push({ filename: outName, url: null });
          clog(`[OK] ${outName} — ${canvas.width}×${canvas.height}px`, 'ok');
        }

        item.status = 'done';
        updateQueueStatus(globalIdx, 'done');
      } catch (err) {
        item.status = 'error';
        updateQueueStatus(globalIdx, 'error');
        errors++;
        clog(`[ERR] Row ${item.row}: ${err.message}`, 'err');
      }

      done++;
      updateProgress(done, selectedList.length);
    }
  }

  const workers = Array.from({ length: CONCURRENCY }, () => processOne());
  await Promise.all(workers);

  clog(`[INFO] Composite selesai: ${done - errors} OK, ${errors} gagal.`, 'info');

  // Generate & download ZIP
  $('prog-text').textContent = 'Membuat ZIP...';
  const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  const zipUrl = URL.createObjectURL(zipBlob);
  const a = document.createElement('a');
  a.href = zipUrl;
  a.download = `imageforge_auto_${Date.now()}.zip`;
  a.click();
  URL.revokeObjectURL(zipUrl);

  clog(`[OK] ZIP didownload (${(zipBlob.size / 1024 / 1024).toFixed(2)} MB)`, 'ok');

  // Show output URLs
  renderOutputUrls();
  setPipeActive(4);

  runDot.className = 'run-status-dot done';
  $('run-status-text').textContent = `✓ Selesai — ${done - errors} gambar berhasil`;

  resetRunState(false);
}

function resetRunState(fullyReset = true) {
  state.isRunning = false;
  const btn = $('btn-run');
  btn.textContent = 'Run Pipeline';
  btn.classList.remove('running');
  if (fullyReset) {
    btn.disabled = true;
  } else {
    btn.disabled = false;
  }
}

// ============================================================
// UPLOAD TO SERVER
// ============================================================
async function uploadToServer(blob, filename, endpoint, method, fieldName, authHeader, baseUrl) {
  const headers = {};
  if (authHeader) headers['Authorization'] = authHeader;

  let response;
  if (method === 'PUT') {
    headers['Content-Type'] = blob.type || 'application/octet-stream';
    response = await fetch(`${endpoint}/${filename}`, {
      method: 'PUT',
      headers,
      body: blob,
    });
  } else {
    const fd = new FormData();
    fd.append(fieldName, blob, filename);
    response = await fetch(endpoint, {
      method: 'POST',
      headers,
      body: fd,
    });
  }

  if (!response.ok) throw new Error(`Server returned ${response.status}`);

  // Try to parse JSON response for URL
  const ct = response.headers.get('content-type') || '';
  if (ct.includes('application/json')) {
    const json = await response.json();
    // Common response field names
    const urlField = json.url || json.path || json.filename || json.file || json.location || json.src;
    if (urlField) {
      if (urlField.startsWith('http')) return urlField;
      return baseUrl.replace(/\/$/, '') + '/' + urlField.replace(/^\//, '');
    }
  }

  // Fallback: construct URL from base + filename
  return baseUrl.replace(/\/$/, '') + '/' + filename;
}

// ============================================================
// OUTPUT URLS
// ============================================================
function renderOutputUrls() {
  const card = $('card-output');
  const list = $('url-output-list');
  const count = $('output-count');

  const withUrls = state.outputUrls.filter(o => o.url);
  count.textContent = `${withUrls.length} URL dihasilkan`;
  card.style.display = 'block';

  list.innerHTML = '';

  if (!withUrls.length) {
    list.innerHTML = '<div style="font-size:12px; color:var(--text-muted); padding:12px; text-align:center;">Mode ZIP Only — tidak ada URL CDN. Upload ke img.offscript.id secara manual.</div>';
    return;
  }

  withUrls.forEach(item => {
    const div = document.createElement('div');
    div.className = 'url-out-item';
    div.innerHTML = `
      <a href="${item.url}" target="_blank" title="${item.url}">${item.url}</a>
      <button class="url-copy-btn" onclick="navigator.clipboard.writeText('${item.url}'); this.textContent='✓'" title="Copy">⎘</button>
    `;
    list.appendChild(div);
  });
}

function copyAllUrls() {
  const urls = state.outputUrls.filter(o => o.url).map(o => o.url).join('\n');
  if (!urls) { alert('Tidak ada URL untuk dicopy.'); return; }
  navigator.clipboard.writeText(urls);
  const btn = event.target;
  btn.textContent = 'Copied!';
  setTimeout(() => btn.textContent = 'Copy All URLs', 2000);
}

function downloadUrlList() {
  const urls = state.outputUrls.filter(o => o.url).map(o => o.url).join('\n');
  if (!urls) { alert('Tidak ada URL.'); return; }
  const blob = new Blob([urls], { type: 'text/plain' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `url_list_${Date.now()}.txt`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ============================================================
// HELPERS
// ============================================================
function dataUrlToBlob(dataUrl) {
  const [header, data] = dataUrl.split(',');
  const mime = header.match(/:(.*?);/)[1];
  const bin = atob(data);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return new Blob([arr], { type: mime });
}

function updateProgress(done, total) {
  const pct = Math.round((done / total) * 100);
  $('prog-fill').style.width = pct + '%';
  $('prog-pct').textContent = pct + '%';
  $('prog-text').textContent = `${done} / ${total}`;
}

function clog(msg, type = '') {
  const el = $('console-log');
  const line = document.createElement('div');
  if (type) line.className = 'log-' + type;
  const ts = new Date().toLocaleTimeString('id-ID', { hour12: false });
  line.textContent = `[${ts}] ${msg}`;
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}

function showConsole() {
  $('card-console').style.display = 'block';
}

// ============================================================
// READINESS CHECK
// ============================================================
function onInputChange() {
  checkReady();
}

function checkReady() {
  const hasSource = state.imageItems.length > 0;
  const hasOverlay = state.overlayImg !== null;
  const ready = hasSource && hasOverlay && !state.isRunning;

  const btn = $('btn-run');
  btn.disabled = !ready;

  const dot = $('run-status-dot');
  const txt = $('run-status-text');

  if (state.isRunning) {
    dot.className = 'run-status-dot running';
    txt.textContent = 'Pipeline berjalan...';
  } else if (ready) {
    dot.className = 'run-status-dot ready';
    txt.textContent = `Siap — ${state.imageItems.length} gambar, overlay dimuat`;
  } else if (!hasSource) {
    dot.className = 'run-status-dot';
    txt.textContent = 'Belum siap — isi source (Excel/Archive)';
  } else if (!hasOverlay) {
    dot.className = 'run-status-dot';
    txt.textContent = 'Belum siap — load overlay image';
  }
}

// ============================================================
// PIPELINE STEP INDICATOR
// ============================================================
function setPipeActive(step) {
  for (let i = 1; i <= 4; i++) {
    const el = $('pipe-' + i);
    if (!el) continue;
    if (i < step) { el.className = 'pipe-step done'; }
    else if (i === step) { el.className = 'pipe-step active'; }
    else { el.className = 'pipe-step'; }
  }
}

// ============================================================
// UPLOAD MODE UI
// ============================================================
function updateUploadModeUI() {
  const mode = document.querySelector('input[name="upload-mode"]:checked').value;
  $('upload-cfg-fields').style.display = mode === 'upload' ? 'block' : 'none';
}

// ============================================================
// SETTINGS UI
// ============================================================
function updateSettingsUI() {
  // currently no conditional UI needed for settings
}

// ============================================================
// DRAG-DROP HELPER
// ============================================================
function setupDrop(dropId, inputId, handler) {
  const drop = $(dropId);
  if (!drop) return;

  drop.addEventListener('dragenter', e => { e.preventDefault(); drop.classList.add('drag-over'); });
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag-over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
  drop.addEventListener('drop', e => {
    e.preventDefault();
    drop.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) handler(f);
  });

  const input = $(inputId);
  if (input) {
    input.addEventListener('change', e => {
      if (e.target.files[0]) handler(e.target.files[0]);
    });
  }
}

// ============================================================
// INIT
// ============================================================
updateUploadModeUI();
checkReady();
