let excelData = [];
let imageUrls = [];
let overlayImage = null;
let selectedRows = new Set();

const excelDrop = document.getElementById('excel-drop');
const overlayDrop = document.getElementById('overlay-drop');
const excelFileInput = document.getElementById('excel-file');
const overlayFileInput = document.getElementById('overlay-file');

['dragenter','dragover'].forEach(e => {
  excelDrop.addEventListener(e, ev => { ev.preventDefault(); excelDrop.classList.add('drag-over'); });
  overlayDrop.addEventListener(e, ev => { ev.preventDefault(); overlayDrop.classList.add('drag-over'); });
});
['dragleave','drop'].forEach(e => {
  excelDrop.addEventListener(e, () => excelDrop.classList.remove('drag-over'));
  overlayDrop.addEventListener(e, () => overlayDrop.classList.remove('drag-over'));
});

excelDrop.addEventListener('drop', ev => {
  ev.preventDefault();
  const f = ev.dataTransfer.files[0];
  if (f) handleExcelFile(f);
});
overlayDrop.addEventListener('drop', ev => {
  ev.preventDefault();
  const f = ev.dataTransfer.files[0];
  if (f) handleOverlayFile(f);
});

excelFileInput.addEventListener('change', e => { if (e.target.files[0]) handleExcelFile(e.target.files[0]); });
overlayFileInput.addEventListener('change', e => { if (e.target.files[0]) handleOverlayFile(e.target.files[0]); });

let selectedColumns = new Set();

function handleExcelFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (!rows.length) { alert('File kosong atau tidak valid.'); return; }

      excelData = rows;
      const headers = rows[0].map((h, i) => h || `Kolom ${i + 1}`);
      selectedColumns.clear();

      // Scan setiap kolom: hitung berapa baris yang berisi URL
      const colUrlCounts = headers.map((_, colIdx) => {
        let count = 0;
        for (let r = 1; r < Math.min(rows.length, 30); r++) {
          const val = String(rows[r][colIdx] || '').trim();
          if (val.startsWith('http') || val.startsWith('//')) count++;
        }
        return count;
      });

      // Render chips multi-select
      const chipsContainer = document.getElementById('col-chips');
      chipsContainer.innerHTML = '';
      headers.forEach((h, i) => {
        const urlCount = colUrlCounts[i];
        const isUrl = urlCount > 0;
        const chip = document.createElement('button');
        chip.className = 'chip' + (isUrl ? ' active' : '');
        chip.dataset.colIdx = i;
        chip.innerHTML = `${h}` + (urlCount > 0 ? ` <span style="font-size:10px;opacity:0.7;">(${urlCount} URL)</span>` : '');
        if (isUrl) selectedColumns.add(i);

        chip.addEventListener('click', () => {
          if (selectedColumns.has(i)) {
            selectedColumns.delete(i);
            chip.classList.remove('active');
          } else {
            selectedColumns.add(i);
            chip.classList.add('active');
          }
          onColumnSelected();
        });
        chipsContainer.appendChild(chip);
      });

      document.getElementById('col-select-wrap').style.display = 'flex';
      document.getElementById('excel-info').style.display = 'block';

      const stats = document.getElementById('excel-stats');
      stats.innerHTML = `
        <div class="stat-mini-item"><div class="stat-mini-val">${rows.length - 1}</div><div class="stat-mini-lbl">Baris data</div></div>
        <div class="stat-mini-item"><div class="stat-mini-val">${headers.length}</div><div class="stat-mini-lbl">Kolom</div></div>
      `;

      const excelIcon = excelDrop.querySelector('.drop-icon svg');
      excelDrop.querySelector('.drop-title').textContent = file.name;
      excelDrop.querySelector('.drop-sub').textContent = `${(file.size/1024).toFixed(1)} KB — ${wb.SheetNames[0]}`;
      excelDrop.querySelector('.drop-icon').style.background = 'rgba(200,245,80,0.1)';
      excelIcon.style.stroke = 'var(--accent)';

      if (selectedColumns.size > 0) onColumnSelected();

      setStep(2);
    } catch (err) {
      alert('Gagal membaca file: ' + err.message);
    }
  };
  reader.readAsBinaryString(file);
}

function onColumnSelected() {
  if (selectedColumns.size === 0) {
    imageUrls = [];
    renderTable();
    document.getElementById('badge-count').textContent = '0 gambar';
    document.getElementById('tbl-count-badge').textContent = '0';
    return;
  }

  const headers = excelData[0].map((h, i) => h || `Kolom ${i + 1}`);
  imageUrls = [];
  for (let i = 1; i < excelData.length; i++) {
    for (const colIdx of selectedColumns) {
      const val = String(excelData[i][colIdx] || '').trim();
      if (val && (val.startsWith('http') || val.startsWith('//'))) {
        imageUrls.push({ url: val, row: i, col: headers[colIdx], status: 'pending' });
      }
    }
  }

  selectedRows = new Set(imageUrls.map((_, i) => i));
  renderTable();
  document.getElementById('badge-count').textContent = imageUrls.length + ' gambar';
  document.getElementById('tbl-count-badge').textContent = imageUrls.length;
  document.getElementById('table-section').style.display = 'block';
  document.getElementById('preview-section').style.display = 'block';
  checkReady();
}

function renderTable() {
  const tbody = document.getElementById('images-tbody');
  if (!imageUrls.length) {
    tbody.innerHTML = `<tr><td colspan="5"><div class="empty-state">Tidak ada URL valid ditemukan di kolom yang dipilih.</div></td></tr>`;
    return;
  }
  tbody.innerHTML = imageUrls.slice(0, 200).map((item, i) => `
    <tr>
      <td><input type="checkbox" ${selectedRows.has(i) ? 'checked' : ''} onchange="toggleRow(${i}, this.checked)" style="accent-color:var(--accent); cursor:pointer;"></td>
      <td style="color:var(--text-muted); font-size:12px;">${item.row}</td>
      <td><span class="badge badge-purple" style="font-size:10px;">${item.col || '-'}</span></td>
      <td><a href="${item.url}" target="_blank" class="url-cell" title="${item.url}">${item.url}</a></td>
      <td><span class="status-dot ${item.status}" id="sdot-${i}"></span><span id="stxt-${i}" style="font-size:12px;">${statusLabel(item.status)}</span></td>
    </tr>
  `).join('') + (imageUrls.length > 200 ? `<tr><td colspan="5" style="text-align:center;padding:12px;font-size:13px;color:var(--text-muted);">... dan ${imageUrls.length - 200} lainnya</td></tr>` : '');

  const chipsWrap = document.getElementById('preview-chips');
  chipsWrap.style.display = 'flex';
  chipsWrap.innerHTML = '<span style="font-size:12px; color:var(--text-muted); line-height:28px;">Sample URL ke-</span>';
  [0, 1, 2, 3, 4].filter(i => i < imageUrls.length).forEach(i => {
    const chip = document.createElement('button');
    chip.className = 'chip' + (i === 0 ? ' active' : '');
    chip.textContent = i + 1;
    chip.onclick = () => {
      document.querySelectorAll('#preview-chips .chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      generatePreview(i);
    };
    chipsWrap.appendChild(chip);
  });
}

function toggleRow(i, checked) {
  if (checked) selectedRows.add(i); else selectedRows.delete(i);
}

document.getElementById('chk-all').addEventListener('change', function() {
  const boxes = document.querySelectorAll('#images-tbody input[type=checkbox]');
  boxes.forEach((b, i) => { b.checked = this.checked; if (this.checked) selectedRows.add(i); else selectedRows.delete(i); });
});

function statusLabel(s) {
  return { pending: 'Menunggu', loading: 'Memproses...', done: 'Selesai', error: 'Gagal' }[s] || s;
}

function handleOverlayFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    const img = new Image();
    img.onload = () => {
      overlayImage = img;
      document.getElementById('overlay-img-preview').style.display = 'flex';
      document.getElementById('overlay-preview-img').src = e.target.result;
      document.getElementById('overlay-settings').style.display = 'grid';
      document.getElementById('overlay-notice').style.display = 'block';


      const icon = overlayDrop.querySelector('.drop-icon svg');
      overlayDrop.querySelector('.drop-title').textContent = file.name;
      overlayDrop.querySelector('.drop-sub').textContent = `${img.naturalWidth}×${img.naturalHeight}px`;
      overlayDrop.querySelector('.drop-icon').style.background = 'rgba(124,109,250,0.12)';
      icon.style.stroke = 'var(--accent2)';

      checkReady();
      generatePreview();
    };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

function checkReady() {
  const ready = imageUrls.length > 0 && overlayImage !== null;
  document.getElementById('btn-process').disabled = !ready;
  if (ready) setStep(3);
}

function setStep(n) {
  [1,2,3,4].forEach(i => {
    const num = document.getElementById('snum-' + i);
    const lbl = document.getElementById('slbl-' + i);
    if (i < n) { num.className = 'step-num done'; num.innerHTML = '✓'; }
    else if (i === n) { num.className = 'step-num active'; num.textContent = i; lbl.className = 'step-label active'; }
    else { num.className = 'step-num'; num.textContent = i; lbl.className = 'step-label'; }
  });
}

async function generatePreview(idx = 0) {
  if (!overlayImage || !imageUrls.length) return;
  const url = imageUrls[idx]?.url;
  if (!url) return;

  const canvas = document.getElementById('preview-canvas');
  const placeholder = document.getElementById('preview-placeholder');
  placeholder.style.display = 'none';
  canvas.style.display = 'block';

  try {
    const baseImg = await loadImage(url);
    const result = await compositeImages(baseImg);
    const ctx = canvas.getContext('2d');
    canvas.width = result.width;
    canvas.height = result.height;
    ctx.drawImage(result, 0, 0);
  } catch (err) {
    placeholder.style.display = 'block';
    placeholder.querySelector('div').textContent = 'Gagal load gambar sample.';
    canvas.style.display = 'none';
  }
}

function loadImage(url) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => res(img);
    img.onerror = () => {
      // try with cors proxy
      const proxy = `https://corsproxy.io/?${encodeURIComponent(url)}`;
      const img2 = new Image();
      img2.crossOrigin = 'anonymous';
      img2.onload = () => res(img2);
      img2.onerror = () => rej(new Error('Cannot load: ' + url));
      img2.src = proxy;
    };
    img.src = url;
  });
}

function updateOverlayNotice() {
  const mode = document.getElementById('overlay-mode').value;
  const ratioSelect = document.getElementById('output-ratio');
  const ratioText = ratioSelect ? ratioSelect.options[ratioSelect.selectedIndex].text : '1:1';
  const notice = document.getElementById('overlay-notice');
  const descriptions = {
    'append-right': `<strong>Append Right:</strong> Overlay ditempelkan di sisi kanan, penuh atas ke bawah. Output ${ratioText}.`,
    'append-left': `<strong>Append Left:</strong> Overlay ditempelkan di sisi kiri, penuh atas ke bawah. Output ${ratioText}.`,
    'append-top': `<strong>Append Top:</strong> Overlay ditempelkan di atas, penuh kiri ke kanan. Output ${ratioText}.`,
    'append-bottom': `<strong>Append Bottom:</strong> Overlay ditempelkan di bawah, penuh kiri ke kanan. Output ${ratioText}.`,
    'center': `<strong>Center:</strong> Overlay ditindih di tengah gambar produk. Output ${ratioText}.`
  };
  notice.innerHTML = descriptions[mode] || '';
}

function compositeImages(baseImg) {
  return new Promise(res => {
    const gap = parseInt(document.getElementById('overlay-gap').value) || 0;
    const mode = document.getElementById('overlay-mode').value;
    const targetAR = parseFloat(document.getElementById('output-ratio')?.value) || 1;

    const bw = baseImg.naturalWidth || baseImg.width;
    const bh = baseImg.naturalHeight || baseImg.height;
    const oNW = overlayImage.naturalWidth;
    const oNH = overlayImage.naturalHeight;
    const overlayAR = oNW / oNH;
    const overlayARinv = oNH / oNW; // height/width

    let canvasW, canvasH, baseX, baseY, ovX, ovY, ovW, ovH;

    if (mode === 'append-right') {
      let baseH = Math.max(bh, Math.ceil((bw + gap) / targetAR));
      let reqH = baseH;
      if (targetAR > overlayAR) {
        reqH = Math.max(baseH, Math.ceil((bw + gap) / (targetAR - overlayAR)));
      }
      canvasH = reqH;
      canvasW = Math.round(canvasH * targetAR);

      ovH = canvasH;
      ovW = Math.round(overlayAR * canvasH);
      const maxOW = canvasW - bw - gap;
      if (ovW > maxOW && maxOW > 0) { ovW = maxOW; ovH = canvasH; }

      baseX = 0;
      baseY = Math.round((canvasH - bh) / 2);
      ovX = canvasW - ovW;
      ovY = 0;

    } else if (mode === 'append-left') {
      let baseH = Math.max(bh, Math.ceil((bw + gap) / targetAR));
      let reqH = baseH;
      if (targetAR > overlayAR) {
        reqH = Math.max(baseH, Math.ceil((bw + gap) / (targetAR - overlayAR)));
      }
      canvasH = reqH;
      canvasW = Math.round(canvasH * targetAR);

      ovH = canvasH;
      ovW = Math.round(overlayAR * canvasH);
      const maxOW = canvasW - bw - gap;
      if (ovW > maxOW && maxOW > 0) { ovW = maxOW; ovH = canvasH; }

      ovX = 0;
      ovY = 0;
      baseX = canvasW - bw;
      baseY = Math.round((canvasH - bh) / 2);

    } else if (mode === 'append-top') {
      const targetARinv = 1 / targetAR;
      let baseW = Math.max(bw, Math.ceil((bh + gap) / targetARinv));
      let reqW = baseW;
      if (targetARinv > overlayARinv) {
        reqW = Math.max(baseW, Math.ceil((bh + gap) / (targetARinv - overlayARinv)));
      }
      canvasW = reqW;
      canvasH = Math.round(canvasW * targetARinv);

      ovW = canvasW;
      ovH = Math.round(overlayARinv * canvasW);
      const maxOH = canvasH - bh - gap;
      if (ovH > maxOH && maxOH > 0) { ovH = maxOH; ovW = canvasW; }

      ovX = 0;
      ovY = 0;
      baseX = Math.round((canvasW - bw) / 2);
      baseY = canvasH - bh;

    } else if (mode === 'append-bottom') {
      const targetARinv = 1 / targetAR;
      let baseW = Math.max(bw, Math.ceil((bh + gap) / targetARinv));
      let reqW = baseW;
      if (targetARinv > overlayARinv) {
        reqW = Math.max(baseW, Math.ceil((bh + gap) / (targetARinv - overlayARinv)));
      }
      canvasW = reqW;
      canvasH = Math.round(canvasW * targetARinv);

      ovW = canvasW;
      ovH = Math.round(overlayARinv * canvasW);
      const maxOH = canvasH - bh - gap;
      if (ovH > maxOH && maxOH > 0) { ovH = maxOH; ovW = canvasW; }

      baseX = Math.round((canvasW - bw) / 2);
      baseY = 0;
      ovX = 0;
      ovY = canvasH - ovH;

    } else {
      // Center: overlay di atas gambar produk
      if (bw / bh > targetAR) {
        canvasW = bw;
        canvasH = Math.round(bw / targetAR);
      } else {
        canvasH = bh;
        canvasW = Math.round(bh * targetAR);
      }
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

    // Background putih
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, canvasW, canvasH);

    // Draw base image
    ctx.drawImage(baseImg, baseX, baseY, bw, bh);

    // Draw overlay
    ctx.drawImage(overlayImage, ovX, ovY, ovW, ovH);

    res(c);
  });
}

function getUniqueFilename(url, fmt, usedNames) {
  try {
    // Extract pathname from URL, removing query params and hash
    let pathname = new URL(url, 'https://x.com').pathname;
    // Get just the filename part (last segment of path)
    let baseName = pathname.split('/').pop() || 'image';
    // Remove the original extension
    const dotIdx = baseName.lastIndexOf('.');
    if (dotIdx > 0) {
      baseName = baseName.substring(0, dotIdx);
    }
    // Sanitize: remove characters not safe for filenames
    baseName = baseName.replace(/[^a-zA-Z0-9_\-\.]/g, '_').replace(/_+/g, '_');
    if (!baseName) baseName = 'image';
    let filename = `${baseName}.${fmt}`;
    // Handle duplicates by appending _2, _3, etc.
    let counter = 2;
    while (usedNames.has(filename)) {
      filename = `${baseName}_${counter}.${fmt}`;
      counter++;
    }
    return filename;
  } catch (e) {
    // Fallback if URL parsing fails
    let fallback = `image_${Date.now()}.${fmt}`;
    return fallback;
  }
}

async function startProcessing() {
  if (!imageUrls.length || !overlayImage) return;

  const btn = document.getElementById('btn-process');
  btn.disabled = true;
  btn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg> Memproses...';

  const progressArea = document.getElementById('progress-area');
  progressArea.classList.add('visible');
  const progressFill = document.getElementById('progress-fill');
  const progressText = document.getElementById('progress-text');
  const progressPct = document.getElementById('progress-pct');
  const progressLog = document.getElementById('progress-log');
  progressLog.innerHTML = '';

  const zip = new JSZip();
  const fmt = document.getElementById('output-fmt').value;
  const mimeType = fmt === 'jpeg' ? 'image/jpeg' : fmt === 'webp' ? 'image/webp' : 'image/png';

  const selectedList = imageUrls.filter((_, i) => selectedRows.has(i));
  const usedFilenames = new Set();
  let done = 0, errors = 0;

  function log(msg, cls = '') {
    const line = document.createElement('div');
    if (cls) line.className = cls;
    line.textContent = msg;
    progressLog.appendChild(line);
    progressLog.scrollTop = progressLog.scrollHeight;
  }

  log(`[INFO] Mulai memproses ${selectedList.length} gambar...`, 'log-info');
  setStep(4);

  const CONCURRENCY = 4;
  let idx = 0;

  async function processNext() {
    while (idx < selectedList.length) {
      const i = idx++;
      const item = selectedList[i];
      const globalIdx = imageUrls.indexOf(item);

      // Update status
      item.status = 'loading';
      updateRowStatus(globalIdx, 'loading');

      try {
        const baseImg = await loadImage(item.url);
        const composite = await compositeImages(baseImg);
        const dataUrl = composite.toDataURL(mimeType, 0.92);
        const base64 = dataUrl.split(',')[1];
        const filename = getUniqueFilename(item.url, fmt, usedFilenames);
        usedFilenames.add(filename);
        zip.file(filename, base64, { base64: true });
        item.status = 'done';
        updateRowStatus(globalIdx, 'done');
        log(`[OK] ${filename} — ${composite.width}×${composite.height}px`, 'log-ok');
      } catch (err) {
        item.status = 'error';
        updateRowStatus(globalIdx, 'error');
        errors++;
        log(`[ERR] Baris ${item.row}: ${err.message}`, 'log-err');
      }

      done++;
      const pct = Math.round((done / selectedList.length) * 100);
      progressFill.style.width = pct + '%';
      progressPct.textContent = pct + '%';
      progressText.textContent = `Memproses ${done} / ${selectedList.length}`;
    }
  }

  const workers = Array.from({ length: CONCURRENCY }, () => processNext());
  await Promise.all(workers);

  log(`[INFO] Selesai! ${done - errors} berhasil, ${errors} gagal.`, 'log-info');
  progressText.textContent = `Selesai — Membuat ZIP...`;

  const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  const zipUrl = URL.createObjectURL(zipBlob);
  const a = document.createElement('a');
  a.href = zipUrl;
  a.download = `imageforge_${Date.now()}.zip`;
  a.click();
  URL.revokeObjectURL(zipUrl);

  progressText.textContent = `ZIP didownload! ${done - errors} gambar berhasil.`;
  log(`[INFO] ZIP berhasil didownload.`, 'log-info');

  btn.disabled = false;
  btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Download ZIP Lagi`;

  const statEl = document.getElementById('excel-stats');
  statEl.innerHTML += `<div class="stat-mini-item"><div class="stat-mini-val" style="color:var(--success)">${done - errors}</div><div class="stat-mini-lbl">Berhasil</div></div>`;
  if (errors) statEl.innerHTML += `<div class="stat-mini-item"><div class="stat-mini-val" style="color:var(--danger)">${errors}</div><div class="stat-mini-lbl">Gagal</div></div>`;
}

function updateRowStatus(i, status) {
  const dot = document.getElementById('sdot-' + i);
  const txt = document.getElementById('stxt-' + i);
  if (dot) { dot.className = 'status-dot ' + status; }
  if (txt) { txt.textContent = statusLabel(status); }
}

function resetAll() {
  excelData = [];
  imageUrls = [];
  overlayImage = null;
  selectedRows = new Set();
  selectedColumns.clear();

  excelFileInput.value = '';
  overlayFileInput.value = '';

  document.getElementById('col-select-wrap').style.display = 'none';
  document.getElementById('excel-info').style.display = 'none';
  document.getElementById('overlay-img-preview').style.display = 'none';
  document.getElementById('overlay-settings').style.display = 'none';
  document.getElementById('overlay-notice').style.display = 'none';
  document.getElementById('table-section').style.display = 'none';
  document.getElementById('preview-section').style.display = 'none';
  document.getElementById('progress-area').classList.remove('visible');

  document.getElementById('images-tbody').innerHTML = `<tr><td colspan="5"><div class="empty-state">
    <svg viewBox="0 0 24 24" fill="none" stroke-width="1.5" stroke-linecap="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
    Upload file Excel terlebih dahulu
  </div></td></tr>`;

  document.getElementById('btn-process').disabled = true;
  document.getElementById('btn-process').innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg> Proses &amp; Download ZIP`;

  const excelIcon = excelDrop.querySelector('.drop-icon svg');
  excelDrop.querySelector('.drop-title').textContent = 'Drag & drop file Excel';
  excelDrop.querySelector('.drop-sub').textContent = 'atau klik untuk browse — .xlsx, .xls, .csv';
  excelDrop.querySelector('.drop-icon').style.background = 'var(--surface3)';
  excelIcon.style.stroke = 'var(--text-muted)';

  const ovIcon = overlayDrop.querySelector('.drop-icon svg');
  overlayDrop.querySelector('.drop-title').textContent = 'Drag & drop gambar overlay';
  overlayDrop.querySelector('.drop-sub').textContent = 'PNG, JPG, SVG, WebP — transparan didukung';
  overlayDrop.querySelector('.drop-icon').style.background = 'var(--surface3)';
  ovIcon.style.stroke = 'var(--text-muted)';

  const previewCanvas = document.getElementById('preview-canvas');
  previewCanvas.style.display = 'none';
  document.getElementById('preview-placeholder').style.display = 'block';
  document.getElementById('preview-placeholder').querySelector('div').textContent = 'Preview akan muncul di sini';

  setStep(1);
  document.querySelectorAll('.step-num').forEach((n, i) => {
    n.textContent = i + 1;
    n.className = 'step-num' + (i === 0 ? ' active' : '');
  });
  document.querySelectorAll('.step-label').forEach((l, i) => {
    l.className = 'step-label' + (i === 0 ? ' active' : '');
  });
  document.getElementById('snum-1').textContent = '1';
}

// ==========================================
// OVERLAY STUDIO JS LOGIC
// ==========================================

let osBaseImages = [];
let osLayerImage = null;

function switchMainTab(tab) {
  document.getElementById('tab-bulk').className = 'tab-btn' + (tab === 'bulk' ? ' active' : '');
  document.getElementById('tab-os').className = 'tab-btn' + (tab === 'os' ? ' active' : '');
  document.getElementById('view-bulk').style.display = tab === 'bulk' ? 'block' : 'none';
  document.getElementById('view-os').style.display = tab === 'os' ? 'block' : 'none';
}

function updateOsUI() {
  const mode = document.getElementById('os-mode').value;
  const posGroup = document.getElementById('os-position-group');
  const yGroup = document.getElementById('os-offset-y-group');
  const xLabel = document.getElementById('os-offset-x-label');

  if (mode === 'overlay') {
    posGroup.style.display = 'block';
    yGroup.style.display = 'block';
    xLabel.textContent = 'Offset X (px)';
  } else {
    posGroup.style.display = 'none';
    yGroup.style.display = 'none';
    xLabel.textContent = 'Gap / Jarak (px)';
  }
}

// Handlers for Layer Drop
const osLayerDrop = document.getElementById('os-layer-drop');
const osLayerFile = document.getElementById('os-layer-file');

osLayerDrop.addEventListener('dragover', e => { e.preventDefault(); osLayerDrop.classList.add('drag-over'); });
osLayerDrop.addEventListener('dragleave', () => osLayerDrop.classList.remove('drag-over'));
osLayerDrop.addEventListener('drop', e => { e.preventDefault(); osLayerDrop.classList.remove('drag-over'); handleOsLayer(e.dataTransfer.files[0]); });
osLayerFile.addEventListener('change', e => { if (e.target.files.length) handleOsLayer(e.target.files[0]); });

function handleOsLayer(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    osLayerImage = new Image();
    osLayerImage.onload = () => {
      document.getElementById('os-layer-preview-img').src = e.target.result;
      document.getElementById('os-layer-preview').style.display = 'flex';
      checkOsReady();
    };
    osLayerImage.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

// Handlers for Base Images Drop
const osBaseDrop = document.getElementById('os-base-drop');
const osBaseFile = document.getElementById('os-base-file');

osBaseDrop.addEventListener('dragover', e => { e.preventDefault(); osBaseDrop.classList.add('drag-over'); });
osBaseDrop.addEventListener('dragleave', () => osBaseDrop.classList.remove('drag-over'));
osBaseDrop.addEventListener('drop', e => { e.preventDefault(); osBaseDrop.classList.remove('drag-over'); handleOsBase(e.dataTransfer.files); });
osBaseFile.addEventListener('change', e => { if (e.target.files.length) handleOsBase(e.target.files); });

function handleOsBase(files) {
  Array.from(files).forEach(file => {
    if(!file.type.startsWith('image/')) return;
    const reader = new FileReader();
    reader.onload = e => {
      const img = new Image();
      img.onload = () => {
        osBaseImages.push({ file, img, src: e.target.result });
        renderOsThumbs();
        checkOsReady();
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

function renderOsThumbs() {
  const container = document.getElementById('os-base-thumbs');
  container.innerHTML = '';
  osBaseImages.forEach((item, i) => {
    const d = document.createElement('div');
    d.className = 'thumb-item';
    d.innerHTML = `<img src="${item.src}"/> <div class="thumb-overlay" onclick="removeOsBase(${i})" style="cursor:pointer; color:#fff;">×</div>`;
    container.appendChild(d);
  });
}

function removeOsBase(index) {
  osBaseImages.splice(index, 1);
  renderOsThumbs();
  checkOsReady();
}

function checkOsReady() {
  const btnZip = document.getElementById('btn-os-process-zip');
  const btnDirect = document.getElementById('btn-os-process-direct');
  if (osBaseImages.length > 0 && osLayerImage) {
    if (btnZip) btnZip.disabled = false;
    if (btnDirect) btnDirect.disabled = false;
    generateOsPreview();
  } else {
    if (btnZip) btnZip.disabled = true;
    if (btnDirect) btnDirect.disabled = true;
  }
}

async function generateOsPreview() {
  if (!osLayerImage || osBaseImages.length === 0) return;
  
  const container = document.getElementById('os-multi-preview-container');
  const ph = document.getElementById('os-preview-placeholder');
  
  container.innerHTML = '';
  
  for (let i = 0; i < osBaseImages.length; i++) {
    const canvas = await compositeOsImage(osBaseImages[i].img);
    container.appendChild(canvas);
  }
  
  container.style.display = 'grid';
  ph.style.display = 'none';
}

function compositeOsImage(baseImg) {
  return new Promise(res => {
    const mode = document.getElementById('os-mode').value;
    const gapX = parseInt(document.getElementById('os-offset-x').value) || 0;
    const gapY = parseInt(document.getElementById('os-offset-y').value) || 0;
    const position = document.getElementById('os-position').value;
    
    const bw = baseImg.naturalWidth || baseImg.width;
    const bh = baseImg.naturalHeight || baseImg.height;
    const lw = osLayerImage.naturalWidth || osLayerImage.width;
    const lh = osLayerImage.naturalHeight || osLayerImage.height;

    let canvasW, canvasH, baseX, baseY, ovX, ovY, ovW, ovH;

    if (mode === 'overlay') {
      canvasW = bw; canvasH = bh;
      baseX = 0; baseY = 0;
      
      if (position === 'fill') {
        ovW = bw; ovH = bh; ovX = 0; ovY = 0;
      } else {
        ovW = lw; ovH = lh;
        if (position === 'center') { ovX = (bw - lw)/2 + gapX; ovY = (bh - lh)/2 + gapY; }
        else if (position === 'top-left') { ovX = gapX; ovY = gapY; }
        else if (position === 'top-right') { ovX = bw - lw - gapX; ovY = gapY; }
        else if (position === 'bottom-left') { ovX = gapX; ovY = bh - lh - gapY; }
        else if (position === 'bottom-right') { ovX = bw - lw - gapX; ovY = bh - lh - gapY; }
      }
    } else {
      // Append modes (gapX is used as gap constraint)
      const gap = Math.max(0, gapX);
      if (mode === 'append-right') {
        canvasW = bw + lw + gap;
        canvasH = Math.max(bh, lh);
        baseX = 0; baseY = (canvasH - bh)/2;
        ovX = bw + gap; ovY = (canvasH - lh)/2;
        ovW = lw; ovH = lh;
      }
      else if (mode === 'append-left') {
        canvasW = bw + lw + gap;
        canvasH = Math.max(bh, lh);
        baseX = lw + gap; baseY = (canvasH - bh)/2;
        ovX = 0; ovY = (canvasH - lh)/2;
        ovW = lw; ovH = lh;
      }
      else if (mode === 'append-top') {
        canvasH = bh + lh + gap;
        canvasW = Math.max(bw, lw);
        baseX = (canvasW - bw)/2; baseY = lh + gap;
        ovX = (canvasW - lw)/2; ovY = 0;
        ovW = lw; ovH = lh;
      }
      else if (mode === 'append-bottom') {
        canvasH = bh + lh + gap;
        canvasW = Math.max(bw, lw);
        baseX = (canvasW - bw)/2; baseY = 0;
        ovX = (canvasW - lw)/2; ovY = bh + gap;
        ovW = lw; ovH = lh;
      }
    }

    const c = document.createElement('canvas');
    c.width = canvasW; c.height = canvasH;
    const ctx = c.getContext('2d');

    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, canvasW, canvasH);
    ctx.drawImage(baseImg, baseX, baseY, bw, bh);
    ctx.drawImage(osLayerImage, ovX, ovY, ovW, ovH);
    
    res(c);
  });
}

function osLog(msg, type='info') {
  const area = document.getElementById('os-progress-log');
  area.innerHTML += `<div class="log-${type}">[${new Date().toLocaleTimeString()}] ${msg}</div>`;
  area.scrollTop = area.scrollHeight;
}

async function startOsProcessing(downloadMode = 'zip') {
  if (osBaseImages.length === 0 || !osLayerImage) return;

  const btnZip = document.getElementById('btn-os-process-zip');
  const btnDirect = document.getElementById('btn-os-process-direct');
  btnZip.disabled = true;
  btnDirect.disabled = true;

  document.getElementById('os-progress-area').classList.add('visible');
  document.getElementById('os-progress-log').innerHTML = '';
  document.getElementById('os-progress-fill').style.width = '0%';
  document.getElementById('os-progress-pct').textContent = '0%';

  const zip = new JSZip();
  const format = document.getElementById('os-format').value;
  const mimeType = 'image/' + (format === 'jpg' ? 'jpeg' : format);
  let ext = format === 'jpeg' ? 'jpg' : format;

  osLog(`Memulai pemrosesan ${osBaseImages.length} file (Mode: ${downloadMode})...`);

  for (let i = 0; i < osBaseImages.length; i++) {
    const item = osBaseImages[i];
    try {
      osLog(`Memproses: ${item.file.name}`);
      const composite = await compositeOsImage(item.img);
      const dataUrl = composite.toDataURL(mimeType, 0.92);
      const oname = item.file.name.replace(/\.[^/.]+$/, "") + '_result.' + ext;

      if (downloadMode === 'zip') {
        const base64 = dataUrl.split(',')[1];
        zip.file(oname, base64, { base64: true });
      } else {
        // Direct download
        const a = document.createElement('a');
        a.href = dataUrl;
        a.download = oname;
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        // Delay singkat agar browser tidak memblokir multiple download secera bersamaan
        await new Promise(r => setTimeout(r, 150));
      }

      osLog(`Selesai: ${item.file.name}`, 'ok');
    } catch (e) {
      osLog(`Gagal: ${item.file.name} - ${e.message}`, 'err');
    }
    const pct = Math.round(((i + 1) / osBaseImages.length) * 100);
    document.getElementById('os-progress-fill').style.width = pct + '%';
    document.getElementById('os-progress-pct').textContent = pct + '%';
  }

  if (downloadMode === 'zip') {
    osLog(`Menyiapkan file ZIP...`);
    try {
      const content = await zip.generateAsync({ type: "blob" });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(content);
      a.download = `overlay_studio_export.zip`;
      a.click();
      URL.revokeObjectURL(a.href);
      osLog(`File ZIP berhasil diunduh!`, 'ok');
    } catch (err) {
      osLog(`Gagal membuat ZIP: ${err.message}`, 'err');
    }
  } else {
    osLog(`Semua file berhasil diproses & diunduh!`, 'ok');
  }

  btnZip.disabled = false;
  btnDirect.disabled = false;
}

// Initial UI setup
updateOsUI();
