let excelData = [];
let imageUrls = [];
let overlayImage = null;
let selectedRows = new Set();
let bulkInputType = "excel";
let archiveBaseImages = [];

const excelDrop = document.getElementById("excel-drop");
const bulkArchiveDrop = document.getElementById("bulk-archive-drop");
const overlayDrop = document.getElementById("overlay-drop");
const excelFileInput = document.getElementById("excel-file");
const bulkArchiveFileInput = document.getElementById("bulk-archive-file");
const overlayFileInput = document.getElementById("overlay-file");

["dragenter", "dragover"].forEach((e) => {
  excelDrop.addEventListener(e, (ev) => {
    ev.preventDefault();
    excelDrop.classList.add("drag-over");
  });
  bulkArchiveDrop.addEventListener(e, (ev) => {
    ev.preventDefault();
    bulkArchiveDrop.classList.add("drag-over");
  });
  overlayDrop.addEventListener(e, (ev) => {
    ev.preventDefault();
    overlayDrop.classList.add("drag-over");
  });
});
["dragleave", "drop"].forEach((e) => {
  excelDrop.addEventListener(e, () => excelDrop.classList.remove("drag-over"));
  bulkArchiveDrop.addEventListener(e, () =>
    bulkArchiveDrop.classList.remove("drag-over"),
  );
  overlayDrop.addEventListener(e, () =>
    overlayDrop.classList.remove("drag-over"),
  );
});

excelDrop.addEventListener("drop", (ev) => {
  ev.preventDefault();
  const f = ev.dataTransfer.files[0];
  if (f) handleExcelFile(f);
});
bulkArchiveDrop.addEventListener("drop", (ev) => {
  ev.preventDefault();
  const f = ev.dataTransfer.files[0];
  if (f) handleBulkArchive(f);
});
overlayDrop.addEventListener("drop", (ev) => {
  ev.preventDefault();
  const f = ev.dataTransfer.files[0];
  if (f) handleOverlayFile(f);
});

excelFileInput.addEventListener("change", (e) => {
  if (e.target.files[0]) handleExcelFile(e.target.files[0]);
});
bulkArchiveFileInput.addEventListener("change", (e) => {
  if (e.target.files[0]) handleBulkArchive(e.target.files[0]);
});
overlayFileInput.addEventListener("change", (e) => {
  if (e.target.files[0]) handleOverlayFile(e.target.files[0]);
});

let selectedColumns = new Set();

function switchBulkInputType(type) {
  bulkInputType = type;
  const exDrop = document.getElementById("excel-drop");
  const arDrop = document.getElementById("bulk-archive-drop");

  document
    .querySelectorAll('input[name="bulk-input-type"]')
    .forEach((input) => {
      input.checked = input.value === type;
    });

  if (type === "excel") {
    exDrop.style.display = "block";
    arDrop.style.display = "none";
    document.getElementById("bulk-input-title").textContent =
      "Upload File Excel";
    document.getElementById("bulk-input-desc").textContent =
      "File .xlsx atau .xls berisi kolom dengan URL gambar CDN.";

    if (excelData.length > 0) {
      document.getElementById("col-select-wrap").style.display = "flex";
      document.getElementById("excel-info").style.display = "block";
      onColumnSelected();
    } else {
      document.getElementById("col-select-wrap").style.display = "none";
      document.getElementById("excel-info").style.display = "none";
      imageUrls = [];
      selectedRows = new Set();
      renderTable();
    }
  } else {
    exDrop.style.display = "none";
    arDrop.style.display = "block";
    document.getElementById("bulk-input-title").textContent =
      "Upload File ZIP / RAR";
    document.getElementById("bulk-input-desc").textContent =
      "File archive .zip atau .rar berisi file gambar offline (PNG, JPG, JPEG, WebP).";
    document.getElementById("col-select-wrap").style.display = "none";

    if (archiveBaseImages.length > 0) {
      document.getElementById("excel-info").style.display = "block";
      imageUrls = [...archiveBaseImages];
      selectedRows = new Set(imageUrls.map((_, i) => i));
      renderTable();
    } else {
      document.getElementById("excel-info").style.display = "none";
      imageUrls = [];
      selectedRows = new Set();
      renderTable();
    }
  }
  checkReady();
}

async function handleBulkArchive(file) {
  if (!file) return;

  const stats = document.getElementById("excel-stats");
  document.getElementById("excel-info").style.display = "block";
  stats.innerHTML = `
    <div class="stat-mini-item" style="flex: 1;">
      <div class="stat-mini-lbl" style="font-size: 13px; color: var(--text);">Mengekstrak archive... silakan tunggu</div>
    </div>
  `;

  try {
    const arrayBuffer = await file.arrayBuffer();
    const extractedFiles = [];
    const isRar = file.name.toLowerCase().endsWith(".rar");

    if (isRar) {
      try {
        const extractor = new Unrar(arrayBuffer);
        const fileList = extractor.getFileList();
        for (const fileItem of fileList) {
          if (fileItem.type === "file") {
            const ext = fileItem.name.split(".").pop().toLowerCase();
            if (["png", "jpg", "jpeg", "webp"].includes(ext)) {
              const extractedData = extractor.extract(fileItem.name);
              const blob = new Blob([extractedData], {
                type: `image/${ext === "jpg" ? "jpeg" : ext}`,
              });
              extractedFiles.push({ name: fileItem.name, blob });
            }
          }
        }
      } catch (rarErr) {
        throw new Error(
          "Gagal mengekstrak RAR: " +
            rarErr.message +
            ". Pastikan file RAR tidak dipassword.",
        );
      }
    } else {
      const zip = new JSZip();
      const content = await zip.loadAsync(arrayBuffer);
      const paths = Object.keys(content.files).filter((path) => {
        const ext = path.split(".").pop().toLowerCase();
        return (
          ["png", "jpg", "jpeg", "webp"].includes(ext) &&
          !content.files[path].dir
        );
      });
      for (const path of paths) {
        const blob = await content.files[path].async("blob");
        extractedFiles.push({ name: path, blob });
      }
    }

    if (extractedFiles.length === 0) {
      throw new Error(
        "Tidak ada file gambar valid (PNG, JPG, WebP) ditemukan di dalam archive.",
      );
    }

    // Clean up existing Object URLs
    archiveBaseImages.forEach((item) => {
      if (item.url && item.url.startsWith("blob:")) {
        URL.revokeObjectURL(item.url);
      }
    });

    archiveBaseImages = extractedFiles.map((item, idx) => {
      const url = URL.createObjectURL(item.blob);
      return {
        url: url,
        filename: item.name.split("/").pop(),
        row: idx + 1,
        col: "Archive",
        status: "pending",
      };
    });

    imageUrls = [...archiveBaseImages];
    selectedRows = new Set(imageUrls.map((_, i) => i));

    stats.innerHTML = `
      <div class="stat-mini-item"><div class="stat-mini-val">${extractedFiles.length}</div><div class="stat-mini-lbl">Gambar Ditemukan</div></div>
      <div class="stat-mini-item"><div class="stat-mini-val">${(file.size / (1024 * 1024)).toFixed(2)} MB</div><div class="stat-mini-lbl">Ukuran Archive</div></div>
    `;

    const dropZone = document.getElementById("bulk-archive-drop");
    const icon = dropZone.querySelector(".drop-icon svg");
    dropZone.querySelector(".drop-title").textContent = file.name;
    dropZone.querySelector(".drop-sub").textContent =
      `${extractedFiles.length} gambar ditemukan`;
    dropZone.querySelector(".drop-icon").style.background =
      "rgba(200,245,80,0.1)";
    icon.style.stroke = "var(--accent)";

    renderTable();
    document.getElementById("badge-count").textContent =
      imageUrls.length + " gambar";
    document.getElementById("tbl-count-badge").textContent = imageUrls.length;
    document.getElementById("table-section").style.display = "block";
    document.getElementById("preview-section").style.display = "block";

    checkReady();
  } catch (err) {
    stats.innerHTML = `
      <div class="stat-mini-item" style="flex: 1;">
        <div class="stat-mini-val" style="color:var(--danger); font-size:14px;">Error</div>
        <div class="stat-mini-lbl">${err.message}</div>
      </div>
    `;
    alert("Gagal mengekstrak archive: " + err.message);
  }
}

function handleExcelFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (!rows.length) {
        alert("File kosong atau tidak valid.");
        return;
      }

      excelData = rows;
      const headers = rows[0].map((h, i) => h || `Kolom ${i + 1}`);
      selectedColumns.clear();

      const colUrlCounts = headers.map((_, colIdx) => {
        let count = 0;
        for (let r = 1; r < Math.min(rows.length, 30); r++) {
          const val = String(rows[r][colIdx] || "").trim();
          if (val.startsWith("http") || val.startsWith("//")) count++;
        }
        return count;
      });

      const chipsContainer = document.getElementById("col-chips");
      chipsContainer.innerHTML = "";
      headers.forEach((h, i) => {
        const urlCount = colUrlCounts[i];
        const isUrl = urlCount > 0;
        const chip = document.createElement("button");
        chip.className = "chip" + (isUrl ? " active" : "");
        chip.dataset.colIdx = i;
        chip.innerHTML =
          `${h}` +
          (urlCount > 0
            ? ` <span style="font-size:10px;opacity:0.7;">(${urlCount} URL)</span>`
            : "");
        if (isUrl) selectedColumns.add(i);

        chip.addEventListener("click", () => {
          if (selectedColumns.has(i)) {
            selectedColumns.delete(i);
            chip.classList.remove("active");
          } else {
            selectedColumns.add(i);
            chip.classList.add("active");
          }
          onColumnSelected();
        });
        chipsContainer.appendChild(chip);
      });

      document.getElementById("col-select-wrap").style.display = "flex";
      document.getElementById("excel-info").style.display = "block";

      const stats = document.getElementById("excel-stats");
      stats.innerHTML = `
        <div class="stat-mini-item"><div class="stat-mini-val">${rows.length - 1}</div><div class="stat-mini-lbl">Baris data</div></div>
        <div class="stat-mini-item"><div class="stat-mini-val">${headers.length}</div><div class="stat-mini-lbl">Kolom</div></div>
      `;

      const excelIcon = excelDrop.querySelector(".drop-icon svg");
      excelDrop.querySelector(".drop-title").textContent = file.name;
      excelDrop.querySelector(".drop-sub").textContent =
        `${(file.size / 1024).toFixed(1)} KB — ${wb.SheetNames[0]}`;
      excelDrop.querySelector(".drop-icon").style.background =
        "rgba(200,245,80,0.1)";
      excelIcon.style.stroke = "var(--accent)";

      if (selectedColumns.size > 0) onColumnSelected();

      setStep(2);
    } catch (err) {
      alert("Gagal membaca file: " + err.message);
    }
  };
  reader.readAsBinaryString(file);
}

function onColumnSelected() {
  if (selectedColumns.size === 0) {
    imageUrls = [];
    renderTable();
    document.getElementById("badge-count").textContent = "0 gambar";
    document.getElementById("tbl-count-badge").textContent = "0";
    return;
  }

  const headers = excelData[0].map((h, i) => h || `Kolom ${i + 1}`);
  imageUrls = [];
  for (let i = 1; i < excelData.length; i++) {
    for (const colIdx of selectedColumns) {
      const val = String(excelData[i][colIdx] || "").trim();
      if (val && (val.startsWith("http") || val.startsWith("//"))) {
        imageUrls.push({
          url: val,
          row: i,
          col: headers[colIdx],
          status: "pending",
        });
      }
    }
  }

  selectedRows = new Set(imageUrls.map((_, i) => i));
  renderTable();
  document.getElementById("badge-count").textContent =
    imageUrls.length + " gambar";
  document.getElementById("tbl-count-badge").textContent = imageUrls.length;
  document.getElementById("table-section").style.display = "block";
  document.getElementById("preview-section").style.display = "block";
  checkReady();
}

function renderTable() {
  const tbody = document.getElementById("images-tbody");
  if (!imageUrls.length) {
    const emptyMsg =
      bulkInputType === "excel"
        ? "Tidak ada URL valid ditemukan di kolom yang dipilih."
        : "Tidak ada gambar yang dimuat dari archive.";
    tbody.innerHTML = `<tr><td colspan="5"><div class="empty-state">${emptyMsg}</div></td></tr>`;
    return;
  }
  tbody.innerHTML =
    imageUrls
      .slice(0, 200)
      .map((item, i) => {
        const isLocal = !!item.filename;
        const urlDisplay = isLocal ? `📦 ${item.filename}` : item.url;
        return `
      <tr>
        <td><input type="checkbox" ${selectedRows.has(i) ? "checked" : ""} onchange="toggleRow(${i}, this.checked)" style="accent-color:var(--accent); cursor:pointer;"></td>
        <td style="color:var(--text-muted); font-size:12px;">${item.row}</td>
        <td><span class="badge badge-purple" style="font-size:10px;">${item.col || "-"}</span></td>
        <td><a href="${item.url}" target="_blank" class="url-cell" title="${isLocal ? item.filename : item.url}">${urlDisplay}</a></td>
        <td><span class="status-dot ${item.status}" id="sdot-${i}"></span><span id="stxt-${i}" style="font-size:12px;">${statusLabel(item.status)}</span></td>
      </tr>
    `;
      })
      .join("") +
    (imageUrls.length > 200
      ? `<tr><td colspan="5" style="text-align:center;padding:12px;font-size:13px;color:var(--text-muted);">... dan ${imageUrls.length - 200} lainnya</td></tr>`
      : "");

  const chipsWrap = document.getElementById("preview-chips");
  chipsWrap.style.display = "flex";
  chipsWrap.innerHTML =
    '<span style="font-size:12px; color:var(--text-muted); line-height:28px;">Sample Gambar ke-</span>';
  [0, 1, 2, 3, 4]
    .filter((i) => i < imageUrls.length)
    .forEach((i) => {
      const chip = document.createElement("button");
      chip.className = "chip" + (i === 0 ? " active" : "");
      chip.textContent = i + 1;
      chip.onclick = () => {
        document
          .querySelectorAll("#preview-chips .chip")
          .forEach((c) => c.classList.remove("active"));
        chip.classList.add("active");
        generatePreview(i);
      };
      chipsWrap.appendChild(chip);
    });
}

function toggleRow(i, checked) {
  if (checked) selectedRows.add(i);
  else selectedRows.delete(i);
}

document.getElementById("chk-all").addEventListener("change", function () {
  const boxes = document.querySelectorAll("#images-tbody input[type=checkbox]");
  boxes.forEach((b, i) => {
    b.checked = this.checked;
    if (this.checked) selectedRows.add(i);
    else selectedRows.delete(i);
  });
});

function statusLabel(s) {
  return (
    {
      pending: "Menunggu",
      loading: "Memproses...",
      done: "Selesai",
      error: "Gagal",
    }[s] || s
  );
}

function handleOverlayFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const img = new Image();
    img.onload = () => {
      overlayImage = img;
      document.getElementById("overlay-img-preview").style.display = "flex";
      document.getElementById("overlay-preview-img").src = e.target.result;
      document.getElementById("overlay-settings").style.display = "grid";
      document.getElementById("overlay-notice").style.display = "block";

      const icon = overlayDrop.querySelector(".drop-icon svg");
      overlayDrop.querySelector(".drop-title").textContent = file.name;
      overlayDrop.querySelector(".drop-sub").textContent =
        `${img.naturalWidth}×${img.naturalHeight}px`;
      overlayDrop.querySelector(".drop-icon").style.background =
        "rgba(124,109,250,0.12)";
      icon.style.stroke = "var(--accent2)";

      checkReady();
      generatePreview();
    };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

function checkReady() {
  const ready = imageUrls.length > 0 && overlayImage !== null;
  document.getElementById("btn-process").disabled = !ready;
  if (ready) setStep(3);
}

function setStep(n) {
  [1, 2, 3, 4].forEach((i) => {
    const num = document.getElementById("snum-" + i);
    const lbl = document.getElementById("slbl-" + i);
    if (i < n) {
      num.className = "step-num done";
      num.innerHTML = "✓";
    } else if (i === n) {
      num.className = "step-num active";
      num.textContent = i;
      lbl.className = "step-label active";
    } else {
      num.className = "step-num";
      num.textContent = i;
      lbl.className = "step-label";
    }
  });
}

async function generatePreview(idx = 0) {
  if (!overlayImage || !imageUrls.length) return;
  const url = imageUrls[idx]?.url;
  if (!url) return;

  const canvas = document.getElementById("preview-canvas");
  const placeholder = document.getElementById("preview-placeholder");
  placeholder.style.display = "none";
  canvas.style.display = "block";

  try {
    const baseImg = await loadImage(url);
    const result = await compositeImages(baseImg);
    const ctx = canvas.getContext("2d");
    canvas.width = result.width;
    canvas.height = result.height;
    ctx.drawImage(result, 0, 0);
  } catch (err) {
    placeholder.style.display = "block";
    placeholder.querySelector("div").textContent = "Gagal load gambar sample.";
    canvas.style.display = "none";
  }
}

function loadImage(url) {
  return new Promise((res, rej) => {
    const img = new Image();
    if (!url.startsWith("blob:") && !url.startsWith("data:")) {
      img.crossOrigin = "anonymous";
    }
    img.onload = () => res(img);
    img.onerror = () => {
      if (url.startsWith("blob:") || url.startsWith("data:")) {
        rej(new Error("Cannot load local image: " + url));
        return;
      }
      const proxy = `https://corsproxy.io/?${encodeURIComponent(url)}`;
      const img2 = new Image();
      img2.crossOrigin = "anonymous";
      img2.onload = () => res(img2);
      img2.onerror = () => rej(new Error("Cannot load: " + url));
      img2.src = proxy;
    };
    img.src = url;
  });
}

function updateOverlayNotice() {
  const mode = document.getElementById("overlay-mode").value;
  const ratioSelect = document.getElementById("output-ratio");
  const ratioText = ratioSelect
    ? ratioSelect.options[ratioSelect.selectedIndex].text
    : "1:1";
  const notice = document.getElementById("overlay-notice");
  const descriptions = {
    "append-right": `<strong>Append Right:</strong> Overlay ditempelkan di sisi kanan, penuh atas ke bawah. Output ${ratioText}.`,
    "append-left": `<strong>Append Left:</strong> Overlay ditempelkan di sisi kiri, penuh atas ke bawah. Output ${ratioText}.`,
    "append-top": `<strong>Append Top:</strong> Overlay ditempelkan di atas, penuh kiri ke kanan. Output ${ratioText}.`,
    "append-bottom": `<strong>Append Bottom:</strong> Overlay ditempelkan di bawah, penuh kiri ke kanan. Output ${ratioText}.`,
    center: `<strong>Center:</strong> Overlay ditindih di tengah gambar produk. Output ${ratioText}.`,
  };
  notice.innerHTML = descriptions[mode] || "";
}

function compositeImages(baseImg) {
  return new Promise((res) => {
    const gap = parseInt(document.getElementById("overlay-gap").value) || 0;
    const zoomPct =
      parseFloat(document.getElementById("product-zoom")?.value) || 100;
    const mode = document.getElementById("overlay-mode").value;
    const targetAR =
      parseFloat(document.getElementById("output-ratio")?.value) || 1;

    const origBw = baseImg.naturalWidth || baseImg.width;
    const origBh = baseImg.naturalHeight || baseImg.height;
    const bw = Math.round(origBw * (zoomPct / 100));
    const bh = Math.round(origBh * (zoomPct / 100));
    const oNW = overlayImage.naturalWidth;
    const oNH = overlayImage.naturalHeight;
    const overlayAR = oNW / oNH;
    const overlayARinv = oNH / oNW; // height/width

    let canvasW, canvasH, baseX, baseY, ovX, ovY, ovW, ovH;

    if (mode === "append-right") {
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
      if (ovW > maxOW && maxOW > 0) {
        ovW = maxOW;
        ovH = canvasH;
      }

      baseX = 0;
      baseY = Math.round((canvasH - bh) / 2);
      ovX = canvasW - ovW;
      ovY = 0;
    } else if (mode === "append-left") {
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
      if (ovW > maxOW && maxOW > 0) {
        ovW = maxOW;
        ovH = canvasH;
      }

      ovX = 0;
      ovY = 0;
      baseX = canvasW - bw;
      baseY = Math.round((canvasH - bh) / 2);
    } else if (mode === "append-top") {
      const targetARinv = 1 / targetAR;
      let baseW = Math.max(bw, Math.ceil((bh + gap) / targetARinv));
      let reqW = baseW;
      if (targetARinv > overlayARinv) {
        reqW = Math.max(
          baseW,
          Math.ceil((bh + gap) / (targetARinv - overlayARinv)),
        );
      }
      canvasW = reqW;
      canvasH = Math.round(canvasW * targetARinv);

      ovW = canvasW;
      ovH = Math.round(overlayARinv * canvasW);
      const maxOH = canvasH - bh - gap;
      if (ovH > maxOH && maxOH > 0) {
        ovH = maxOH;
        ovW = canvasW;
      }

      ovX = 0;
      ovY = 0;
      baseX = Math.round((canvasW - bw) / 2);
      baseY = canvasH - bh;
    } else if (mode === "append-bottom") {
      const targetARinv = 1 / targetAR;
      let baseW = Math.max(bw, Math.ceil((bh + gap) / targetARinv));
      let reqW = baseW;
      if (targetARinv > overlayARinv) {
        reqW = Math.max(
          baseW,
          Math.ceil((bh + gap) / (targetARinv - overlayARinv)),
        );
      }
      canvasW = reqW;
      canvasH = Math.round(canvasW * targetARinv);

      ovW = canvasW;
      ovH = Math.round(overlayARinv * canvasW);
      const maxOH = canvasH - bh - gap;
      if (ovH > maxOH && maxOH > 0) {
        ovH = maxOH;
        ovW = canvasW;
      }

      baseX = Math.round((canvasW - bw) / 2);
      baseY = 0;
      ovX = 0;
      ovY = canvasH - ovH;
    } else {
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

    const c = document.createElement("canvas");
    c.width = canvasW;
    c.height = canvasH;
    const ctx = c.getContext("2d");
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, canvasW, canvasH);
    ctx.drawImage(baseImg, baseX, baseY, bw, bh);
    ctx.drawImage(overlayImage, ovX, ovY, ovW, ovH);

    res(c);
  });
}

function getUniqueFilename(url, fmt, usedNames, customFilename = null) {
  try {
    let baseName = "image";
    if (customFilename) {
      baseName = customFilename;
    } else {
      let pathname = new URL(url, "https://x.com").pathname;
      baseName = pathname.split("/").pop() || "image";
    }
    const dotIdx = baseName.lastIndexOf(".");
    if (dotIdx > 0) {
      baseName = baseName.substring(0, dotIdx);
    }
    baseName = baseName.replace(/[^a-zA-Z0-9_\-\.]/g, "_").replace(/_+/g, "_");
    if (!baseName) baseName = "image";
    let filename = `${baseName}.${fmt}`;
    let counter = 2;
    while (usedNames.has(filename)) {
      filename = `${baseName}_${counter}.${fmt}`;
      counter++;
    }
    return filename;
  } catch (e) {
    let fallback = `image_${Date.now()}.${fmt}`;
    return fallback;
  }
}

async function startProcessing() {
  if (!imageUrls.length || !overlayImage) return;

  const btn = document.getElementById("btn-process");
  btn.disabled = true;
  btn.innerHTML =
    '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg> Memproses...';

  const progressArea = document.getElementById("progress-area");
  progressArea.classList.add("visible");
  const progressFill = document.getElementById("progress-fill");
  const progressText = document.getElementById("progress-text");
  const progressPct = document.getElementById("progress-pct");
  const progressLog = document.getElementById("progress-log");
  progressLog.innerHTML = "";

  const zip = new JSZip();
  const fmt = document.getElementById("output-fmt").value;
  const mimeType =
    fmt === "jpeg" ? "image/jpeg" : fmt === "webp" ? "image/webp" : "image/png";

  const selectedList = imageUrls.filter((_, i) => selectedRows.has(i));
  const usedFilenames = new Set();
  let done = 0,
    errors = 0;

  function log(msg, cls = "") {
    const line = document.createElement("div");
    if (cls) line.className = cls;
    line.textContent = msg;
    progressLog.appendChild(line);
    progressLog.scrollTop = progressLog.scrollHeight;
  }

  log(`[INFO] Mulai memproses ${selectedList.length} gambar...`, "log-info");
  setStep(4);

  const CONCURRENCY = 4;
  let idx = 0;

  async function processNext() {
    while (idx < selectedList.length) {
      const i = idx++;
      const item = selectedList[i];
      const globalIdx = imageUrls.indexOf(item);
      item.status = "loading";
      updateRowStatus(globalIdx, "loading");

      try {
        const baseImg = await loadImage(item.url);
        const composite = await compositeImages(baseImg);
        const dataUrl = composite.toDataURL(mimeType, 0.92);
        const base64 = dataUrl.split(",")[1];
        const filename = getUniqueFilename(
          item.url,
          fmt,
          usedFilenames,
          item.filename,
        );
        usedFilenames.add(filename);
        zip.file(filename, base64, { base64: true });
        item.status = "done";
        updateRowStatus(globalIdx, "done");
        log(
          `[OK] ${filename} — ${composite.width}×${composite.height}px`,
          "log-ok",
        );
      } catch (err) {
        item.status = "error";
        updateRowStatus(globalIdx, "error");
        errors++;
        log(`[ERR] Baris ${item.row}: ${err.message}`, "log-err");
      }

      done++;
      const pct = Math.round((done / selectedList.length) * 100);
      progressFill.style.width = pct + "%";
      progressPct.textContent = pct + "%";
      progressText.textContent = `Memproses ${done} / ${selectedList.length}`;
    }
  }

  const workers = Array.from({ length: CONCURRENCY }, () => processNext());
  await Promise.all(workers);

  log(
    `[INFO] Selesai! ${done - errors} berhasil, ${errors} gagal.`,
    "log-info",
  );
  progressText.textContent = `Selesai — Membuat ZIP...`;

  const zipBlob = await zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  const zipUrl = URL.createObjectURL(zipBlob);
  const a = document.createElement("a");
  a.href = zipUrl;
  a.download = `imageforge_${Date.now()}.zip`;
  a.click();
  URL.revokeObjectURL(zipUrl);

  progressText.textContent = `ZIP didownload! ${done - errors} gambar berhasil.`;
  log(`[INFO] ZIP berhasil didownload.`, "log-info");

  btn.disabled = false;
  btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Download ZIP Lagi`;

  const statEl = document.getElementById("excel-stats");
  statEl.innerHTML += `<div class="stat-mini-item"><div class="stat-mini-val" style="color:var(--success)">${done - errors}</div><div class="stat-mini-lbl">Berhasil</div></div>`;
  if (errors)
    statEl.innerHTML += `<div class="stat-mini-item"><div class="stat-mini-val" style="color:var(--danger)">${errors}</div><div class="stat-mini-lbl">Gagal</div></div>`;
}

function updateRowStatus(i, status) {
  const dot = document.getElementById("sdot-" + i);
  const txt = document.getElementById("stxt-" + i);
  if (dot) {
    dot.className = "status-dot " + status;
  }
  if (txt) {
    txt.textContent = statusLabel(status);
  }
}

function resetAll() {
  excelData = [];
  imageUrls = [];
  overlayImage = null;
  selectedRows = new Set();
  selectedColumns.clear();

  // Clear archive Object URLs
  archiveBaseImages.forEach((item) => {
    if (item.url && item.url.startsWith("blob:")) {
      URL.revokeObjectURL(item.url);
    }
  });
  archiveBaseImages = [];

  excelFileInput.value = "";
  document.getElementById("bulk-archive-file").value = "";
  overlayFileInput.value = "";

  document.getElementById("col-select-wrap").style.display = "none";
  document.getElementById("excel-info").style.display = "none";
  document.getElementById("overlay-img-preview").style.display = "none";
  document.getElementById("overlay-settings").style.display = "none";
  document.getElementById("overlay-notice").style.display = "none";
  document.getElementById("table-section").style.display = "none";
  document.getElementById("preview-section").style.display = "none";
  document.getElementById("progress-area").classList.remove("visible");

  document.getElementById("images-tbody").innerHTML =
    `<tr><td colspan="5"><div class="empty-state">
    <svg viewBox="0 0 24 24" fill="none" stroke-width="1.5" stroke-linecap="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
    Upload file Excel atau ZIP/RAR terlebih dahulu
  </div></td></tr>`;

  document.getElementById("btn-process").disabled = true;
  document.getElementById("btn-process").innerHTML =
    `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg> Proses &amp; Download ZIP`;

  const excelIcon = excelDrop.querySelector(".drop-icon svg");
  excelDrop.querySelector(".drop-title").textContent = "Drag & drop file Excel";
  excelDrop.querySelector(".drop-sub").textContent =
    "atau klik untuk browse — .xlsx, .xls, .csv";
  excelDrop.querySelector(".drop-icon").style.background = "var(--surface3)";
  excelIcon.style.stroke = "var(--text-muted)";

  const archiveDropZone = document.getElementById("bulk-archive-drop");
  const archiveIcon = archiveDropZone.querySelector(".drop-icon svg");
  archiveDropZone.querySelector(".drop-title").textContent =
    "Drag & drop file ZIP / RAR";
  archiveDropZone.querySelector(".drop-sub").textContent =
    "atau klik untuk browse — .zip, .rar";
  archiveDropZone.querySelector(".drop-icon").style.background =
    "var(--surface3)";
  archiveIcon.style.stroke = "var(--text-muted)";

  const ovIcon = overlayDrop.querySelector(".drop-icon svg");
  overlayDrop.querySelector(".drop-title").textContent =
    "Drag & drop gambar overlay";
  overlayDrop.querySelector(".drop-sub").textContent =
    "PNG, JPG, SVG, WebP — transparan didukung";
  overlayDrop.querySelector(".drop-icon").style.background = "var(--surface3)";
  ovIcon.style.stroke = "var(--text-muted)";

  const previewCanvas = document.getElementById("preview-canvas");
  previewCanvas.style.display = "none";
  document.getElementById("preview-placeholder").style.display = "block";
  document
    .getElementById("preview-placeholder")
    .querySelector("div").textContent = "Preview akan muncul di sini";

  setStep(1);
  document.querySelectorAll(".step-num").forEach((n, i) => {
    n.textContent = i + 1;
    n.className = "step-num" + (i === 0 ? " active" : "");
  });
  document.querySelectorAll(".step-label").forEach((l, i) => {
    l.className = "step-label" + (i === 0 ? " active" : "");
  });
  document.getElementById("snum-1").textContent = "1";
}

let osBaseImages = [];
let osLayerImage = null;

function switchMainTab(tab) {
  document.getElementById("tab-bulk").className =
    "tab-btn" + (tab === "bulk" ? " active" : "");
  document.getElementById("tab-os").className =
    "tab-btn" + (tab === "os" ? " active" : "");
  const viewOnlyBulk = document.getElementById("view-only-bulk");
  if (viewOnlyBulk)
    viewOnlyBulk.style.display = tab === "only-bulk" ? "block" : "none";

  const tabResize = document.getElementById("tab-resize");
  if (tabResize)
    tabResize.className = "tab-btn" + (tab === "resize" ? " active" : "");
  const viewResize = document.getElementById("view-resize");
  if (viewResize)
    viewResize.style.display = tab === "resize" ? "block" : "none";
}

function updateOsUI() {
  const mode = document.getElementById("os-mode").value;
  const posGroup = document.getElementById("os-position-group");
  const yGroup = document.getElementById("os-offset-y-group");
  const xLabel = document.getElementById("os-offset-x-label");

  if (mode === "overlay") {
    posGroup.style.display = "block";
    yGroup.style.display = "block";
    xLabel.textContent = "Offset X (px)";
  } else {
    posGroup.style.display = "none";
    yGroup.style.display = "none";
    xLabel.textContent = "Gap / Jarak (px)";
  }
}

const osLayerDrop = document.getElementById("os-layer-drop");
const osLayerFile = document.getElementById("os-layer-file");

osLayerDrop.addEventListener("dragover", (e) => {
  e.preventDefault();
  osLayerDrop.classList.add("drag-over");
});
osLayerDrop.addEventListener("dragleave", () =>
  osLayerDrop.classList.remove("drag-over"),
);
osLayerDrop.addEventListener("drop", (e) => {
  e.preventDefault();
  osLayerDrop.classList.remove("drag-over");
  handleOsLayer(e.dataTransfer.files[0]);
});
osLayerFile.addEventListener("change", (e) => {
  if (e.target.files.length) handleOsLayer(e.target.files[0]);
});

function handleOsLayer(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    osLayerImage = new Image();
    osLayerImage.onload = () => {
      document.getElementById("os-layer-preview-img").src = e.target.result;
      document.getElementById("os-layer-preview").style.display = "flex";
      checkOsReady();
    };
    osLayerImage.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

const osBaseDrop = document.getElementById("os-base-drop");
const osBaseFile = document.getElementById("os-base-file");

osBaseDrop.addEventListener("dragover", (e) => {
  e.preventDefault();
  osBaseDrop.classList.add("drag-over");
});
osBaseDrop.addEventListener("dragleave", () =>
  osBaseDrop.classList.remove("drag-over"),
);
osBaseDrop.addEventListener("drop", (e) => {
  e.preventDefault();
  osBaseDrop.classList.remove("drag-over");
  handleOsBase(e.dataTransfer.files);
});
osBaseFile.addEventListener("change", (e) => {
  if (e.target.files.length) handleOsBase(e.target.files);
});

function handleOsBase(files) {
  if (!files || files.length === 0) return;
  const file = files[0]; // Hanya ambil satu file
  if (!file.type.startsWith("image/")) {
    alert("Harap pilih file gambar yang valid.");
    return;
  }
  
  const reader = new FileReader();
  reader.onload = (e) => {
    const img = new Image();
    img.onload = () => {
      osBaseImages = [{ file, img, src: e.target.result }];
      renderOsThumbs();
      checkOsReady();
      // Auto-generate preview if layer image is also available
      if (typeof generateOsPreview === "function") {
          generateOsPreview();
      }
    };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

function renderOsThumbs() {
  const container = document.getElementById("os-base-thumbs");
  container.innerHTML = "";
  osBaseImages.forEach((item, i) => {
    const d = document.createElement("div");
    d.className = "thumb-item";
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
  const btnZip = document.getElementById("btn-os-process-zip");
  const btnDirect = document.getElementById("btn-os-process-direct");
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

  const container = document.getElementById("os-multi-preview-container");
  const ph = document.getElementById("os-preview-placeholder");

  container.innerHTML = "";

  for (let i = 0; i < osBaseImages.length; i++) {
    const canvas = await compositeOsImage(osBaseImages[i].img);
    container.appendChild(canvas);
  }

  container.style.display = "grid";
  ph.style.display = "none";
}

function compositeOsImage(baseImg) {
  return new Promise((res) => {
    const mode = document.getElementById("os-mode").value;
    const gapX = parseInt(document.getElementById("os-offset-x").value) || 0;
    const gapY = parseInt(document.getElementById("os-offset-y").value) || 0;
    const position = document.getElementById("os-position").value;

    const bw = baseImg.naturalWidth || baseImg.width;
    const bh = baseImg.naturalHeight || baseImg.height;
    const lw = osLayerImage.naturalWidth || osLayerImage.width;
    const lh = osLayerImage.naturalHeight || osLayerImage.height;

    let canvasW, canvasH, baseX, baseY, ovX, ovY, ovW, ovH;

    if (mode === "overlay") {
      canvasW = bw;
      canvasH = bh;
      baseX = 0;
      baseY = 0;

      if (position === "fill") {
        ovW = bw;
        ovH = bh;
        ovX = 0;
        ovY = 0;
      } else {
        ovW = lw;
        ovH = lh;
        if (position === "center") {
          ovX = (bw - lw) / 2 + gapX;
          ovY = (bh - lh) / 2 + gapY;
        } else if (position === "top-left") {
          ovX = gapX;
          ovY = gapY;
        } else if (position === "top-right") {
          ovX = bw - lw - gapX;
          ovY = gapY;
        } else if (position === "bottom-left") {
          ovX = gapX;
          ovY = bh - lh - gapY;
        } else if (position === "bottom-right") {
          ovX = bw - lw - gapX;
          ovY = bh - lh - gapY;
        }
      }
    } else {
      // Append modes (gapX is used as gap constraint)
      const gap = Math.max(0, gapX);
      if (mode === "append-right") {
        canvasW = bw + lw + gap;
        canvasH = Math.max(bh, lh);
        baseX = 0;
        baseY = (canvasH - bh) / 2;
        ovX = bw + gap;
        ovY = (canvasH - lh) / 2;
        ovW = lw;
        ovH = lh;
      } else if (mode === "append-left") {
        canvasW = bw + lw + gap;
        canvasH = Math.max(bh, lh);
        baseX = lw + gap;
        baseY = (canvasH - bh) / 2;
        ovX = 0;
        ovY = (canvasH - lh) / 2;
        ovW = lw;
        ovH = lh;
      } else if (mode === "append-top") {
        canvasH = bh + lh + gap;
        canvasW = Math.max(bw, lw);
        baseX = (canvasW - bw) / 2;
        baseY = lh + gap;
        ovX = (canvasW - lw) / 2;
        ovY = 0;
        ovW = lw;
        ovH = lh;
      } else if (mode === "append-bottom") {
        canvasH = bh + lh + gap;
        canvasW = Math.max(bw, lw);
        baseX = (canvasW - bw) / 2;
        baseY = 0;
        ovX = (canvasW - lw) / 2;
        ovY = bh + gap;
        ovW = lw;
        ovH = lh;
      }
    }

    const c = document.createElement("canvas");
    c.width = canvasW;
    c.height = canvasH;
    const ctx = c.getContext("2d");

    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, canvasW, canvasH);
    ctx.drawImage(baseImg, baseX, baseY, bw, bh);
    ctx.drawImage(osLayerImage, ovX, ovY, ovW, ovH);

    res(c);
  });
}

function osLog(msg, type = "info") {
  const area = document.getElementById("os-progress-log");
  area.innerHTML += `<div class="log-${type}">[${new Date().toLocaleTimeString()}] ${msg}</div>`;
  area.scrollTop = area.scrollHeight;
}

async function startOsProcessing(downloadMode = "zip") {
  if (osBaseImages.length === 0 || !osLayerImage) return;

  const btnZip = document.getElementById("btn-os-process-zip");
  const btnDirect = document.getElementById("btn-os-process-direct");
  btnZip.disabled = true;
  btnDirect.disabled = true;

  document.getElementById("os-progress-area").classList.add("visible");
  document.getElementById("os-progress-log").innerHTML = "";
  document.getElementById("os-progress-fill").style.width = "0%";
  document.getElementById("os-progress-pct").textContent = "0%";

  const zip = new JSZip();
  const format = document.getElementById("os-format").value;
  const mimeType = "image/" + (format === "jpg" ? "jpeg" : format);
  let ext = format === "jpeg" ? "jpg" : format;

  osLog(
    `Memulai pemrosesan ${osBaseImages.length} file (Mode: ${downloadMode})...`,
  );

  for (let i = 0; i < osBaseImages.length; i++) {
    const item = osBaseImages[i];
    try {
      osLog(`Memproses: ${item.file.name}`);
      const composite = await compositeOsImage(item.img);
      const dataUrl = composite.toDataURL(mimeType, 0.92);
      const oname = item.file.name.replace(/\.[^/.]+$/, "") + "_result." + ext;

      if (downloadMode === "zip") {
        const base64 = dataUrl.split(",")[1];
        zip.file(oname, base64, { base64: true });
      } else {
        // Direct download
        const a = document.createElement("a");
        a.href = dataUrl;
        a.download = oname;
        a.style.display = "none";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        // Delay singkat agar browser tidak memblokir multiple download secera bersamaan
        await new Promise((r) => setTimeout(r, 150));
      }

      osLog(`Selesai: ${item.file.name}`, "ok");
    } catch (e) {
      osLog(`Gagal: ${item.file.name} - ${e.message}`, "err");
    }
    const pct = Math.round(((i + 1) / osBaseImages.length) * 100);
    document.getElementById("os-progress-fill").style.width = pct + "%";
    document.getElementById("os-progress-pct").textContent = pct + "%";
  }

  if (downloadMode === "zip") {
    osLog(`Menyiapkan file ZIP...`);
    try {
      const content = await zip.generateAsync({ type: "blob" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(content);
      a.download = `overlay_studio_export.zip`;
      a.click();
      URL.revokeObjectURL(a.href);
      osLog(`File ZIP berhasil diunduh!`, "ok");
    } catch (err) {
      osLog(`Gagal membuat ZIP: ${err.message}`, "err");
    }
  } else {
    osLog(`Semua file berhasil diproses & diunduh!`, "ok");
  }

  btnZip.disabled = false;
  btnDirect.disabled = false;
}

// Initial UI setup
updateOsUI();

// ==========================================
// BULK DOWNLOAD ONLY LOGIC
// ==========================================

let bdExcelData = [];
let bdImageUrls = [];
let bdSelectedRows = new Set();
let bdSelectedColumns = new Set();

const bdExcelDrop = document.getElementById("bd-excel-drop");
const bdExcelFileInput = document.getElementById("bd-excel-file");

["dragenter", "dragover"].forEach((e) => {
  bdExcelDrop.addEventListener(e, (ev) => {
    ev.preventDefault();
    bdExcelDrop.classList.add("drag-over");
  });
});
["dragleave", "drop"].forEach((e) => {
  bdExcelDrop.addEventListener(e, () =>
    bdExcelDrop.classList.remove("drag-over"),
  );
});

bdExcelDrop.addEventListener("drop", (ev) => {
  ev.preventDefault();
  const f = ev.dataTransfer.files[0];
  if (f) bdHandleExcelFile(f);
});

bdExcelFileInput.addEventListener("change", (e) => {
  if (e.target.files[0]) bdHandleExcelFile(e.target.files[0]);
});

function bdHandleExcelFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (!rows.length) {
        alert("File kosong atau tidak valid.");
        return;
      }

      bdExcelData = rows;
      const headers = rows[0].map((h, i) => h || `Kolom ${i + 1}`);
      bdSelectedColumns.clear();

      const colUrlCounts = headers.map((_, colIdx) => {
        let count = 0;
        for (let r = 1; r < Math.min(rows.length, 30); r++) {
          const val = String(rows[r][colIdx] || "").trim();
          if (val.startsWith("http") || val.startsWith("//")) count++;
        }
        return count;
      });

      const chipsContainer = document.getElementById("bd-col-chips");
      chipsContainer.innerHTML = "";
      headers.forEach((h, i) => {
        const urlCount = colUrlCounts[i];
        const isUrl = urlCount > 0;
        const chip = document.createElement("button");
        chip.className = "chip" + (isUrl ? " active" : "");
        chip.dataset.colIdx = i;
        chip.innerHTML =
          `${h}` +
          (urlCount > 0
            ? ` <span style="font-size:10px;opacity:0.7;">(${urlCount} URL)</span>`
            : "");
        if (isUrl) bdSelectedColumns.add(i);

        chip.addEventListener("click", () => {
          if (bdSelectedColumns.has(i)) {
            bdSelectedColumns.delete(i);
            chip.classList.remove("active");
          } else {
            bdSelectedColumns.add(i);
            chip.classList.add("active");
          }
          bdOnColumnSelected();
        });
        chipsContainer.appendChild(chip);
      });

      document.getElementById("bd-col-select-wrap").style.display = "flex";
      document.getElementById("bd-excel-info").style.display = "block";

      const stats = document.getElementById("bd-excel-stats");
      stats.innerHTML = `
        <div class="stat-mini-item"><div class="stat-mini-val">${rows.length - 1}</div><div class="stat-mini-lbl">Baris data</div></div>
        <div class="stat-mini-item"><div class="stat-mini-val">${headers.length}</div><div class="stat-mini-lbl">Kolom</div></div>
      `;

      const excelIcon = bdExcelDrop.querySelector(".drop-icon svg");
      bdExcelDrop.querySelector(".drop-title").textContent = file.name;
      bdExcelDrop.querySelector(".drop-sub").textContent =
        `${(file.size / 1024).toFixed(1)} KB — ${wb.SheetNames[0]}`;
      bdExcelDrop.querySelector(".drop-icon").style.background =
        "rgba(200,245,80,0.1)";
      excelIcon.style.stroke = "var(--accent)";

      if (bdSelectedColumns.size > 0) bdOnColumnSelected();
      document.getElementById("bd-snum-2").className = "step-num active";
    } catch (err) {
      alert("Gagal membaca file: " + err.message);
    }
  };
  reader.readAsBinaryString(file);
}

function bdOnColumnSelected() {
  if (bdSelectedColumns.size === 0) {
    bdImageUrls = [];
    bdRenderTable();
    document.getElementById("bd-badge-count").textContent = "0 gambar";
    document.getElementById("bd-tbl-count-badge").textContent = "0";
    return;
  }

  const headers = bdExcelData[0].map((h, i) => h || `Kolom ${i + 1}`);
  bdImageUrls = [];
  for (let i = 1; i < bdExcelData.length; i++) {
    for (const colIdx of bdSelectedColumns) {
      // Find existing object with URL
      const val = String(bdExcelData[i][colIdx] || "").trim();
      if (val && (val.startsWith("http") || val.startsWith("//"))) {
        bdImageUrls.push({
          url: val,
          row: i,
          col: headers[colIdx],
          status: "pending",
        });
      }
    }
  }

  bdSelectedRows = new Set(bdImageUrls.map((_, i) => i));
  bdRenderTable();
  document.getElementById("bd-badge-count").textContent =
    bdImageUrls.length + " gambar";
  document.getElementById("bd-tbl-count-badge").textContent =
    bdImageUrls.length;
  document.getElementById("bd-table-section").style.display = "block";
  document.getElementById("bd-settings-section").style.display = "block";
  bdCheckReady();
}

function bdRenderTable() {
  const tbody = document.getElementById("bd-images-tbody");
  if (!bdImageUrls.length) {
    tbody.innerHTML = `<tr><td colspan="5"><div class="empty-state">Tidak ada URL valid ditemukan di kolom yang dipilih.</div></td></tr>`;
    return;
  }
  tbody.innerHTML =
    bdImageUrls
      .slice(0, 200)
      .map(
        (item, i) => `
    <tr>
      <td><input type="checkbox" ${bdSelectedRows.has(i) ? "checked" : ""} onchange="bdToggleRow(${i}, this.checked)" style="accent-color:var(--accent); cursor:pointer;"></td>
      <td style="color:var(--text-muted); font-size:12px;">${item.row}</td>
      <td><span class="badge badge-purple" style="font-size:10px;">${item.col || "-"}</span></td>
      <td><a href="${item.url}" target="_blank" class="url-cell" title="${item.url}">${item.url}</a></td>
      <td><span class="status-dot ${item.status}" id="bd-sdot-${i}"></span><span id="bd-stxt-${i}" style="font-size:12px;">${statusLabel(item.status)}</span></td>
    </tr>
  `,
      )
      .join("") +
    (bdImageUrls.length > 200
      ? `<tr><td colspan="5" style="text-align:center;padding:12px;font-size:13px;color:var(--text-muted);">... dan ${bdImageUrls.length - 200} lainnya</td></tr>`
      : "");
}

function bdToggleRow(i, checked) {
  if (checked) bdSelectedRows.add(i);
  else bdSelectedRows.delete(i);
  document.getElementById("bd-tbl-count-badge").textContent =
    bdSelectedRows.size;
  bdCheckReady();
}

document.getElementById("bd-chk-all").addEventListener("change", function () {
  const boxes = document.querySelectorAll(
    "#bd-images-tbody input[type=checkbox]",
  );
  boxes.forEach((b, i) => {
    b.checked = this.checked;
    if (this.checked) bdSelectedRows.add(i);
    else bdSelectedRows.delete(i);
  });
  document.getElementById("bd-tbl-count-badge").textContent =
    bdSelectedRows.size;
  bdCheckReady();
});

function bdCheckReady() {
  const ready = bdSelectedRows.size > 0;
  document.getElementById("bd-btn-process").disabled = !ready;
}

function bdGetOriginalMimeType(url) {
  const ext = url.split(".").pop().toLowerCase().split("?")[0];
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "png") return "image/png";
  if (ext === "webp") return "image/webp";
  if (ext === "gif") return "image/gif";
  return "image/jpeg"; // fallback
}

function bdGetOriginalExt(url) {
  const ext = url.split(".").pop().toLowerCase().split("?")[0];
  if (ext === "jpg" || ext === "jpeg") return "jpg";
  if (ext === "png") return "png";
  if (ext === "webp") return "webp";
  if (ext === "gif") return "gif";
  return "jpg"; // fallback
}

async function bdStartProcessing() {
  if (!bdSelectedRows.size) return;

  const btn = document.getElementById("bd-btn-process");
  btn.disabled = true;
  btn.innerHTML =
    '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg> Memproses...';

  const progressArea = document.getElementById("bd-progress-area");
  progressArea.classList.add("visible");
  const progressFill = document.getElementById("bd-progress-fill");
  const progressText = document.getElementById("bd-progress-text");
  const progressPct = document.getElementById("bd-progress-pct");
  const progressLog = document.getElementById("bd-progress-log");
  progressLog.innerHTML = "";

  const zip = new JSZip();
  const fmt = document.getElementById("bd-output-fmt").value;

  const selectedList = bdImageUrls.filter((_, i) => bdSelectedRows.has(i));
  const usedFilenames = new Set();
  let done = 0,
    errors = 0;

  function log(msg, cls = "") {
    const line = document.createElement("div");
    if (cls) line.className = cls;
    line.textContent = msg;
    progressLog.appendChild(line);
    progressLog.scrollTop = progressLog.scrollHeight;
  }

  log(`[INFO] Mulai mendownload ${selectedList.length} gambar...`, "log-info");

  const CONCURRENCY = 4;
  let idx = 0;

  async function processNext() {
    while (idx < selectedList.length) {
      const i = idx++;
      const item = selectedList[i];
      const globalIdx = bdImageUrls.indexOf(item);
      item.status = "loading";
      bdUpdateRowStatus(globalIdx, "loading");

      try {
        const baseImg = await loadImage(item.url);
        let base64;
        let extToUse;

        if (fmt === "original") {
          const c = document.createElement("canvas");
          c.width = baseImg.naturalWidth || baseImg.width;
          c.height = baseImg.naturalHeight || baseImg.height;
          const ctx = c.getContext("2d");
          ctx.drawImage(baseImg, 0, 0);
          const mType = bdGetOriginalMimeType(item.url);
          extToUse = bdGetOriginalExt(item.url);
          const dataUrl = c.toDataURL(mType, 0.95);
          base64 = dataUrl.split(",")[1];
        } else {
          const c = document.createElement("canvas");
          c.width = baseImg.naturalWidth || baseImg.width;
          c.height = baseImg.naturalHeight || baseImg.height;
          const ctx = c.getContext("2d");
          if (fmt === "jpeg") {
            ctx.fillStyle = "#FFFFFF";
            ctx.fillRect(0, 0, c.width, c.height);
          }
          ctx.drawImage(baseImg, 0, 0);
          const mimeType =
            fmt === "jpeg"
              ? "image/jpeg"
              : fmt === "webp"
                ? "image/webp"
                : "image/png";
          extToUse = fmt;
          const dataUrl = c.toDataURL(mimeType, 0.92);
          base64 = dataUrl.split(",")[1];
        }

        const filename = getUniqueFilename(item.url, extToUse, usedFilenames);
        usedFilenames.add(filename);
        zip.file(filename, base64, { base64: true });

        item.status = "done";
        bdUpdateRowStatus(globalIdx, "done");
        log(`[OK] ${filename} didownload`, "log-ok");
      } catch (err) {
        item.status = "error";
        bdUpdateRowStatus(globalIdx, "error");
        errors++;
        log(`[ERR] Baris ${item.row}: ${err.message}`, "log-err");
      }

      done++;
      const pct = Math.round((done / selectedList.length) * 100);
      progressFill.style.width = pct + "%";
      progressPct.textContent = pct + "%";
      progressText.textContent = `Mendownload ${done} / ${selectedList.length}`;
    }
  }

  const workers = Array.from({ length: CONCURRENCY }, () => processNext());
  await Promise.all(workers);

  log(
    `[INFO] Selesai! ${done - errors} berhasil, ${errors} gagal.`,
    "log-info",
  );
  progressText.textContent = `Selesai — Membuat ZIP...`;

  const zipBlob = await zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  const zipUrl = URL.createObjectURL(zipBlob);
  const a = document.createElement("a");
  a.href = zipUrl;
  a.download = `bulk_download_${Date.now()}.zip`;
  a.click();
  URL.revokeObjectURL(zipUrl);

  progressText.textContent = `ZIP didownload! ${done - errors} gambar berhasil.`;
  log(`[INFO] ZIP berhasil didownload.`, "log-info");

  btn.disabled = false;
  btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Download ZIP Lagi`;
}

function bdUpdateRowStatus(i, status) {
  const dot = document.getElementById("bd-sdot-" + i);
  const txt = document.getElementById("bd-stxt-" + i);
  if (dot) {
    dot.className = "status-dot " + status;
  }
  if (txt) {
    txt.textContent = statusLabel(status);
  }
}

function bdResetAll() {
  bdExcelData = [];
  bdImageUrls = [];
  bdSelectedRows = new Set();
  bdSelectedColumns.clear();

  bdExcelFileInput.value = "";

  document.getElementById("bd-col-select-wrap").style.display = "none";
  document.getElementById("bd-excel-info").style.display = "none";
  document.getElementById("bd-table-section").style.display = "none";
  document.getElementById("bd-settings-section").style.display = "none";
  document.getElementById("bd-progress-area").classList.remove("visible");

  document.getElementById("bd-images-tbody").innerHTML =
    `<tr><td colspan="5"><div class="empty-state">
    <svg viewBox="0 0 24 24" fill="none" stroke-width="1.5" stroke-linecap="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
    Upload file Excel terlebih dahulu
  </div></td></tr>`;

  document.getElementById("bd-btn-process").disabled = true;
  document.getElementById("bd-btn-process").innerHTML =
    `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg> Download ZIP`;

  const excelIcon = bdExcelDrop.querySelector(".drop-icon svg");
  bdExcelDrop.querySelector(".drop-title").textContent =
    "Drag & drop file Excel";
  bdExcelDrop.querySelector(".drop-sub").textContent =
    "atau klik untuk browse — .xlsx, .xls, .csv";
  bdExcelDrop.querySelector(".drop-icon").style.background = "var(--surface3)";
  excelIcon.style.stroke = "var(--text-muted)";

  document.getElementById("bd-snum-2").className = "step-num";
}

// ==========================================
// IMAGE RESIZER LOGIC
// ==========================================

let rsUploadedFile = null;
let rsType = "single";

const rsDropZone = document.getElementById("rs-drop-zone");
const rsFileInput = document.getElementById("rs-file-input");

function updateRsUI() {
  rsType = document.querySelector('input[name="rs-type"]:checked').value;
  const title = document.getElementById("rs-drop-title");
  const sub = document.getElementById("rs-drop-sub");
  const input = document.getElementById("rs-file-input");

  if (rsType === "single") {
    title.textContent = "Drag & drop file gambar";
    sub.textContent = "atau klik untuk browse — PNG, JPG, WebP";
    input.accept = "image/png, image/jpeg, image/jpg, image/webp";
  } else {
    title.textContent = "Drag & drop file RAR / ZIP";
    sub.textContent = "atau klik untuk browse — .zip, .rar";
    input.accept = ".zip, .rar";
  }

  // Clear file if type changed
  rsUploadedFile = null;
  document.getElementById("rs-file-info").style.display = "none";
  document.getElementById("rs-btn-process").disabled = true;
  document
    .getElementById("rs-drop-zone")
    .querySelector(".drop-icon").style.background = "var(--surface3)";
}

["dragenter", "dragover"].forEach((e) => {
  rsDropZone.addEventListener(e, (ev) => {
    ev.preventDefault();
    rsDropZone.classList.add("drag-over");
  });
});
["dragleave", "drop"].forEach((e) => {
  rsDropZone.addEventListener(e, () =>
    rsDropZone.classList.remove("drag-over"),
  );
});

rsDropZone.addEventListener("drop", (ev) => {
  ev.preventDefault();
  const f = ev.dataTransfer.files[0];
  if (f) handleRsFile(f);
});

rsFileInput.addEventListener("change", (e) => {
  if (e.target.files[0]) handleRsFile(e.target.files[0]);
});

function handleRsFile(file) {
  rsUploadedFile = file;
  const stats = document.getElementById("rs-stats");
  document.getElementById("rs-file-info").style.display = "block";

  const sizeMb = (file.size / (1024 * 1024)).toFixed(2);
  stats.innerHTML = `
    <div class="stat-mini-item"><div class="stat-mini-val">${sizeMb} MB</div><div class="stat-mini-lbl">Ukuran Asli</div></div>
    <div class="stat-mini-item"><div class="stat-mini-val">${file.name.split(".").pop().toUpperCase()}</div><div class="stat-mini-lbl">Tipe File</div></div>
  `;

  document.getElementById("rs-drop-title").textContent = file.name;
  document.getElementById("rs-drop-sub").textContent =
    `${(file.size / 1024).toFixed(1)} KB`;
  document
    .getElementById("rs-drop-zone")
    .querySelector(".drop-icon").style.background = "rgba(200,245,80,0.1)";

  document.getElementById("rs-btn-process").disabled = false;
  document.getElementById("rs-snum-2").className = "step-num active";
}

async function rsStartProcessing() {
  if (!rsUploadedFile) return;

  const btn = document.getElementById("rs-btn-process");
  btn.disabled = true;
  btn.innerHTML =
    '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg> Memproses...';

  const progressArea = document.getElementById("rs-progress-area");
  progressArea.classList.add("visible");
  const progressFill = document.getElementById("rs-progress-fill");
  const progressText = document.getElementById("rs-progress-text");
  const progressPct = document.getElementById("rs-progress-pct");
  const progressLog = document.getElementById("rs-progress-log");
  progressLog.innerHTML = "";

  const targetInput = document.getElementById("rs-target-size");
  const targetMB = targetInput ? parseFloat(targetInput.value) || 1 : 1;

  function log(msg, cls = "") {
    const line = document.createElement("div");
    if (cls) line.className = cls;
    line.textContent = msg;
    progressLog.appendChild(line);
    progressLog.scrollTop = progressLog.scrollHeight;
  }

  function formatSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(2) + " MB";
  }

  try {
    log(
      `[INFO] Target ukuran maksimal: ${targetMB} MB (${formatSize(targetMB * 1024 * 1024)})`,
    );

    if (rsType === "single") {
      const originalSize = rsUploadedFile.size;
      log(
        `[INFO] Memproses file tunggal: ${rsUploadedFile.name} (${formatSize(originalSize)})`,
      );
      const result = await resizeImageIfNeeded(rsUploadedFile);
      if (result) {
        const newSize = result.blob.size;
        const reduction = ((1 - newSize / originalSize) * 100).toFixed(1);
        downloadBlob(result.blob, result.filename);
        if (newSize < originalSize) {
          log(
            `[OK] Selesai! ${formatSize(originalSize)} → ${formatSize(newSize)} (berkurang ${reduction}%)`,
            "log-ok",
          );
        } else {
          log(
            `[OK] File sudah di bawah target (${formatSize(newSize)}), tidak perlu dikompres.`,
            "log-ok",
          );
        }
      }
      progressFill.style.width = "100%";
      progressPct.textContent = "100%";
      progressText.textContent = "Selesai!";
    } else {
      log(`[INFO] Membaca archive: ${rsUploadedFile.name}`);
      const arrayBuffer = await rsUploadedFile.arrayBuffer();
      const outputZip = new JSZip();
      let extractedFiles = [];

      const isRar = rsUploadedFile.name.toLowerCase().endsWith(".rar");

      if (isRar) {
        log(`[INFO] Mendeteksi format RAR. Mengekstrak...`);
        try {
          const extractor = new Unrar(arrayBuffer);
          const fileList = extractor.getFileList();
          for (const file of fileList) {
            if (file.type === "file") {
              const ext = file.name.split(".").pop().toLowerCase();
              if (["png", "jpg", "jpeg", "webp"].includes(ext)) {
                const extractedData = extractor.extract(file.name);
                extractedFiles.push({
                  name: file.name,
                  data: new Blob([extractedData], {
                    type: `image/${ext === "jpg" ? "jpeg" : ext}`,
                  }),
                });
              }
            }
          }
        } catch (rarErr) {
          throw new Error(
            "Gagal mengekstrak RAR: " +
              rarErr.message +
              ". Pastikan file RAR tidak dipassword dan bukan RAR5 (beberapa library lama hanya mendukung RAR4).",
          );
        }
      } else {
        log(`[INFO] Mendeteksi format ZIP. Mengekstrak...`);
        const zip = new JSZip();
        const content = await zip.loadAsync(arrayBuffer);
        const paths = Object.keys(content.files).filter((path) => {
          const ext = path.split(".").pop().toLowerCase();
          return (
            ["png", "jpg", "jpeg", "webp"].includes(ext) &&
            !content.files[path].dir
          );
        });
        for (const path of paths) {
          const blob = await content.files[path].async("blob");
          extractedFiles.push({ name: path, data: blob });
        }
      }

      if (extractedFiles.length === 0) {
        throw new Error(
          "Tidak ada file gambar valid (PNG, JPG, WebP) ditemukan di dalam archive.",
        );
      }

      log(`[INFO] Menemukan ${extractedFiles.length} gambar untuk diproses.`);
      let compressedCount = 0;
      let skippedCount = 0;

      for (let i = 0; i < extractedFiles.length; i++) {
        const item = extractedFiles[i];
        const filename = item.name.split("/").pop();
        const originalSize = item.data.size;

        log(
          `[INFO] Memproses (${i + 1}/${extractedFiles.length}): ${filename} (${formatSize(originalSize)})`,
        );
        const result = await resizeImageIfNeeded(item.data, filename);

        if (result) {
          outputZip.file(item.name, result.blob);
          const newSize = result.blob.size;
          if (newSize < originalSize) {
            const reduction = ((1 - newSize / originalSize) * 100).toFixed(1);
            log(
              `[OK] ${filename}: ${formatSize(originalSize)} → ${formatSize(newSize)} (-${reduction}%)`,
              "log-ok",
            );
            compressedCount++;
          } else {
            log(
              `[OK] ${filename}: sudah di bawah target (${formatSize(newSize)})`,
              "log-ok",
            );
            skippedCount++;
          }
        } else {
          outputZip.file(item.name, item.data);
          skippedCount++;
        }

        const pct = Math.round(((i + 1) / extractedFiles.length) * 100);
        progressFill.style.width = pct + "%";
        progressPct.textContent = pct + "%";
        progressText.textContent = `Memproses ${i + 1} / ${extractedFiles.length}`;
      }

      log(
        `[INFO] Ringkasan: ${compressedCount} dikompres, ${skippedCount} sudah OK.`,
      );
      log(`[INFO] Membuat ZIP hasil...`);
      const zipBlob = await outputZip.generateAsync({ type: "blob" });
      downloadBlob(
        zipBlob,
        `resized_${rsUploadedFile.name.replace(/\.[^/.]+$/, "")}.zip`,
      );
      log(
        `[OK] Archive berhasil didownload! (${formatSize(zipBlob.size)})`,
        "log-ok",
      );
    }
  } catch (err) {
    log(`[ERR] ${err.message}`, "log-err");
  }

  btn.disabled = false;
  btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg> Proses &amp; Download`;
}

async function resizeImageIfNeeded(fileBlob, filename = null) {
  // Read target size from UI input (in MB), default to 1MB
  const targetInput = document.getElementById("rs-target-size");
  const targetMB = targetInput ? parseFloat(targetInput.value) || 1 : 1;
  const MAX_SIZE = Math.round(targetMB * 1024 * 1024);

  const fname = filename || rsUploadedFile.name;
  const ext = fname.split(".").pop().toLowerCase();

  if (fileBlob.size <= MAX_SIZE) {
    return { blob: fileBlob, filename: fname };
  }

  // Load image
  const img = await new Promise((res, rej) => {
    const url = URL.createObjectURL(fileBlob);
    const i = new Image();
    i.onload = () => {
      URL.revokeObjectURL(url);
      res(i);
    };
    i.onerror = () => rej(new Error("Gagal memuat gambar: " + fname));
    i.src = url;
  });

  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");

  const origWidth = img.naturalWidth;
  const origHeight = img.naturalHeight;
  let width = origWidth;
  let height = origHeight;
  let quality = 0.92;
  let currentBlob = fileBlob;
  let mimeType = "image/jpeg";
  if (ext === "png") mimeType = "image/png";
  else if (ext === "webp") mimeType = "image/webp";
  else mimeType = "image/jpeg";

  const isPng = mimeType === "image/png";

  // For PNG: quality parameter is ignored by canvas.toBlob, so we MUST reduce dimensions.
  // Strategy: estimate a good initial scale based on file size ratio, then iterate.

  if (isPng) {
    // PNG compression: iteratively scale down dimensions
    // Start with an estimated scale factor based on size ratio
    // PNG file size is roughly proportional to pixel count (width * height)
    let scaleFactor = Math.sqrt(MAX_SIZE / fileBlob.size) * 0.9; // slightly aggressive
    scaleFactor = Math.min(scaleFactor, 1);

    let iterations = 0;
    const MAX_ITERATIONS = 25;

    while (currentBlob.size > MAX_SIZE && iterations < MAX_ITERATIONS) {
      iterations++;

      width = Math.max(Math.floor(origWidth * scaleFactor), 16);
      height = Math.max(Math.floor(origHeight * scaleFactor), 16);

      canvas.width = width;
      canvas.height = height;
      ctx.clearRect(0, 0, width, height);
      ctx.drawImage(img, 0, 0, width, height);

      currentBlob = await new Promise((r) => canvas.toBlob(r, "image/png"));
      if (!currentBlob) break;

      if (currentBlob.size > MAX_SIZE) {
        // Calculate how much more we need to shrink
        const sizeRatio = MAX_SIZE / currentBlob.size;
        // Adjust scale factor proportionally (PNG size ~ pixel count)
        scaleFactor *= Math.sqrt(sizeRatio) * 0.95;
        scaleFactor = Math.max(scaleFactor, 0.01); // minimum 1% of original
      }
    }
  } else {
    // JPEG/WebP: first try reducing quality, then reduce dimensions
    let iterations = 0;
    const MAX_ITERATIONS = 25;

    while (currentBlob.size > MAX_SIZE && iterations < MAX_ITERATIONS) {
      iterations++;

      if (quality > 0.3) {
        // Reduce quality more aggressively based on size ratio
        const sizeRatio = MAX_SIZE / currentBlob.size;
        if (sizeRatio < 0.5) {
          quality -= 0.15;
        } else {
          quality -= 0.08;
        }
        quality = Math.max(quality, 0.3);
      } else {
        // Quality is already low, reduce dimensions
        const sizeRatio = MAX_SIZE / currentBlob.size;
        const dimScale = Math.sqrt(sizeRatio) * 0.95;
        width = Math.max(Math.floor(width * dimScale), 16);
        height = Math.max(Math.floor(height * dimScale), 16);
      }

      canvas.width = width;
      canvas.height = height;
      ctx.clearRect(0, 0, width, height);
      // Fill white background for JPEG (no alpha)
      if (mimeType === "image/jpeg") {
        ctx.fillStyle = "#FFFFFF";
        ctx.fillRect(0, 0, width, height);
      }
      ctx.drawImage(img, 0, 0, width, height);

      currentBlob = await new Promise((r) =>
        canvas.toBlob(r, mimeType, quality),
      );
      if (!currentBlob) break;
    }
  }

  return { blob: currentBlob, filename: fname };
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

updateRsUI();
