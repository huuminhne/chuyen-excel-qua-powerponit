'use strict';

const DEFAULT_SAMPLE = {
  viTitle: 'THÔNG BÁO TỶ GIÁ MUA, BÁN NGOẠI TỆ NGÀY 28/4/2026 – LẦN 01',
  enTitle: '(Foreign Exchange rates date 28/4/2026 – Count 01)',
  rows: [
    { stt: '1', code: 'EUR', buyCash: '30,296', buyTransfer: '30,418', sellCash: '31,600', sellTransfer: '31,600' },
    { stt: '2', code: 'GBP', buyCash: '-', buyTransfer: '35,155', sellCash: '-', sellTransfer: '36,165' },
    { stt: '3', code: 'HKD', buyCash: '-', buyTransfer: '3,305', sellCash: '-', sellTransfer: '3,420' },
    { stt: '4', code: 'USD', buyCash: '26,116', buyTransfer: '26,146', sellCash: '26,366', sellTransfer: '26,366' },
    { stt: '5', code: 'CHF', buyCash: '-', buyTransfer: '33,061', sellCash: '-', sellTransfer: '33,996' },
    { stt: '6', code: 'JPY', buyCash: '-', buyTransfer: '162', sellCash: '-', sellTransfer: '169' },
    { stt: '7', code: 'AUD', buyCash: '-', buyTransfer: '18,660', sellCash: '-', sellTransfer: '19,256' },
    { stt: '8', code: 'SGD', buyCash: '-', buyTransfer: '20,395', sellCash: '-', sellTransfer: '20,979' },
    { stt: '9', code: 'THB', buyCash: '-', buyTransfer: '796', sellCash: '-', sellTransfer: '831' },
    { stt: '10', code: 'SEK', buyCash: '-', buyTransfer: '2,804', sellCash: '-', sellTransfer: '2,911' },
    { stt: '11', code: 'NOK', buyCash: '-', buyTransfer: '2,777', sellCash: '-', sellTransfer: '2,893' },
    { stt: '12', code: 'DKK', buyCash: '-', buyTransfer: '4,062', sellCash: '-', sellTransfer: '4,203' },
    { stt: '13', code: 'CAD', buyCash: '-', buyTransfer: '19,053', sellCash: '-', sellTransfer: '19,632' },
    { stt: '14', code: 'NZD', buyCash: '-', buyTransfer: '15,323', sellCash: '-', sellTransfer: '15,857' }
  ]
};

const SLIDE_SIZE = { w: 1600, h: 900 };
const PPT_W = 10;
const PPT_H = 5.625;
const COLORS = {
  border: 'D35A5A',
  fillLight: 'FAF2F2',
  fillDark: 'EEDDDD',
  text: '2A2A2A',
  title: 'FFF8EA'
};
const LAYOUT = {
  title: { x: 800, viY: 118, enY: 153, viSize: 34, enSize: 25, width: 1220 },
  table: {
    x: 210,
    y: 184,
    bottom: 745,
    cols: [96, 235, 230, 180, 220, 220],
    headerHeights: [29, 29, 40, 54]
  }
};

const els = {
  fileInput: document.getElementById('fileInput'),
  dropzone: document.getElementById('dropzone'),
  fileInfo: document.getElementById('fileInfo'),
  parseStatus: document.getElementById('parseStatus'),
  viTitle: document.getElementById('viTitle'),
  enTitle: document.getElementById('enTitle'),
  loadSampleBtn: document.getElementById('loadSampleBtn'),
  resetBtn: document.getElementById('resetBtn'),
  addRowBtn: document.getElementById('addRowBtn'),
  exportPptxBtn: document.getElementById('exportPptxBtn'),
  exportVideoBtn: document.getElementById('exportVideoBtn'),
  musicToggle: document.getElementById('musicToggle'),
  videoDuration: document.getElementById('videoDuration'),
  previewCanvas: document.getElementById('previewCanvas'),
  exportCanvas: document.getElementById('exportCanvas'),
  ratesBody: document.getElementById('ratesBody'),
  progressWrap: document.getElementById('progressWrap'),
  progressBar: document.getElementById('progressBar'),
  progressText: document.getElementById('progressText'),
  downloadLink: document.getElementById('downloadLink'),
  videoPreviewWrap: document.getElementById('videoPreviewWrap'),
  videoPreview: document.getElementById('videoPreview')
};

const state = {
  viTitle: DEFAULT_SAMPLE.viTitle,
  enTitle: DEFAULT_SAMPLE.enTitle,
  rows: cloneRows(DEFAULT_SAMPLE.rows)
};

let bgImage = null;
let bgImagePromise = null;
let musicLoadPromise = null;
let musicDurationPromise = null;

bootstrap();

function bootstrap() {
  bindEvents();
  loadBgImage().then(() => {
    renderAll();
    setStatus('Sẵn sàng. Bạn có thể nạp file Excel hoặc dùng mẫu có sẵn.', false);
  }).catch(err => {
    console.error(err);
    renderAll();
    setStatus('Đã sẵn sàng, nhưng không tải được ảnh nền mẫu. Vẫn có thể xuất file.', true);
  });
}

function bindEvents() {
  els.fileInput.addEventListener('change', handleFileSelected);
  els.dropzone.addEventListener('click', e => {
    if (e.target.tagName !== 'INPUT') els.fileInput.click();
  });
  els.dropzone.addEventListener('dragover', e => {
    e.preventDefault();
    els.dropzone.classList.add('drag');
  });
  els.dropzone.addEventListener('dragleave', () => els.dropzone.classList.remove('drag'));
  els.dropzone.addEventListener('drop', e => {
    e.preventDefault();
    els.dropzone.classList.remove('drag');
    const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) readWorkbook(file);
  });

  els.loadSampleBtn.addEventListener('click', () => {
    applySample();
    setStatus('Đã khôi phục dữ liệu mẫu.', false);
  });

  els.resetBtn.addEventListener('click', () => {
    applySample();
    setStatus('Đã khôi phục dữ liệu mẫu.', false);
  });

  els.viTitle.addEventListener('input', () => {
    state.viTitle = els.viTitle.value;
    requestRender();
  });
  els.enTitle.addEventListener('input', () => {
    state.enTitle = els.enTitle.value;
    requestRender();
  });

  els.addRowBtn.addEventListener('click', () => {
    state.rows.push({
      stt: String(state.rows.length + 1),
      code: '',
      buyCash: '-',
      buyTransfer: '-',
      sellCash: '-',
      sellTransfer: '-'
    });
    renderTable();
    requestRender();
  });

  els.exportPptxBtn.addEventListener('click', exportPptx);
  els.exportVideoBtn.addEventListener('click', exportVideo);
}

function applySample() {
  state.viTitle = DEFAULT_SAMPLE.viTitle;
  state.enTitle = DEFAULT_SAMPLE.enTitle;
  state.rows = cloneRows(DEFAULT_SAMPLE.rows);
  syncInputs();
  renderTable();
  requestRender();
}

function cloneRows(rows) {
  return rows.map(row => ({ ...row }));
}

function syncInputs() {
  els.viTitle.value = state.viTitle;
  els.enTitle.value = state.enTitle;
}

function setStatus(message, isError = false) {
  els.parseStatus.textContent = message;
  els.parseStatus.style.color = isError ? '#99152B' : '#6B7280';
}

function handleFileSelected(e) {
  const file = e.target.files && e.target.files[0];
  if (file) readWorkbook(file);
}

async function readWorkbook(file) {
  els.fileInfo.textContent = `Đã chọn: ${file.name}`;
  setStatus('Đang đọc file Excel…', false);

  try {
    const arrayBuffer = await file.arrayBuffer();
    const wb = XLSX.read(arrayBuffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });

    const titles = extractTitles(rows);
    const rates = extractRates(rows);

    state.viTitle = titles.vi || DEFAULT_SAMPLE.viTitle;
    state.enTitle = titles.en || DEFAULT_SAMPLE.enTitle;
    state.rows = rates.length ? rates : cloneRows(DEFAULT_SAMPLE.rows);

    syncInputs();
    renderTable();
    requestRender();

    if (rates.length) {
      setStatus(`Đã nhận ${rates.length} dòng tỷ giá từ file. Bạn có thể xuất ngay hoặc chỉnh thêm.`, false);
    } else {
      setStatus('Không tự nhận diện được bảng tỷ giá trong file. Đã nạp dữ liệu mẫu để bạn chỉnh.', true);
    }
  } catch (err) {
    console.error(err);
    setStatus('Không đọc được file này. Hãy kiểm tra lại định dạng Excel.', true);
    els.fileInfo.textContent = '';
  }
}

function extractTitles(rows) {
  let vi = '';
  let en = '';

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    for (const cell of row) {
      if (typeof cell === 'string' && /THÔNG BÁO/i.test(cell)) {
        vi = normalizeText(cell);
        for (let k = i; k < Math.min(rows.length, i + 6); k++) {
          for (const c of (rows[k] || [])) {
            if (typeof c === 'string' && /^\s*\(/.test(c)) {
              en = normalizeText(c);
              break;
            }
          }
          if (en) break;
        }
        return { vi, en };
      }
    }
  }
  return { vi, en };
}

function extractRates(rows) {
  const out = [];
  for (const row of rows) {
    const stt = row[0];
    const code = row[1];
    const sttStr = typeof stt === 'number' ? String(stt) : (typeof stt === 'string' ? stt.trim() : '');
    const codeStr = typeof code === 'string' ? code.trim().toUpperCase() : '';
    if (!/^\d{1,2}$/.test(sttStr) || !/^[A-Z]{2,4}$/.test(codeStr)) continue;

    out.push({
      stt: sttStr,
      code: codeStr,
      buyCash: formatRate(row[2]),
      buyTransfer: formatRate(row[3]),
      sellCash: formatRate(row[4]),
      sellTransfer: formatRate(row[5])
    });
  }
  return out;
}

function formatRate(v) {
  if (v === undefined || v === null || v === '') return '-';
  if (typeof v === 'number') return Math.round(v).toLocaleString('en-US');
  const t = String(v).trim();
  if (!t || t === '-' || t === '−') return '-';
  const n = Number(t.replace(/,/g, ''));
  if (Number.isFinite(n)) return Math.round(n).toLocaleString('en-US');
  return t;
}

function normalizeText(text) {
  return String(text).replace(/\s+/g, ' ').trim();
}

function renderTable() {
  els.ratesBody.innerHTML = '';
  state.rows.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.dataset.index = String(index);

    tr.appendChild(makeCell(row.stt, 'stt', false));
    tr.appendChild(makeCell(row.code, 'code', true));
    tr.appendChild(makeCell(row.buyCash, 'buyCash', false));
    tr.appendChild(makeCell(row.buyTransfer, 'buyTransfer', false));
    tr.appendChild(makeCell(row.sellCash, 'sellCash', false));
    tr.appendChild(makeCell(row.sellTransfer, 'sellTransfer', false));

    const action = document.createElement('td');
    action.className = 'row-action';
    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'ghost';
    del.textContent = 'Xoá';
    del.addEventListener('click', () => {
      state.rows.splice(index, 1);
      renderTable();
      requestRender();
    });
    action.appendChild(del);
    tr.appendChild(action);

    els.ratesBody.appendChild(tr);
  });
}

function makeCell(value, key, codeClass) {
  const td = document.createElement('td');
  td.contentEditable = 'true';
  td.textContent = value;
  if (codeClass) td.classList.add('code');

  td.addEventListener('input', () => {
    const row = td.parentElement;
    const idx = Number(row.dataset.index);
    state.rows[idx][key] = normalizeText(td.textContent);
    requestRender();
  });

  td.addEventListener('blur', () => {
    if (key === 'code') td.textContent = td.textContent.trim().toUpperCase();
    if (!td.textContent.trim() && key !== 'stt' && key !== 'code') td.textContent = '-';
    state.rows[Number(td.parentElement.dataset.index)][key] = normalizeText(td.textContent);
    requestRender();
  });

  return td;
}

function requestRender() {
  syncInputs();
  drawToCanvas(els.previewCanvas, 1);
}

function renderAll() {
  syncInputs();
  renderTable();
  drawToCanvas(els.previewCanvas, 1);
}

function loadBgImage() {
  if (bgImagePromise) return bgImagePromise;
  bgImagePromise = new Promise((resolve, reject) => {
    if (typeof SLIDE_BG_DATA !== 'string') {
      reject(new Error('Missing SLIDE_BG_DATA'));
      return;
    }
    const img = new Image();
    img.onload = () => {
      bgImage = img;
      resolve(img);
    };
    img.onerror = reject;
    img.src = SLIDE_BG_DATA;
  });
  return bgImagePromise;
}

function drawToCanvas(canvas, scale = 1) {
  const w = Math.round(SLIDE_SIZE.w * scale);
  const h = Math.round(SLIDE_SIZE.h * scale);
  if (canvas.width !== w) canvas.width = w;
  if (canvas.height !== h) canvas.height = h;

  const ctx = canvas.getContext('2d');
  ctx.save();
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  ctx.clearRect(0, 0, SLIDE_SIZE.w, SLIDE_SIZE.h);

  if (bgImage) {
    ctx.drawImage(bgImage, 0, 0, SLIDE_SIZE.w, SLIDE_SIZE.h);
  } else {
    const grad = ctx.createLinearGradient(0, 0, 0, SLIDE_SIZE.h);
    grad.addColorStop(0, '#99152B');
    grad.addColorStop(0.54, '#F4D9B0');
    grad.addColorStop(1, '#B31C42');
    ctx.fillStyle = grad;
    ctx.fillRect(0, 0, SLIDE_SIZE.w, SLIDE_SIZE.h);
  }

  drawTitles(ctx);
  drawTable(ctx);

  ctx.restore();
}

function drawTitles(ctx) {
  const title = state.viTitle || DEFAULT_SAMPLE.viTitle;
  const subtitle = state.enTitle || DEFAULT_SAMPLE.enTitle;

  ctx.fillStyle = `#${COLORS.title}`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';

  ctx.font = `400 ${LAYOUT.title.viSize}px "Times New Roman", serif`;
  wrapCentered(ctx, title, LAYOUT.title.x, LAYOUT.title.viY, 38, LAYOUT.title.width);

  ctx.font = `400 ${LAYOUT.title.enSize}px "Times New Roman", serif`;
  wrapCentered(ctx, subtitle, LAYOUT.title.x, LAYOUT.title.enY, 30, LAYOUT.title.width);
}

function wrapCentered(ctx, text, x, y, lineHeight, maxWidth = 1040) {
  const words = String(text).split(' ');
  const lines = [];
  let current = '';

  words.forEach(word => {
    const test = current ? `${current} ${word}` : word;
    if (ctx.measureText(test).width > maxWidth && current) {
      lines.push(current);
      current = word;
    } else {
      current = test;
    }
  });
  if (current) lines.push(current);

  lines.forEach((line, i) => ctx.fillText(line, x, y + i * lineHeight));
}

function drawTable(ctx) {
  const tableX = LAYOUT.table.x;
  const tableY = LAYOUT.table.y;
  const cols = LAYOUT.table.cols;
  const rows = buildTableLayout();

  const drawRect = (x, y, w, h, fill, stroke) => {
    ctx.fillStyle = fill;
    ctx.fillRect(x, y, w, h);
    ctx.strokeStyle = stroke;
    ctx.lineWidth = 1;
    ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
  };

  const drawText = (txt, x, y, w, h, fontSize, bold, fill) => {
    ctx.fillStyle = fill;
    ctx.font = `${bold ? 700 : 400} ${fontSize}px "Times New Roman", serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    const lines = String(txt).split('\n');
    const lineH = fontSize * 1.04;
    const startY = y + h / 2 - ((lines.length - 1) * lineH) / 2;
    lines.forEach((line, i) => ctx.fillText(line, x + w / 2, startY + i * lineH, w - 8));
  };

  let y = tableY;
  const headerSpans = [
    [
      { text: 'STT\n/No.', col: 0, span: 1, rowSpan: 4 },
      { text: 'Ngoại tệ\n(Foreign currency)', col: 1, span: 1, rowSpan: 4 },
      { text: 'Tỷ giá mua', col: 2, span: 2, rowSpan: 1 },
      { text: 'Tỷ giá bán', col: 4, span: 2, rowSpan: 1 }
    ],
    [
      { text: '(Buying rate)', col: 2, span: 2, rowSpan: 1 },
      { text: '(Selling rate)', col: 4, span: 2, rowSpan: 1 }
    ],
    [
      { text: 'Tiền mặt', col: 2, span: 1, rowSpan: 1 },
      { text: 'Chuyển khoản', col: 3, span: 1, rowSpan: 1 },
      { text: 'Tiền mặt', col: 4, span: 1, rowSpan: 1 },
      { text: 'Chuyển khoản', col: 5, span: 1, rowSpan: 1 }
    ],
    [
      { text: '(Cash)', col: 2, span: 1, rowSpan: 1 },
      { text: '(Transferring)', col: 3, span: 1, rowSpan: 1 },
      { text: '(Cash)', col: 4, span: 1, rowSpan: 1 },
      { text: '(Transferring)', col: 5, span: 1, rowSpan: 1 }
    ]
  ];

  const headerHeights = rows.headerHeights;
  for (let r = 0; r < 4; r++) {
    let x = tableX;
    const h = headerHeights[r];
    let col = 0;
    for (const head of headerSpans[r]) {
      while (col < head.col) {
        x += cols[col];
        col += 1;
      }
      const w = cols.slice(head.col, head.col + head.span).reduce((a, b) => a + b, 0);
      const hh = head.rowSpan === 4 ? headerHeights.reduce((a, b) => a + b, 0) : h;
      drawRect(x, y, w, hh, r < 2 ? `#${COLORS.fillLight}` : `#${COLORS.fillDark}`, `#${COLORS.border}`);
      drawText(head.text, x, y, w, hh, r === 0 ? 18 : 17, true, `#${COLORS.text}`);
      x += w;
      col = head.col + head.span;
    }
    y += h;
  }

  state.rows.forEach((row, i) => {
    const fill = i % 2 === 0 ? `#${COLORS.fillLight}` : `#${COLORS.fillDark}`;
    const rh = rows.rowHeights[i];
    let x = tableX;
    const vals = [row.stt, row.code, row.buyCash, row.buyTransfer, row.sellCash, row.sellTransfer];
    vals.forEach((val, colIdx) => {
      drawRect(x, y, cols[colIdx], rh, fill, `#${COLORS.border}`);
      drawText(val || '-', x, y, cols[colIdx], rh, colIdx === 1 ? 17 : 16, colIdx === 1, `#${COLORS.text}`);
      x += cols[colIdx];
    });
    y += rh;
  });
}

function buildTableLayout() {
  const headerHeights = LAYOUT.table.headerHeights;
  const totalHeader = headerHeights.reduce((a, b) => a + b, 0);
  const available = Math.max(240, LAYOUT.table.bottom - LAYOUT.table.y - totalHeader);
  const rows = Math.max(1, state.rows.length || 1);
  const rowH = Math.max(22, Math.min(29, Math.round(available / rows)));
  return {
    headerHeights,
    rowHeights: Array.from({ length: rows }, () => rowH)
  };
}

function drawFooter(ctx) {
  // Footer/logo live inside the background image to match the reference slide exactly.
}


function drawAnimatedFrame(canvas, progress) {
  // Keep the exported video visually identical to the slide: one clean static frame.
  drawToCanvas(canvas, 1);
}

async function exportPptx() {
  try {
    setBusy(true, 'Đang tạo PowerPoint…');
    const png = await canvasToDataURL(els.exportCanvas, 2);
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';

    const slide = pptx.addSlide();
    slide.addImage({ data: png, x: 0, y: 0, w: PPT_W, h: PPT_H });

    const fname = buildBaseFileName() + '.pptx';
    await pptx.writeFile({ fileName: fname });

    setBusy(false);
    setDownloadMessage(`Đã tạo xong ${fname}`, null);
  } catch (err) {
    console.error(err);
    setBusy(false);
    setDownloadMessage('Không tạo được PowerPoint. Hãy thử lại trên Chrome hoặc Edge.', true);
  }
}

async function exportVideo() {
  if (!window.MediaRecorder || !HTMLCanvasElement.prototype.captureStream) {
    setDownloadMessage('Trình duyệt này chưa hỗ trợ ghi video từ canvas.', true);
    return;
  }

  try {
    setBusy(true, 'Đang tạo video…');
    const canvas = els.exportCanvas;

    let durationMs = Number(els.videoDuration.value) || 30000;
    if (els.musicToggle.checked) {
      try {
        await ensureMusicLoaded();
        durationMs = Math.max(1000, await getMusicDurationMs());
      } catch (musicErr) {
        console.warn('Music duration fallback:', musicErr);
      }
    }

    // Audio in exported videos is more reliable when we wait for the track to be live
    // before starting the recorder.
    const stream = canvas.captureStream(24);
    let combinedStream = stream;
    let audio = null;
    let audioCtx = null;

    if (els.musicToggle.checked) {
      try {
        const music = typeof AGRIBANK_MUSIC_DATA === 'string' ? AGRIBANK_MUSIC_DATA : '';
        if (music) {
          audio = new Audio(music);
          audio.preload = 'auto';
          audio.loop = false;
          audio.volume = 0.95;
          audio.muted = false;
          audio.crossOrigin = 'anonymous';
          const audioPack = await captureAudioStream(audio);
          if (audioPack.stream) {
            combinedStream = new MediaStream([
              ...stream.getVideoTracks(),
              ...audioPack.stream.getAudioTracks()
            ]);
            audioCtx = audioPack.audioCtx;
          }
        }
      } catch (musicErr) {
        console.warn('Music capture skipped:', musicErr);
      }
    }

    drawAnimatedFrame(canvas, 0);

    // Prefer a WebM encoder when audio is enabled; it is the most stable path in browsers.
    const mimeType = els.musicToggle.checked ? pickWebmMimeType() : pickMimeType();
    const options = mimeType ? { mimeType, videoBitsPerSecond: 1000000, audioBitsPerSecond: 128000 } : { videoBitsPerSecond: 1000000, audioBitsPerSecond: 128000 };
    const recorder = new MediaRecorder(combinedStream, options);
    const chunks = [];

    recorder.ondataavailable = e => {
      if (e.data && e.data.size) chunks.push(e.data);
    };

    const cleanupAudio = () => {
      try {
        if (audio) {
          audio.pause();
          audio.currentTime = 0;
        }
        if (audioCtx && audioCtx.close) audioCtx.close().catch(() => {});
      } catch (_) {}
    };

    const blobPromise = new Promise((resolve, reject) => {
      recorder.onstop = () => {
        cleanupAudio();
        resolve(new Blob(chunks, { type: recorder.mimeType || 'video/webm' }));
      };
      recorder.onerror = err => {
        cleanupAudio();
        reject(err);
      };
    });

    if (audio) {
      try {
        await audio.play();
        await waitForAudioStart(audio);
      } catch (err) {
        console.warn(err);
      }
    }

    recorder.start(250);

    const start = performance.now();
    let stopped = false;
    els.progressWrap.hidden = false;
    els.progressBar.style.width = '0%';

    const tick = (now) => {
      if (stopped) return;
      const elapsed = now - start;
      const progress = Math.min(1, elapsed / durationMs);

      drawAnimatedFrame(canvas, progress);

      els.progressBar.style.width = `${Math.round(progress * 100)}%`;
      els.progressText.textContent = `Đang ghi video: ${formatTime(elapsed)} / ${formatTime(durationMs)}${els.musicToggle.checked ? ' (đã ghép nhạc nền)' : ''}`;

      if (elapsed >= durationMs) {
        stopped = true;
        if (recorder.state !== 'inactive') recorder.stop();
        return;
      }
      requestAnimationFrame(tick);
    };

    requestAnimationFrame(tick);

    const blob = await blobPromise;
    const url = URL.createObjectURL(blob);
    const ext = /webm/i.test(recorder.mimeType || '') ? 'webm' : (/mp4/i.test(recorder.mimeType || '') ? 'mp4' : 'webm');
    const fname = buildBaseFileName() + '.' + ext;

    els.downloadLink.hidden = false;
    els.downloadLink.href = url;
    els.downloadLink.download = fname;
    els.downloadLink.textContent = `Tải ${fname}`;

    if (els.videoPreview && els.videoPreviewWrap) {
      els.videoPreview.src = url;
      els.videoPreview.load();
      els.videoPreviewWrap.hidden = false;
    }

    setBusy(false);
    els.progressText.textContent = `Đã tạo xong ${fname}${els.musicToggle.checked ? ' (video có âm thanh)' : ''}`;
  } catch (err) {
    console.error(err);
    setBusy(false);
    setDownloadMessage('Không tạo được video. Hãy thử lại trên Chrome hoặc Edge.', true);
  }
}

function pickMimeType() {
  const list = [
    'video/mp4;codecs=avc1.42E01E,mp4a.40.2',
    'video/mp4;codecs=avc1.42E01E',
    'video/mp4',
    'video/webm;codecs=vp8,opus',
    'video/webm;codecs=vp8',
    'video/webm'
  ];
  for (const type of list) {
    if (window.MediaRecorder && MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(type)) {
      return type;
    }
  }
  return '';
}

function pickWebmMimeType() {
  const list = [
    'video/webm;codecs=vp9,opus',
    'video/webm;codecs=vp8,opus',
    'video/webm;codecs=vp8',
    'video/webm'
  ];
  for (const type of list) {
    if (window.MediaRecorder && MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(type)) {
      return type;
    }
  }
  return '';
}

function waitForAudioStart(audio) {
  return new Promise(resolve => {
    if (!audio) {
      resolve();
      return;
    }
    if (!audio.paused && audio.currentTime > 0) {
      resolve();
      return;
    }
    const done = () => {
      audio.removeEventListener('playing', done);
      audio.removeEventListener('timeupdate', done);
      audio.removeEventListener('canplay', done);
      resolve();
    };
    audio.addEventListener('playing', done, { once: true });
    audio.addEventListener('timeupdate', done, { once: true });
    audio.addEventListener('canplay', done, { once: true });
    setTimeout(done, 500);
  });
}

async function ensureMusicLoaded() {
  if (window.AGRIBANK_MUSIC_DATA) return;
  if (musicLoadPromise) return musicLoadPromise;

  musicLoadPromise = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'music-data.js';
    script.onload = () => resolve();
    script.onerror = reject;
    document.head.appendChild(script);
  });

  return musicLoadPromise;
}

async function getMusicDurationMs() {
  if (musicDurationPromise) return musicDurationPromise;
  musicDurationPromise = new Promise((resolve, reject) => {
    const src = typeof AGRIBANK_MUSIC_DATA === 'string' ? AGRIBANK_MUSIC_DATA : '';
    if (!src) {
      reject(new Error('Missing music source'));
      return;
    }
    const audio = new Audio(src);
    audio.preload = 'metadata';
    audio.onloadedmetadata = () => {
      const ms = Number.isFinite(audio.duration) ? Math.round(audio.duration * 1000) : NaN;
      if (Number.isFinite(ms) && ms > 0) resolve(ms);
      else reject(new Error('Invalid audio duration'));
    };
    audio.onerror = () => reject(new Error('Failed to load audio metadata'));
  });
  return musicDurationPromise;
}

async function captureAudioStream(audio) {
  if (audio.captureStream) {
    return { stream: audio.captureStream(), audioCtx: null };
  }
  if (audio.mozCaptureStream) {
    return { stream: audio.mozCaptureStream(), audioCtx: null };
  }

  const AudioCtx = window.AudioContext || window.webkitAudioContext;
  if (!AudioCtx) throw new Error('AudioContext not supported');

  const audioCtx = new AudioCtx();
  const src = audioCtx.createMediaElementSource(audio);
  const dest = audioCtx.createMediaStreamDestination();
  src.connect(dest);
  src.connect(audioCtx.destination);
  if (audioCtx.state === 'suspended' && audioCtx.resume) {
    await audioCtx.resume();
  }
  return { stream: dest.stream, audioCtx };
}

function canvasToDataURL(canvas, scale) {
  return new Promise(resolve => {
    drawToCanvas(canvas, scale);
    resolve(canvas.toDataURL('image/png'));
  });
}

function buildBaseFileName() {
  const dateMatch = String(state.viTitle).match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!dateMatch) return 'ThongBaoTyGia';
  return `ThongBaoTyGia_${dateMatch[1]}${dateMatch[2]}${dateMatch[3]}`;
}

function formatTime(ms) {
  const s = Math.max(0, Math.floor(ms / 1000));
  const m = Math.floor(s / 60);
  return `${m}:${String(s % 60).padStart(2, '0')}`;
}

function setBusy(isBusy, text) {
  els.exportPptxBtn.disabled = isBusy;
  els.exportVideoBtn.disabled = isBusy;
  els.loadSampleBtn.disabled = isBusy;
  els.resetBtn.disabled = isBusy;
  els.addRowBtn.disabled = isBusy;
  els.progressWrap.hidden = !isBusy;
  if (text) els.progressText.textContent = text;
  els.progressBar.style.width = isBusy ? '0%' : '100%';
}

function setDownloadMessage(msg, isError) {
  els.downloadLink.hidden = true;
  els.progressWrap.hidden = false;
  els.progressText.textContent = msg;
  els.progressText.style.color = isError ? '#99152B' : '#6B7280';
}
