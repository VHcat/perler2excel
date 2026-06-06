// =====================================================
// app.js - 拼豆图转 Excel 像素画 v2.0 主逻辑
// =====================================================

// --- 全局状态 ---
let cropper = null;
let currentGridData = null;   // 缓存网格数据 [{code, rgb, isEmpty}]
let currentCounts = {};
let originalFileName = '';
let previewStale = false;     // 参数变更时标记为 true，导出时检查

// DOM 引用
const $ = id => document.getElementById(id);
const imageEl       = $('image');
const previewCanvas = $('previewCanvas');
const previewCtx    = previewCanvas.getContext('2d');
const statusText    = $('statusText');
const progressBar   = $('progressBar');
const progressFill  = $('progressFill');
const statsPanel    = $('statsPanel');
const tooltip       = $('previewTooltip');

// --- 色卡管理 (Task 4) ---
const PALETTES = {
  'artkal-s': PALETTE_ARTKAL_S,
  'mard':     PALETTE_MARD,
};

function getCurrentPalette() {
  return PALETTES[$('paletteSelect').value] || PALETTE_MARD;
}

// 切换色卡后自动重新预计算 LAB
$('paletteSelect').addEventListener('change', () => {
  precomputeLAB(getCurrentPalette());
  if (currentGridData) {
    setStatus(`已切换至 ${$('paletteSelect').selectedOptions[0].text} 色卡，正在重新匹配...`);
    generatePreview();
  } else {
    setStatus(`已切换至 ${$('paletteSelect').selectedOptions[0].text} 色卡，共 ${getCurrentPalette().length} 色。请上传图片并生成预览。`);
  }
});

// --- LAB 色彩空间算法 ---
function rgbToLab(rgb) {
  let r = rgb[0]/255, g = rgb[1]/255, b = rgb[2]/255;
  r = r > 0.04045 ? Math.pow((r+0.055)/1.055, 2.4) : r/12.92;
  g = g > 0.04045 ? Math.pow((g+0.055)/1.055, 2.4) : g/12.92;
  b = b > 0.04045 ? Math.pow((b+0.055)/1.055, 2.4) : b/12.92;
  let x = (r*0.4124 + g*0.3576 + b*0.1805) / 0.95047;
  let y = (r*0.2126 + g*0.7152 + b*0.0722) / 1.00000;
  let z = (r*0.0193 + g*0.1192 + b*0.9505) / 1.08883;
  x = x > 0.008856 ? Math.cbrt(x) : (7.787*x) + 16/116;
  y = y > 0.008856 ? Math.cbrt(y) : (7.787*y) + 16/116;
  z = z > 0.008856 ? Math.cbrt(z) : (7.787*z) + 16/116;
  return [(116*y)-16, 500*(x-y), 200*(y-z)];
}

function deltaE(a, b) {
  const dL = a[0]-b[0], da = a[1]-b[1], db = a[2]-b[2];
  return Math.sqrt(dL*dL + da*da + db*db);
}

// 预计算色卡 LAB 缓存
let paletteLABCache = [];
function precomputeLAB(palette) {
  paletteLABCache = palette.map(c => ({ ...c, lab: rgbToLab(c.rgb) }));
}
precomputeLAB(getCurrentPalette());

function nearestColor(rgb) {
  const target = rgbToLab(rgb);
  let best = paletteLABCache[0], minD = Infinity;
  for (const c of paletteLABCache) {
    const d = deltaE(target, c.lab);
    if (d < minD) { minD = d; best = c; }
  }
  return best;
}

// --- 工具函数 ---
function setStatus(msg) { statusText.textContent = msg; }
function showProgress(pct) {
  progressBar.classList.add('active');
  progressFill.style.width = pct + '%';
}
function hideProgress() {
  progressBar.classList.remove('active');
  progressFill.style.width = '0%';
}

// Task 2: 参数钳位
function clampInt(id, min, max) {
  const el = $(id);
  let v = parseInt(el.value) || min;
  v = Math.max(min, Math.min(max, v));
  el.value = v;
  return v;
}

// --- Task 8: 引导栏 ---
(function initGuide() {
  if (localStorage.getItem('guide_closed') === '1') {
    $('guideBar').classList.add('hidden');
  }
  $('closeGuide').addEventListener('click', () => {
    $('guideBar').classList.add('hidden');
    localStorage.setItem('guide_closed', '1');
  });
})();

// --- 拖拽上传图片 ---
const paneBody = document.querySelector('.workspace .pane-body');
['dragenter', 'dragover'].forEach(evt => {
  paneBody.addEventListener(evt, e => { e.preventDefault(); paneBody.classList.add('drag-over'); });
});
['dragleave', 'drop'].forEach(evt => {
  paneBody.addEventListener(evt, e => { e.preventDefault(); paneBody.classList.remove('drag-over'); });
});
paneBody.addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  if (!file || !file.type.startsWith('image/')) return;
  // 触发与文件选择器相同的加载逻辑
  const dt = new DataTransfer();
  dt.items.add(file);
  $('imageInput').files = dt.files;
  $('imageInput').dispatchEvent(new Event('change'));
});

// --- 灵敏度滑块显隐（仅在勾选“忽略白色”时显示） ---
$('ignoreWhite').addEventListener('change', () => {
  $('thresholdLabel').style.display = $('ignoreWhite').checked ? '' : 'none';
});

// --- 参数变更标记过期 ---
['cols', 'rows', 'pixelSize', 'boardSize', 'paletteSelect', 'ignoreWhite', 'whiteThreshold', 'showCodes', 'colorOnly', 'includeLegend'].forEach(id => {
  $(id).addEventListener('change', () => { previewStale = true; });
});

// --- 快捷键 ---
document.addEventListener('keydown', e => {
  // 如果焦点在输入框中且不是 Enter，跳过
  if (e.target.tagName === 'INPUT' && e.key !== 'Enter') return;
  if (e.key === 'Enter') { e.preventDefault(); generatePreview(); }
  if (e.key === 'e' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); exportExcel(); }
});

// --- 编辑/浏览模式切换（手机端） ---
const isMobile = () => window.innerWidth <= 768;
let cropEditMode = !isMobile(); // 手机端默认浏览模式

function applyCropMode() {
  const workspace = document.querySelector('.workspace');
  const btn = $('modeToggle');
  if (cropEditMode) {
    workspace.classList.remove('browse-mode');
    btn.classList.remove('browse');
    $('modeIcon').textContent = '\u270E'; // ✎
    $('modeText').textContent = '编辑中';
    if (cropper) { cropper.setDragMode('crop'); updateGridOverlay(); }
  } else {
    $('gridOverlay').style.display = 'none';
    workspace.classList.add('browse-mode');
    btn.classList.add('browse');
    $('modeIcon').textContent = '\u270B'; // ✋
    $('modeText').textContent = '可滑动';
    if (cropper) cropper.setDragMode('none');
  }
}

$('modeToggle').addEventListener('click', () => {
  cropEditMode = !cropEditMode;
  applyCropMode();
});

// --- 拼豆网格叠加层 ---
function updateGridOverlay() {
  const overlay = $('gridOverlay');
  if (!cropper || !cropEditMode) {
    overlay.style.display = 'none';
    return;
  }

  const box = cropper.getCropBoxData();
  let cols, rows;

  if (calibratedCellW !== null) {
    // 校准模式：行列数 = 裁切框 / 单个格子
    cols = Math.max(1, Math.min(200, Math.round(box.width / calibratedCellW)));
    rows = Math.max(1, Math.min(200, Math.round(box.height / calibratedCellH)));
    $('cols').value = cols;
    $('rows').value = rows;
    if (document.activeElement !== $('cellPx')) {
      $('cellPx').value = Math.round((calibratedCellW + calibratedCellH) / 2);
    }
  } else {
    cols = parseInt($('cols').value) || 26;
    rows = parseInt($('rows').value) || 23;
    if (document.activeElement !== $('cellPx')) {
      $('cellPx').value = Math.round(box.width / cols);
    }
  }

  overlay.style.display = 'block';
  overlay.style.left   = box.left + 'px';
  overlay.style.top    = box.top + 'px';
  overlay.style.width  = box.width + 'px';
  overlay.style.height = box.height + 'px';
  overlay.style.setProperty('--grid-cw', (box.width / cols) + 'px');
  overlay.style.setProperty('--grid-ch', (box.height / rows) + 'px');
}

// --- 格子校准 ---
let calibratedCellW = null;
let calibratedCellH = null;

$('calibrateBtn').addEventListener('click', () => {
  if (!cropper) return;

  if (calibratedCellW !== null) {
    // 重新校准：清除旧值，恢复手动模式
    calibratedCellW = null;
    calibratedCellH = null;
    const btn = $('calibrateBtn');
    btn.textContent = '📐 设为一格';
    btn.classList.remove('calibrated');
    btn.title = '先框选一个格子 → 点此按钮 → 再框选识别范围';
    updateGridOverlay();
    return;
  }

  // 记录当前选框作为单个格子的尺寸
  const box = cropper.getCropBoxData();
  calibratedCellW = box.width;
  calibratedCellH = box.height;

  // 将裁切框扩展到整张图
  const cd = cropper.getCanvasData();
  cropper.setCropBoxData({ left: cd.left, top: cd.top, width: cd.width, height: cd.height });

  // 解锁比例，自由调整识别范围
  if (aspectLocked) $('lockAspect').click();

  $('cellPx').value = Math.round((box.width + box.height) / 2);

  const btn = $('calibrateBtn');
  btn.textContent = '✓ 已设定 ' + Math.round(box.width) + '×' + Math.round(box.height) + 'px';
  btn.classList.add('calibrated');
  btn.title = '点击重新设定格子大小';

  applyCellPx();
});

// --- 比例锁定（默认解锁，自由矩形） ---
let aspectLocked = false;

$('lockAspect').addEventListener('click', () => {
  aspectLocked = !aspectLocked;
  const btn = $('lockAspect');
  if (aspectLocked) {
    btn.textContent = '🔒'; // 🔒
    btn.classList.add('locked');
    btn.title = '🔒 锁定：列/行决定裁切框比例\n🔓 解锁：裁切框自由拖拽，像素/格独立调节';
    if (cropper) {
      const cols = parseInt($('cols').value) || 26;
      const rows = parseInt($('rows').value) || 23;
      cropper.setAspectRatio(cols / rows);
      updateGridOverlay();
    }
  } else {
    btn.textContent = '🔓'; // 🔓
    btn.classList.remove('locked');
    btn.title = '🔒 锁定：列/行决定裁切框比例\n🔓 解锁：裁切框自由拖拽，像素/格独立调节';
    // 解锁时也更新一次网格，让用户看到当前格子划分
    updateGridOverlay();
  }
});

// --- 像素/格 → 列/行联动 ---
function applyCellPx() {
  if (!cropper) return;
  const cellPx = Math.max(2, Math.min(500, parseFloat($('cellPx').value) || 2));

  if (calibratedCellW !== null) {
    // 校准模式：缩放已校准的格子尺寸
    const oldAvg = Math.round((calibratedCellW + calibratedCellH) / 2);
    if (oldAvg > 0) {
      const scale = cellPx / oldAvg;
      calibratedCellW = Math.max(1, Math.round(calibratedCellW * scale));
      calibratedCellH = Math.max(1, Math.round(calibratedCellH * scale));
    }
    updateGridOverlay();
    return;
  }

  // 手动模式
  const box = cropper.getCropBoxData();
  const cols = Math.max(1, Math.min(200, Math.round(box.width / cellPx)));
  const rows = Math.max(1, Math.min(200, Math.round(box.height / cellPx)));
  $('cols').value = cols;
  $('rows').value = rows;
  if (aspectLocked) cropper.setAspectRatio(cols / rows);
  updateGridOverlay();
}

$('cellPx').addEventListener('input', applyCellPx);

function adjustCellPx(delta) {
  if (!cropper) return;
  const cur = parseFloat($('cellPx').value) || Math.round(cropper.getCropBoxData().width / (parseInt($('cols').value) || 26));
  // 自适应步长：值小调 0.5，中大调 1~5，很大调 10~25
  const step = cur < 10 ? 0.5 : cur < 30 ? 1 : cur < 60 ? 5 : cur < 150 ? 10 : 25;
  $('cellPx').value = Math.max(2, Math.min(500, cur + delta * step));
  applyCellPx();
}

$('cellPxDown').addEventListener('click', () => adjustCellPx(-1));
$('cellPxUp').addEventListener('click', () => adjustCellPx(1));

// --- 下载弹窗关闭 ---
$('dlClose').addEventListener('click', () => {
  $('dlOverlay').classList.remove('active');
  // 释放 Data URL 内存
  $('dlLink').href = '';
});

// --- 图片加载 ---
$('imageInput').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (!file) return;
  originalFileName = file.name.replace(/\.[^.]+$/, '');
  const reader = new FileReader();
  reader.onload = function(ev) {
    imageEl.src = ev.target.result;
    imageEl.style.display = 'block';
    if (cropper) cropper.destroy();
    cropper = new Cropper(imageEl, {
      viewMode: 1,
      dragMode: cropEditMode ? 'crop' : 'none',
      guides: false,
      crop() { updateGridOverlay(); },
      ready() {
        const cols = clampInt('cols', 1, 200);
        const rows = clampInt('rows', 1, 200);
        if (aspectLocked) cropper.setAspectRatio(cols / rows);
        // 初始化时应用当前模式（手机端默认浏览模式）
        applyCropMode();
        // 初始化像素/格显示
        const box = cropper.getCropBoxData();
        $('cellPx').value = Math.round(box.width / cols);
      }
    });
    setStatus('图片已加载，请框选区域后点击"生成预览"。');
  };
  reader.readAsDataURL(file);
});

// --- Task 1: 行列比例联动 ---
['cols', 'rows'].forEach(id => {
  $(id).addEventListener('input', () => {
    if (!cropper) return;
    const cols = clampInt('cols', 1, 200);
    const rows = clampInt('rows', 1, 200);
    if (aspectLocked) cropper.setAspectRatio(cols / rows);
    updateGridOverlay();
  });
});

// --- Task 7: 预览 Tooltip ---
let previewCellSize = 15;
let gridCols = 0, gridRows = 0;
let basePreviewCellSize = 15; // 未缩放时的基础格子大小
let zoomLevel = 1;            // 当前缩放倍率

// --- 缩放功能 ---
function applyZoom() {
  const canvas = previewCanvas;
  canvas.style.width  = (canvas.width  * zoomLevel) + 'px';
  canvas.style.height = (canvas.height * zoomLevel) + 'px';
  $('zoomLabel').textContent = Math.round(zoomLevel * 100) + '%';
}

function setZoom(newZoom) {
  zoomLevel = Math.max(0.5, Math.min(5, newZoom));
  applyZoom();
}

$('zoomIn').addEventListener('click',  () => setZoom(zoomLevel + 0.5));
$('zoomOut').addEventListener('click', () => setZoom(zoomLevel - 0.5));
$('zoomReset').addEventListener('click', () => setZoom(1));

$('previewScroll').addEventListener('wheel', e => {
  if (!currentGridData) return;
  e.preventDefault();
  const delta = e.deltaY < 0 ? 0.25 : -0.25;
  setZoom(zoomLevel + delta);
}, { passive: false });

previewCanvas.addEventListener('mousemove', e => {
  if (!currentGridData) return;
  const rect = previewCanvas.getBoundingClientRect();
  // 考虑 CSS 缩放后的实际像素比
  const scaleX = previewCanvas.width / rect.width;
  const scaleY = previewCanvas.height / rect.height;
  const x = (e.clientX - rect.left) * scaleX;
  const y = (e.clientY - rect.top)  * scaleY;
  const c = Math.floor(x / previewCellSize);
  const r = Math.floor(y / previewCellSize);
  if (r < 0 || r >= gridRows || c < 0 || c >= gridCols) { tooltip.style.display = 'none'; return; }
  const cell = currentGridData[r][c];
  if (!cell) { tooltip.style.display = 'none'; return; }

  const swatch = tooltip.querySelector('.color-swatch');
  const text   = tooltip.querySelector('.tooltip-text');
  if (cell.isEmpty) {
    swatch.style.background = '#e5e7eb';
    text.textContent = '(空/忽略)';
  } else {
    const hex = '#' + cell.rgb.map(v => v.toString(16).padStart(2,'0')).join('');
    swatch.style.background = hex;
    text.textContent = `${cell.code}  RGB(${cell.rgb.join(', ')})`;
  }
  tooltip.style.display = 'block';
  tooltip.style.left = (e.clientX + 12) + 'px';
  tooltip.style.top  = (e.clientY + 12) + 'px';
});

previewCanvas.addEventListener('mouseleave', () => { tooltip.style.display = 'none'; });

// --- 生成预览 (Task 5, 6, 7, 10) ---
$('btnPreview').addEventListener('click', generatePreview);

function generatePreview() {
  if (!cropper) return alert('请先选择图片！');

  // Task 2: 参数校验
  const cols = clampInt('cols', 1, 200);
  const rows = clampInt('rows', 1, 200);
  const ignoreWhite   = $('ignoreWhite').checked;
  const showCodes     = $('showCodes').checked;
  const whiteThreshold = parseInt($('whiteThreshold').value); // 70~98，默认90

  setStatus('正在使用 LAB 算法进行色彩匹配...');
  showProgress(10);

  // 全分辨率裁切，每格可采样数百像素，保证色彩准确度
  const croppedCanvas = cropper.getCroppedCanvas();
  const ctx = croppedCanvas.getContext('2d', { willReadFrequently: true });
  const cw = croppedCanvas.width, ch = croppedCanvas.height;

  gridCols = cols; gridRows = rows;
  currentGridData = [];
  currentCounts = {};

  // 预览格子大小：根据列数自适应
  previewCellSize = Math.max(8, Math.min(20, Math.floor(400 / Math.max(cols, rows))));
  previewCanvas.width  = cols * previewCellSize;
  previewCanvas.height = rows * previewCellSize;

  // 整数化每格尺寸，保证跨设备确定性（避免浮点边界不一致）
  const cellW = Math.floor(cw / cols);
  const cellH = Math.floor(ch / rows);

  // 一次性读取全部像素
  const allPixels = ctx.getImageData(0, 0, cw, ch).data;

  // 每格取中心 65% 区域的像素均值
  const insetX = Math.floor(cellW * 0.175);
  const insetY = Math.floor(cellH * 0.175);
  const sampleW = cellW - 2 * insetX;
  const sampleH = cellH - 2 * insetY;

  for (let r = 0; r < rows; r++) {
    const rowData = [];
    for (let c = 0; c < cols; c++) {
      let sumR = 0, sumG = 0, sumB = 0, count = 0;
      const startX = c * cellW + insetX;
      const startY = r * cellH + insetY;
      for (let y = startY; y < startY + sampleH; y++) {
        for (let x = startX; x < startX + sampleW; x++) {
          const idx = (y * cw + x) * 4;
          sumR += allPixels[idx];
          sumG += allPixels[idx + 1];
          sumB += allPixels[idx + 2];
          count++;
        }
      }
      const avgRgb = [Math.round(sumR / count), Math.round(sumG / count), Math.round(sumB / count)];

      // Task 6: 判断是否为白色/透明（阈值可调）
      const lab = rgbToLab(avgRgb);
      const isWhite = ignoreWhite && lab[0] > whiteThreshold && Math.abs(lab[1]) < 8 && Math.abs(lab[2]) < 8;

      if (isWhite) {
        rowData.push({ code: '-', rgb: avgRgb, isEmpty: true });
        // 画灰色格子
        previewCtx.fillStyle = '#e5e7eb';
        previewCtx.fillRect(c * previewCellSize, r * previewCellSize, previewCellSize, previewCellSize);
      } else {
        const matched = nearestColor(avgRgb);
        rowData.push({ code: matched.code, rgb: matched.rgb, isEmpty: false });
        currentCounts[matched.code] = (currentCounts[matched.code] || 0) + 1;
        previewCtx.fillStyle = `rgb(${matched.rgb[0]},${matched.rgb[1]},${matched.rgb[2]})`;
        previewCtx.fillRect(c * previewCellSize, r * previewCellSize, previewCellSize, previewCellSize);

        // Task 7: 显示色号文字
        if (showCodes && previewCellSize >= 10) {
          const lum = 0.2126*matched.rgb[0] + 0.7152*matched.rgb[1] + 0.0722*matched.rgb[2];
          previewCtx.fillStyle = lum < 115 ? 'rgba(255,255,255,0.85)' : 'rgba(0,0,0,0.75)';
          previewCtx.font = `bold ${Math.max(6, previewCellSize * 0.45)}px system-ui`;
          previewCtx.textAlign = 'center';
          previewCtx.textBaseline = 'middle';
          previewCtx.fillText(matched.code, c*previewCellSize + previewCellSize/2, r*previewCellSize + previewCellSize/2);
        }
      }
      // 格子边框
      previewCtx.strokeStyle = 'rgba(255,255,255,0.25)';
      previewCtx.strokeRect(c*previewCellSize, r*previewCellSize, previewCellSize, previewCellSize);
    }
    currentGridData.push(rowData);
  }

  showProgress(100);
  setTimeout(hideProgress, 400);

  // 应用缩放
  basePreviewCellSize = previewCellSize;
  zoomLevel = 1;
  applyZoom();

  previewStale = false; // 预览已更新，标记为新鲜
  renderStats();
  setStatus(`预览已生成：${cols}×${rows} = ${cols*rows} 格。满意后点击"导出 Excel"。`);
}

// --- Task 5: 色豆统计面板 ---
function renderStats() {
  const entries = Object.entries(currentCounts);
  if (!entries.length) {
    statsPanel.className = 'stats-panel empty';
    statsPanel.innerHTML = '预览生成后将显示色豆用量统计';
    return;
  }
  statsPanel.className = 'stats-panel';

  // 按数量降序
  entries.sort((a, b) => b[1] - a[1]);
  const total = entries.reduce((s, [, c]) => s + c, 0);
  const palette = getCurrentPalette();

  let html = `<div class="stats-header"><h4>色豆用量统计</h4><span class="total">共 ${total} 颗（${entries.length} 种颜色）</span></div>`;
  html += '<div class="stats-grid">';
  for (const [code, count] of entries) {
    const colorObj = palette.find(p => p.code === code);
    const hex = colorObj ? '#' + colorObj.rgb.map(v => v.toString(16).padStart(2,'0')).join('') : '#ccc';
    html += `<div class="stat-item">
      <span class="stat-swatch" style="background:${hex}"></span>
      <span class="stat-code">${code}</span>
      <span class="stat-count">×${count}</span>
    </div>`;
  }
  html += '</div>';
  statsPanel.innerHTML = html;
}

// --- Excel 导出 (Task 3, 9, 10, 11) ---
$('btnExport').addEventListener('click', exportExcel);

function rgbToArgbHex(rgb) {
  return 'FF' + rgb.map(x => x.toString(16).padStart(2, '0').toUpperCase()).join('');
}

function readableFontColor(rgb) {
  return (0.2126*rgb[0] + 0.7152*rgb[1] + 0.0722*rgb[2]) < 115 ? 'FFFFFFFF' : 'FF000000';
}

async function exportExcel() {
  if (!currentGridData) return alert('请先点击“生成预览”计算网格数据！');
  if (previewStale) {
    if (!confirm('参数已修改，当前导出的将是上一次预览结果。\n是否先重新生成预览？')) return;
    generatePreview();
  }

  // Task 2: 参数校验
  const cols = clampInt('cols', 1, 200);
  const rows = clampInt('rows', 1, 200);
  const pixelSize  = clampInt('pixelSize', 8, 48);
  const boardSize  = clampInt('boardSize', 1, 100);
  const colorOnly  = $('colorOnly').checked;
  const includeLegend = $('includeLegend').checked;

  setStatus('正在打包生成 Excel 文件...');
  showProgress(5);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Pixel Art', { views: [{ showGridLines: false }] });

  const colWidth  = pixelSize <= 12 ? pixelSize/12 : (pixelSize-5)/7;
  const rowHeight = pixelSize * 72 / 96;

  for (let c = 1; c <= cols; c++) sheet.getColumn(c).width = colWidth;
  for (let r = 1; r <= rows; r++) sheet.getRow(r).height = rowHeight;

  // Task 3: 钉板粗线辅助线

  function getBorders(r, c) {
    // r, c 是 0-based 索引；钉板边界处画粗线
    const thick = { style: 'medium', color: { argb: 'FFD32F2F' } };
    const thin  = { style: 'thin',   color: { argb: 'FFC9D1D9' } };
    return {
      top:    r % boardSize === 0                          ? thick : thin,
      bottom: (r + 1) % boardSize === 0 || r === rows - 1  ? thick : thin,
      left:   c % boardSize === 0                          ? thick : thin,
      right:  (c + 1) % boardSize === 0 || c === cols - 1  ? thick : thin,
    };
  }

  // 写入网格
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cellData = currentGridData[r][c];
      const cell = sheet.getCell(r + 1, c + 1);
      const borders = getBorders(r, c);

      if (cellData.isEmpty) {
        // 空白格：浅灰填充，视觉上与有颜色格子区分
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } };
        cell.border = borders;
        continue;
      }

      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rgbToArgbHex(cellData.rgb) } };
      cell.border = borders;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      if (!colorOnly) {
        cell.value = cellData.code;
        cell.font = { name: 'Arial', size: 7, color: { argb: readableFontColor(cellData.rgb) } };
      }
    }
    // Task 10: 进度
    if (r % 5 === 0) showProgress(Math.round((r / rows) * 80) + 5);
  }

  // Task 11: 图例
  if (includeLegend) {
    const legendCol = cols + 2;
    sheet.getColumn(legendCol).width = 12;
    sheet.getColumn(legendCol + 1).width = 10;
    sheet.getColumn(legendCol + 2).width = 14;

    const headerFont = { bold: true, size: 11 };
    sheet.getCell(1, legendCol).value = '色号';
    sheet.getCell(1, legendCol).font = headerFont;
    sheet.getCell(1, legendCol + 1).value = '数量';
    sheet.getCell(1, legendCol + 1).font = headerFont;
    sheet.getCell(1, legendCol + 2).value = '颜色';
    sheet.getCell(1, legendCol + 2).font = headerFont;

    const palette = getCurrentPalette();
    let rowOff = 2;
    for (const [code, count] of Object.entries(currentCounts).sort((a,b) => b[1]-a[1])) {
      const colorObj = palette.find(p => p.code === code);
      if (!colorObj) continue;
      sheet.getCell(rowOff, legendCol).value = code;
      sheet.getCell(rowOff, legendCol + 1).value = count;
      const swatch = sheet.getCell(rowOff, legendCol + 2);
      swatch.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rgbToArgbHex(colorObj.rgb) } };
      swatch.border = { top: {style:'thin',color:{argb:'FFC9D1D9'}}, bottom: {style:'thin',color:{argb:'FFC9D1D9'}}, left: {style:'thin',color:{argb:'FFC9D1D9'}}, right: {style:'thin',color:{argb:'FFC9D1D9'}} };
      rowOff++;
    }
  }

  showProgress(95);

  // Task 9: 文件命名
  const name = $('fileName').value.trim() || originalFileName || 'pixel_art';
  const filename = `${name}_${Date.now()}.xlsx`;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

  // === 策略一：Web Share API（iOS 15.4+ / Android Chrome）—— 最原生 ===
  // 整体 try-catch：canShare()/new File() 在部分浏览器会抛异常而非返回 false
  try {
    const file = new File([blob], filename, { type: blob.type });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], title: '拼豆 Excel' });
      showProgress(100); setTimeout(hideProgress, 500);
      setStatus(`Excel 已分享：${filename}`);
      return;
    }
  } catch (e) {
    // canShare / File 构造 / share 抛错 → 回退到平台专用方案
  }

  // === 策略二：平台回退 ===
  const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) ||
    (/Macintosh/.test(navigator.userAgent) && navigator.maxTouchPoints > 1);

  if (isIOS) {
    // iOS 无 Share API（< 15.4）：弹窗 + Data URL 长按下载
    // iOS Safari 忽略 download 属性，只能通过长按链接触发保存
    showIOSDownloadModal(blob, filename);
  } else {
    // 桌面端 + Android：Blob URL + <a download> 均可正常触发下载
    triggerBlobDownload(blob, filename);
  }
}

function triggerBlobDownload(blob, filename) {
  const blobUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = blobUrl;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  setTimeout(() => {
    document.body.removeChild(link);
    URL.revokeObjectURL(blobUrl);
  }, 3000);
  showProgress(100); setTimeout(hideProgress, 500);
  setStatus(`Excel 生成完毕，已下载：${filename}`);
}

function showIOSDownloadModal(blob, filename) {
  const reader = new FileReader();
  reader.onload = function() {
    $('dlFilename').textContent = filename;
    $('dlLink').href = reader.result;
    $('dlLink').download = filename;
    $('dlHint').textContent = '请长按上方蓝色按钮，选择「下载链接」保存文件';
    $('dlOverlay').classList.add('active');
    showProgress(100); setTimeout(hideProgress, 500);
    setStatus(`Excel 已生成：${filename}`);
  };
  reader.readAsDataURL(blob);
}
