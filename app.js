/* global Html5Qrcode, Html5QrcodeSupportedFormats, XLSX */
'use strict';

const $ = (s) => document.querySelector(s);
const $$ = (s) => document.querySelectorAll(s);

/* ================= 设置存储 ================= */
const Settings = {
  get(key, def = '') {
    try { return localStorage.getItem('pd_' + key) ?? def; } catch { return def; }
  },
  set(key, val) { localStorage.setItem('pd_' + key, val); },
};

/* ================= 账号 ================= */
let presetAccounts = [];   // 来自 data/账号名单.json 的固定账号
let presetAdminPin = '';   // 管理员密码：删除记录 / 清空数据时校验
async function loadPresetAccounts() {
  const files = ['data/账号名单.json', 'data/accounts.json'];
  for (const f of files) {
    try {
      // 加时间戳 + no-store：确保每次打开都取到最新的账号名单，不被缓存挡掉
      const res = await fetch(f + '?t=' + Date.now(), { cache: 'no-store' });
      if (!res.ok) continue;
      const text = (await res.text()).replace(/^\uFEFF/, '');
      const data = JSON.parse(text);
      presetAdminPin = data.adminPassword ? String(data.adminPassword) : '';
      presetAccounts = (data.accounts || [])
        .filter((a) => a && a.name && a.password != null)
        .map((a) => ({ name: String(a.name).trim(), pin: String(a.password) }));
      return;
    } catch {}
  }
}

/* 管理员密码校验：配置了 adminPassword 就用它，否则退回当前账号密码 */
function checkAdminPin(pin) {
  if (presetAdminPin) return pin === presetAdminPin;
  return Auth.verify(Auth.current(), pin);
}

const Auth = {
  local() {
    try { return JSON.parse(localStorage.getItem('pd_accts') || '[]'); } catch { return []; }
  },
  save(list) { localStorage.setItem('pd_accts', JSON.stringify(list)); },
  hash(s) {
    let h = 5381;
    for (let i = 0; i < s.length; i++) h = ((h * 33) ^ s.charCodeAt(i)) >>> 0;
    return h.toString(16) + '-' + s.length;
  },
  /* 预置账号优先，本机临时账号补充在后（同名以预置为准） */
  list() {
    const names = new Set(presetAccounts.map((a) => a.name));
    return [
      ...presetAccounts.map((a) => ({ name: a.name, pin: a.pin, preset: true })),
      ...this.local().filter((a) => !names.has(a.name)),
    ];
  },
  find(name) { return this.list().find((a) => a.name === name); },
  add(name, pin) {
    const list = this.local();
    if (this.list().some((a) => a.name === name)) return false;
    list.push({ name, pin: this.hash(pin), created: Date.now() });
    this.save(list);
    return true;
  },
  verify(name, pin) {
    const a = this.find(name);
    if (!a) return false;
    return a.preset ? a.pin === pin : a.pin === this.hash(pin);
  },
  touch(name) {
    const list = this.local();
    const a = list.find((x) => x.name === name);
    if (a) { a.last = Date.now(); this.save(list); }
  },
  login(name) {
    localStorage.setItem('pd_session', name);
    Settings.set('operator', name);
  },
  logout() {
    localStorage.removeItem('pd_session');
    Settings.set('operator', '');
  },
  current() {
    const s = localStorage.getItem('pd_session') || '';
    return this.find(s) ? s : '';
  },
};

/* ================= IndexedDB ================= */
let db = null;
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('pandian', 1);
    req.onupgradeneeded = () => {
      const d = req.result;
      if (!d.objectStoreNames.contains('records')) {
        d.createObjectStore('records', { keyPath: 'id', autoIncrement: true });
      }
    };
    req.onsuccess = () => { db = req.result; resolve(db); };
    req.onerror = () => reject(req.error);
  });
}
function tx(mode, fn) {
  return new Promise((resolve, reject) => {
    const t = db.transaction('records', mode);
    const store = t.objectStore('records');
    const out = fn(store);
    t.oncomplete = () => resolve(out && out.result !== undefined ? out.result : undefined);
    t.onerror = () => reject(t.error);
  });
}
const dbAdd = (r) => tx('readwrite', (s) => s.add(r));
const dbPut = (r) => tx('readwrite', (s) => s.put(r));
const dbDel = (id) => tx('readwrite', (s) => s.delete(id));
const dbClear = () => tx('readwrite', (s) => s.clear());
function dbAll() {
  return new Promise((resolve, reject) => {
    const t = db.transaction('records', 'readonly');
    const req = t.objectStore('records').getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

/* ================= 物料数据 ================= */
let bySap = new Map();
let byAcc = new Map();
let matSource = '内置';

async function loadMaterials() {
  const custom = Settings.get('materials_custom');
  let data;
  if (custom) {
    try { data = JSON.parse(custom); matSource = '自定义导入'; }
    catch { data = null; }
  }
  if (!data) {
    const res = await fetch('data/materials.json');
    data = await res.json();
    matSource = '内置';
  }
  bySap = new Map();
  byAcc = new Map();
  for (const [k, v] of Object.entries(data.by_sap)) bySap.set(k, v);
  for (const [k, v] of Object.entries(data.by_acc)) byAcc.set(k.toUpperCase(), v);
  updateMatInfo();
}
function updateMatInfo() {
  $('#matInfo').textContent = `当前：${matSource}物料表，共 ${bySap.size} 个 SAP 代号 / ${byAcc.size} 个 Accpac 代号`;
}
function lookupMaterial(raw) {
  const code = raw.trim();
  if (/^\d{8}$/.test(code) && bySap.has(code)) {
    return { code, codeType: 'SAP', ...bySap.get(code) };
  }
  const up = code.toUpperCase();
  if (byAcc.has(up)) {
    const m = byAcc.get(up);
    return { code, codeType: 'ACCPAC', ...m };
  }
  // 8位数字但不在表里，仍按 SAP 记
  if (/^\d{8}$/.test(code)) return { code, codeType: 'SAP', desc: '' };
  return null;
}

/* ================= 录入流程状态 ================= */
const Flow = {
  location: Settings.get('location'),
  step: 'code',        // code | batch | qty
  current: null,       // 当前正在录入的物料
  onScan: null,        // 当前扫码回调
};

/* ================= 扫码器 ================= */
let scanner = null;
let scannerRunning = false;
let lastScan = { text: '', at: 0 };
let camTrack = null;
let camCaps = {};
let torchOn = false;
let lastScanCfg = null;

/* 扫码器启停串行化：避免上一步的停止还没完成就启动下一步，导致
   「Cannot clear while scan is ongoing」而落到“摄像头不可用” */
let scanChain = Promise.resolve();
function enqueueScan(fn) {
  const p = scanChain.then(fn, fn);
  scanChain = p.catch(() => {});
  return p;
}

function startScanner(onResult, tip) {
  return enqueueScan(() => doStartScanner(onResult, tip));
}

/* 把浏览器抛出的英文摄像头错误翻译成现场人员看得懂的提示 */
function describeCameraError(e) {
  const name = (e && e.name) || '';
  const msg = String((e && e.message) || e || '');
  if (name === 'NotAllowedError' || /Permission|denied/i.test(msg)) {
    return '摄像头权限被拒绝：请在浏览器地址栏左侧允许摄像头，或直接使用下方手动输入';
  }
  if (name === 'NotFoundError' || /no camera|Requested device not found/i.test(msg)) {
    return '没有检测到摄像头，请使用下方手动输入';
  }
  if (name === 'NotReadableError' || /Could not start video|in use/i.test(msg)) {
    return '摄像头被其他程序占用，关闭其他相机应用后重试，或用下方手动输入';
  }
  if (name === 'OverconstrainedError') {
    return '当前设备不支持该摄像头参数，请刷新页面重试，或用下方手动输入';
  }
  return '摄像头无法启动，请刷新页面重试，或用下方手动输入';
}

async function doStartScanner(onResult, tip) {
  await stopScannerNow();
  Flow.onScan = onResult;
  lastScanCfg = { onResult, tip };
  $('#readerWrap').classList.remove('hidden');
  $('#scanTip').textContent = (tip || '对准条码') + ' · 让条码尽量占满框';
  try {
    if (!scanner) {
      scanner = new Html5Qrcode('reader', {
        verbose: false,
        // 优先调用系统原生识别，对小码、轻微模糊的容忍度更高
        experimentalFeatures: { useBarCodeDetectorIfSupported: true },
      });
    }
    await scanner.start(
      { facingMode: 'environment' },
      {
        fps: 15,
        // 不设 qrbox：整幅画面参与解码，避免小码被裁剪丢失像素
        formatsToSupport: [
          Html5QrcodeSupportedFormats.QR_CODE,
          Html5QrcodeSupportedFormats.CODE_128,
          Html5QrcodeSupportedFormats.CODE_39,
          Html5QrcodeSupportedFormats.CODE_93,
          Html5QrcodeSupportedFormats.EAN_13,
          Html5QrcodeSupportedFormats.EAN_8,
          Html5QrcodeSupportedFormats.ITF,
          Html5QrcodeSupportedFormats.CODABAR,
          Html5QrcodeSupportedFormats.DATA_MATRIX,
        ],
      },
      (text) => {
        const now = Date.now();
        if (text === lastScan.text && now - lastScan.at < 1500) return; // 去抖
        lastScan = { text, at: now };
        beep();
        if (navigator.vibrate) navigator.vibrate(80);
        if (Flow.onScan) Flow.onScan(text.trim());
      },
      () => {}
    );
    scannerRunning = true;
    await tuneCamera();
  } catch (e) {
    scannerRunning = false;
    try { if (scanner) await scanner.clear(); } catch {}
    scanner = null;
    $('#scanTip').textContent = describeCameraError(e);
  }
}

/* 摄像头调优：高分辨率 + 连续对焦 + 补光 + 变焦 + 点击对焦 */
async function tuneCamera() {
  const video = document.querySelector('#reader video');
  if (!video) return;
  video.setAttribute('playsinline', 'true');
  const stream = video.srcObject;
  camTrack = stream && stream.getVideoTracks ? stream.getVideoTracks()[0] : null;
  if (!camTrack) return;
  camCaps = (camTrack.getCapabilities && camTrack.getCapabilities()) || {};

  const advanced = [];
  if (Array.isArray(camCaps.focusMode) && camCaps.focusMode.includes('continuous')) {
    advanced.push({ focusMode: 'continuous' });
  }
  try {
    await camTrack.applyConstraints({ width: { ideal: 1920 }, height: { ideal: 1080 }, advanced });
  } catch {}

  const hasTorch = !!camCaps.torch;
  torchOn = false;
  $('#btnTorch').classList.toggle('hidden', !hasTorch);
  $('#btnTorch').classList.remove('on');

  if (camCaps.zoom && camCaps.zoom.max > camCaps.zoom.min) {
    const r = $('#zoomRange');
    r.min = camCaps.zoom.min;
    r.max = camCaps.zoom.max;
    r.step = (camCaps.zoom.max - camCaps.zoom.min) / 20;
    let cur = camCaps.zoom.min;
    try { cur = camTrack.getSettings().zoom || camCaps.zoom.min; } catch {}
    // 自动放大到 2x：小二维码 / 小条码在画面里占的像素翻倍，识别率明显提升
    let target = cur;
    if (camCaps.zoom.min <= 1.05 && camCaps.zoom.max >= 2) target = 2;
    if (target !== cur) {
      try {
        await camTrack.applyConstraints({ advanced: [{ zoom: target }] });
      } catch { target = cur; }
    }
    r.value = target;
    if (target > cur) $('#scanTip').textContent += ' · 已放大 ' + (Math.round(target * 10) / 10) + 'x';
    $('#zoomWrap').classList.remove('hidden');
  } else {
    $('#zoomWrap').classList.add('hidden');
  }

  video.onclick = async (e) => {
    const box = video.getBoundingClientRect();
    const x = Math.min(1, Math.max(0, (e.clientX - box.left) / box.width));
    const y = Math.min(1, Math.max(0, (e.clientY - box.top) / box.height));
    try {
      await camTrack.applyConstraints({ advanced: [{ pointsOfInterest: [{ x, y }], focusMode: 'single-shot' }] });
      toast('已对该点对焦');
      setTimeout(() => {
        camTrack.applyConstraints({ advanced: [{ focusMode: 'continuous' }] }).catch(() => {});
      }, 1600);
    } catch {
      toast('该手机不支持点击对焦，请前后移动调整距离');
    }
  };
  $('#btnTapFocus').classList.remove('hidden');
}

async function stopScannerNow() {
  $('#readerWrap').classList.add('hidden');
  $('#btnTapFocus').classList.add('hidden');
  $('#zoomWrap').classList.add('hidden');
  $('#btnTorch').classList.add('hidden');
  if (camTrack && torchOn) {
    try { await camTrack.applyConstraints({ advanced: [{ torch: false }] }); } catch {}
    torchOn = false;
  }
  camTrack = null;
  camCaps = {};
  if (scanner) {
    const withTimeout = (pr, ms) => Promise.race([pr.catch(() => {}), new Promise((r) => setTimeout(r, ms))]);
    try { if (scannerRunning) await withTimeout(scanner.stop(), 2000); } catch {}
    scannerRunning = false;
    // 释放实例，下一步重新创建，避免复用导致的 clear 状态异常
    try { await withTimeout(scanner.clear(), 1200); } catch {}
    scanner = null;
  }
}
function stopScanner() {
  return enqueueScan(() => stopScannerNow());
}

function bindCamTools() {
  $('#btnSnap').onclick = snapDecode;
  $('#btnTorch').onclick = async () => {
    if (!camTrack || !camCaps.torch) return;
    torchOn = !torchOn;
    try {
      await camTrack.applyConstraints({ advanced: [{ torch: torchOn }] });
      $('#btnTorch').classList.toggle('on', torchOn);
    } catch { toast('该手机无法开启补光'); }
  };
  $('#zoomRange').oninput = async (e) => {
    if (!camTrack || !camCaps.zoom) return;
    try { await camTrack.applyConstraints({ advanced: [{ zoom: +e.target.value }] }); } catch {}
  };
}
/* ============ 拍照识别（小码兜底） ============ */
async function snapDecode() {
  const video = document.querySelector('#reader video');
  if (!video || !video.videoWidth) { toast('摄像头还没准备好，请稍等'); return; }
  const cfg = lastScanCfg;
  const W = video.videoWidth, H = video.videoHeight;
  const full = document.createElement('canvas');
  full.width = W; full.height = H;
  full.getContext('2d').drawImage(video, 0, 0, W, H);

  await stopScanner();
  $('#readerWrap').classList.remove('hidden');
  $('#scanTip').textContent = `正在识别（${W}×${H}）…`;

  const text = await decodeCanvasVariants(full);
  if (text) {
    beep();
    if (navigator.vibrate) navigator.vibrate(80);
    if (Flow.onScan) Flow.onScan(text.trim());
    return;
  }
  toast('这张没识别出来：让条码占满取景框，或点画面手动对焦后再试', 4500);
  if (cfg) await startScanner(cfg.onResult, cfg.tip);
}

/* 同一张图生成多种候选：原图、放大 2 倍、放大+灰度增强 */
async function decodeCanvasVariants(src) {
  const tries = [
    { ratio: 1, scale: 1, gray: false },
    { ratio: 0.92, scale: 2, gray: false },
    { ratio: 0.92, scale: 2, gray: true },
  ];
  for (const t of tries) {
    const canvas = cropScale(src, t.ratio, t.scale, t.gray);
    const text = await decodeOne(canvas);
    if (text) return text;
  }
  return '';
}

function cropScale(src, ratio, scale, gray) {
  const sw = Math.floor(src.width * ratio);
  const sh = Math.floor(src.height * ratio);
  const sx = Math.floor((src.width - sw) / 2);
  const sy = Math.floor((src.height - sh) / 2);
  let outW = Math.floor(sw * scale);
  if (outW > 2600) outW = 2600;
  const k = outW / sw;
  const c = document.createElement('canvas');
  c.width = outW;
  c.height = Math.max(1, Math.floor(sh * k));
  const ctx = c.getContext('2d');
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(src, sx, sy, sw, sh, 0, 0, c.width, c.height);
  if (gray) {
    try {
      const img = ctx.getImageData(0, 0, c.width, c.height);
      const d = img.data;
      for (let i = 0; i < d.length; i += 4) {
        let g = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
        g = g < 110 ? 0 : g > 170 ? 255 : g;   // 增强黑白对比
        d[i] = d[i + 1] = d[i + 2] = g;
      }
      ctx.putImageData(img, 0, 0);
    } catch {}
  }
  return c;
}

async function decodeOne(canvas) {
  // 1) 系统原生识别（安卓 Chrome 支持，准确率高）
  if ('BarcodeDetector' in window) {
    try {
      const det = new window.BarcodeDetector();
      const res = await det.detect(canvas);
      if (res && res.length) return res[0].rawValue;
    } catch {}
  }
  // 2) 退回本地解码库
  try {
    const blob = await new Promise((r) => canvas.toBlob(r, 'image/png'));
    if (!blob) return '';
    const file = new File([blob], 'snap.png', { type: 'image/png' });
    const el = document.createElement('div');
    el.id = 'snapReader';
    el.style.display = 'none';
    document.body.appendChild(el);
    const q = new Html5Qrcode('snapReader', { verbose: false });
    const txt = await q.scanFile(file, false);
    try { q.clear(); } catch {}
    el.remove();
    return txt || '';
  } catch {
    return '';
  }
}
function beep() {
  try {
    const ctx = beep._ctx || (beep._ctx = new (window.AudioContext || window.webkitAudioContext)());
    const o = ctx.createOscillator();
    const g = ctx.createGain();
    o.connect(g); g.connect(ctx.destination);
    o.frequency.value = 1560; g.gain.value = 0.08;
    o.start(); o.stop(ctx.currentTime + 0.09);
  } catch {}
}

/* ================= 步骤 UI ================= */
function renderStep() {
  const card = $('#stepCard');
  const manualBox = $('#manualBox');
  const qtyForm = $('#qtyForm');
  manualBox.classList.add('hidden');
  qtyForm.classList.add('hidden');
  if (!Auth.current()) { stopScanner(); return; }

  if (!Flow.location) {
    card.innerHTML = `
      <div class="step-line"><span class="step-num">1</span>请先扫描库位二维码</div>
      <div class="desc-line"><span class="desc-text">确定所在库位后才能录入，如 AC102</span></div>`;
    manualBox.classList.remove('hidden');
    $('#manualInput').placeholder = '或手动输入库位编号';
    startScanner(handleLocationScan, '对准库位二维码');
    return;
  }

  if (Flow.step === 'code') {
    card.innerHTML = `
      <div class="step-line"><span class="step-num">1</span>扫描物料代号<span class="step-tag">SAP 优先 · 8位数字</span></div>`;
    manualBox.classList.remove('hidden');
    $('#manualInput').placeholder = '或手动输入物料代号';
    $('#manualInput').value = '';
    startScanner(handleCodeScan, '对准物料代号条码');
  } else if (Flow.step === 'batch') {
    const c = Flow.current;
    card.innerHTML = `
      <div class="step-line"><span class="step-num">2</span>扫描批次<span class="code-chip">${escapeHtml(c.code)}</span></div>
      <div class="desc-line">
        <span class="desc-text${c.desc ? '' : ' warn-text'}">${c.desc ? escapeHtml(c.desc) : '物料表中未找到，请核对代号'}</span>
        <button id="btnBackCode" class="link-btn">改代号</button>
      </div>`;
    $('#btnBackCode').onclick = () => { Flow.step = 'code'; Flow.current = null; renderStep(); };
    const descEl = card.querySelector('.desc-text');
    if (descEl) descEl.onclick = () => descEl.classList.toggle('expanded');
    manualBox.classList.remove('hidden');
    $('#manualInput').placeholder = '或手动输入批次号';
    $('#manualInput').value = '';
    startScanner(handleBatchScan, '对准批次条码');
  } else if (Flow.step === 'qty') {
    stopScanner();
    const c = Flow.current;
    card.innerHTML = `
      <div class="step-line"><span class="step-num">3</span>输入数量
        <span class="step-tag">
          <button id="btnEditCode2" class="link-btn">改代号</button>
          <button id="btnEditBatch2" class="link-btn">改批次</button>
        </span>
      </div>`;
    $('#btnEditCode2').onclick = () => { Flow.step = 'code'; Flow.current = null; renderStep(); };
    $('#btnEditBatch2').onclick = () => { Flow.step = 'batch'; renderStep(); };
    qtyForm.classList.remove('hidden');
    $('#summaryCard').innerHTML = `
      <div class="row"><span>物料代号</span><b>${c.code}（${c.codeType}）</b></div>
      ${c.desc ? `<div class="row"><span>描述</span><b>${escapeHtml(c.desc)}</b></div>` : ''}
      <div class="row"><span>批次</span><b>${escapeHtml(c.batch)}</b></div>
      <div class="row"><span>库位</span><b>${escapeHtml(Flow.location)}</b></div>`;
    $('#qtyInput').value = '';
    $('#expInput').value = '';
    $('#mfgInput').value = '';
    $('#unitCustom').value = '';
    selectUnit('KG');
    setTimeout(() => $('#qtyInput').focus(), 100);
  }
}
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));
}
function setLocation(loc) {
  Flow.location = loc;
  Settings.set('location', loc);
  $('#locText').textContent = loc;
  $('#locBanner').classList.remove('unset');
  $('#locBanner').classList.add('set');
}

/* ================= 扫码回调 ================= */
/* 从二维码/条码内容里提取库位号：
   兼容「R101」「r101」「区位代码：R101」「库位: AC-102」「https://x?loc=R101」等写法 */
function extractLocation(text) {
  let s = String(text || '');
  // 全角转半角
  s = s.replace(/[\uFF01-\uFF5E]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)).toUpperCase();
  const m = s.match(/[A-Z]{1,4}[\s\-_]*\d{2,4}/);
  if (!m) return '';
  const loc = m[0].replace(/[\s\-_]/g, '');
  return /^[A-Z]{1,4}\d{2,4}$/.test(loc) ? loc : '';
}

function handleLocationScan(text) {
  const loc = extractLocation(text);
  if (!loc) {
    toast(`「${String(text).trim().slice(0, 20)}」不像库位编码（应为字母+数字，如 AC102），已忽略`);
    return;
  }
  setLocation(loc);
  toast(`库位已切换：${loc}`);
  renderStep();
}
async function handleCodeScan(text) {
  const v = text.trim();
  if (!v) return;
  await stopScanner();
  const m = lookupMaterial(v);
  if (m) {
    Flow.current = { ...m };
    if (!m.desc) toast('该代号不在物料表中，描述为空');
  } else {
    Flow.current = { code: v, codeType: '未知', sap: '', acc: '', desc: '' };
    toast('未识别的代号，已按原样记录');
  }
  Flow.step = 'batch';
  renderStep();
}
async function handleBatchScan(text) {
  if (!text.trim()) return;
  Flow.current.batch = text.trim();
  Flow.step = 'qty';
  renderStep();
}

/* ================= 保存记录 ================= */
let selectedUnit = 'KG';
function selectUnit(u) {
  selectedUnit = u;
  $$('.unit-chip').forEach((c) => c.classList.toggle('sel', c.dataset.unit === u));
}
async function saveRecord(changeLoc) {
  const qty = parseFloat($('#qtyInput').value);
  if (isNaN(qty) || qty < 0) { toast('请输入有效的数量'); return; }
  const customU = $('#unitCustom').value.trim();
  const unit = customU || selectedUnit;
  const c = Flow.current;
  const rec = {
    location: Flow.location,
    code: c.code,
    codeType: c.codeType,
    sap: c.sap || (c.codeType === 'SAP' ? c.code : ''),
    acc: c.acc || (c.codeType === 'ACCPAC' ? c.code : ''),
    desc: c.desc || '',
    batch: c.batch,
    qty,
    unit,
    exp: $('#expInput').value || '',
    mfg: $('#mfgInput').value || '',
    operator: Settings.get('operator') || '未设置',
    time: new Date().toISOString(),
    synced: 0,
  };
  const id = await dbAdd(rec);
  renderTodayStats();
  if (changeLoc) {
    setLocation('');
    Flow.location = '';
    Settings.set('location', '');
    $('#locText').textContent = '未设置';
    $('#locBanner').classList.add('unset');
    $('#locBanner').classList.remove('set');
  }
  Flow.step = 'code';
  Flow.current = null;
  renderStep();
  toastWithUndo(`已保存：${rec.code} × ${qty} ${unit}`, async () => {
    await dbDel(id);
    renderTodayStats();
    toast('已撤销');
  });
}

/* ================= Toast ================= */
let toastTimer = null;
function toast(msg, ms = 2200) {
  const t = $('#toast');
  t.innerHTML = `<span>${escapeHtml(msg)}</span>`;
  t.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.add('hidden'), ms);
}
function toastWithUndo(msg, onUndo) {
  const t = $('#toast');
  t.innerHTML = '';
  const span = document.createElement('span');
  span.textContent = msg;
  const btn = document.createElement('button');
  btn.textContent = '撤销';
  btn.onclick = () => { t.classList.add('hidden'); onUndo(); };
  t.appendChild(span); t.appendChild(btn);
  t.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.add('hidden'), 5000);
}

/* ================= 统计 ================= */
async function renderTodayStats() {
  const all = (await dbAll()).filter((r) => !r.voided);
  const today = new Date().toDateString();
  const n = all.filter((r) => new Date(r.time).toDateString() === today).length;
  const last = all.slice().sort((a, b) => (a.time < b.time ? 1 : -1))[0];
  const left = last ? `上一条 ${last.code}${last.batch ? ' · ' + last.batch : ''}` : '还没有记录';
  $('#todayStats').innerHTML =
    `<span>${escapeHtml(left)}</span><span>今日 ${n} 条 · 共 ${all.length} 条</span>`;
}

/* ================= 清单页 ================= */
let listCache = [];
let showVoided = false;   // 是否查看回收站（已作废记录）
async function renderList() {
  listCache = (await dbAll()).sort((a, b) => (a.time < b.time ? 1 : -1));
  const bin = listCache.filter((r) => r.voided).length;
  $('#btnVoided').textContent = `回收站（${bin}）`;
  $('#btnPurge').classList.toggle('hidden', !showVoided || !bin);
  applyListFilter();
}
function applyListFilter() {
  const kw = $('#listFilter').value.trim().toLowerCase();
  const base = listCache.filter((r) => (showVoided ? r.voided : !r.voided));
  const rows = kw
    ? base.filter((r) =>
        [r.code, r.batch, r.location, r.desc].some((f) => (f || '').toLowerCase().includes(kw)))
    : base;
  $('#listCount').textContent = showVoided ? `回收站 ${rows.length} 条` : `共 ${rows.length} 条`;
  const box = $('#recordList');
  box.innerHTML = rows.map((r) => `
    <div class="record-card${r.voided ? ' voided' : ''}" data-id="${r.id}">
      <div class="top">
        <span class="code">${escapeHtml(r.code)}</span>
        <span class="qty">${r.qty} ${escapeHtml(r.unit)}</span>
      </div>
      <div class="mid">库位 ${escapeHtml(r.location)} · 批次 ${escapeHtml(r.batch)}</div>
      ${r.desc ? `<div class="desc">${escapeHtml(r.desc)}</div>` : ''}
      <div class="foot">
        <span><i class="sync-dot ${r.synced ? 'yes' : 'no'}"></i>${r.synced ? '已同步' : '未同步'} · ${formatTime(r.time)} · ${escapeHtml(r.operator)}</span>
        <span class="ops">
          ${r.voided
            ? '<button class="btn btn-outline act-restore">恢复</button>'
            : '<button class="btn btn-outline act-edit">编辑</button><button class="btn btn-danger act-del">作废</button>'}
        </span>
      </div>
    </div>`).join('');
  box.querySelectorAll('.act-edit').forEach((b) =>
    b.addEventListener('click', (e) => openEdit(+e.target.closest('.record-card').dataset.id)));
  box.querySelectorAll('.act-restore').forEach((b) =>
    b.addEventListener('click', async (e) => {
      const id = +e.target.closest('.record-card').dataset.id;
      const rec = listCache.find((x) => x.id === id);
      if (!rec) return;
      const { voided, voidedAt, ...keep } = rec;
      await dbPut(keep);
      await renderList(); renderTodayStats(); renderExportStats();
      toast('已恢复到清单');
    }));
  box.querySelectorAll('.act-del').forEach((b) =>
    b.addEventListener('click', async (e) => {
      const id = +e.target.closest('.record-card').dataset.id;
      const rec = listCache.find((x) => x.id === id);
      if (!rec) return;
      // 软删除：只标记作废，数据仍在本机，可从回收站恢复
      await dbPut({ ...rec, voided: 1, voidedAt: Date.now() });
      await renderList(); renderTodayStats(); renderExportStats();
      toastWithUndo('已移入回收站，数据不会丢失', async () => {
        const { voided: _v, voidedAt: _t, ...back } = rec;
        await dbPut(back);
        await renderList(); renderTodayStats(); renderExportStats();
        toast('已恢复该记录');
      });
    }));
}
function formatTime(iso) {
  const d = new Date(iso);
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getMonth() + 1}-${d.getDate()} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

/* ================= 编辑弹窗 ================= */
let editingId = null;
function openEdit(id) {
  const r = listCache.find((x) => x.id === id);
  if (!r) return;
  editingId = id;
  $('#editInfo').textContent = `${r.code} · 库位 ${r.location}`;
  $('#editBatch').value = r.batch;
  $('#editQty').value = r.qty;
  $('#editUnit').value = r.unit;
  $('#editExp').value = r.exp || '';
  $('#editMfg').value = r.mfg || '';
  $('#editModal').classList.remove('hidden');
}
async function saveEdit() {
  const r = listCache.find((x) => x.id === editingId);
  if (!r) return;
  const qty = parseFloat($('#editQty').value);
  if (isNaN(qty) || qty < 0) { toast('数量无效'); return; }
  r.batch = $('#editBatch').value.trim();
  r.qty = qty;
  r.unit = $('#editUnit').value.trim() || r.unit;
  r.exp = $('#editExp').value;
  r.mfg = $('#editMfg').value;
  r.synced = 0;
  await dbPut(r);
  $('#editModal').classList.add('hidden');
  renderList();
  toast('已保存修改');
}

/* ================= 导出 ================= */
const EXPORT_COLS = [
  ['库位', 'location'], ['物料代号', 'code'], ['代号类型', 'codeType'],
  ['SAP代号', 'sap'], ['ACCPAC代号', 'acc'], ['描述', 'desc'],
  ['批次', 'batch'], ['数量', 'qty'], ['单位', 'unit'],
  ['保质期至', 'exp'], ['生产日期', 'mfg'],
  ['录入时间', (r) => { const d = new Date(r.time); const p = (n) => String(n).padStart(2, '0'); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`; }],
  ['操作人', 'operator'],
];
function buildSheetRows(records) {
  const head = EXPORT_COLS.map(([h]) => h);
  const rows = records.map((r) => EXPORT_COLS.map(([, k]) => (typeof k === 'function' ? k(r) : r[k] ?? '')));
  return [head, ...rows];
}
async function exportFile(type) {
  const all = (await dbAll()).filter((r) => !r.voided).sort((a, b) => (a.time < b.time ? -1 : 1));
  if (!all.length) { toast('没有可导出的记录'); return; }
  const aoa = buildSheetRows(all);
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  const fname = `录入_${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
  if (type === 'xlsx') {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = aoa[0].map((_, i) => ({ wch: Math.max(10, ...aoa.map((row) => String(row[i] ?? '').length)) }));
    XLSX.utils.book_append_sheet(wb, ws, '录入记录');
    XLSX.writeFile(wb, fname + '.xlsx');
  } else {
    const csv = '﻿' + aoa.map((row) => row.map((c) => `"${String(c ?? '').replace(/"/g, '""')}"`).join(',')).join('\r\n');
    downloadBlob(new Blob([csv], { type: 'text/csv;charset=utf-8' }), fname + '.csv');
  }
  toast(`已导出 ${all.length} 条记录`);
}
function downloadBlob(blob, fname) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = fname;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

/* ================= 备份 / 恢复 ================= */
function stamp() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
}
async function backupAll(silent) {
  const records = await dbAll();
  const payload = {
    type: 'pandian-backup',
    version: 1,
    exportedAt: new Date().toISOString(),
    operator: Auth.current(),
    location: Settings.get('location'),
    webhook: Settings.get('webhook'),
    materialsCustom: localStorage.getItem('pd_materials_custom') || null,
    records,
  };
  downloadBlob(
    new Blob([JSON.stringify(payload)], { type: 'application/json' }),
    `数据备份_${stamp()}_${records.length}条.json`
  );
  Settings.set('lastBackup', new Date().toISOString());
  if (!silent) toast(`已备份 ${records.length} 条记录，请把文件存到网盘或发给自己`);
}
async function restoreBackup(file) {
  let data;
  try {
    data = JSON.parse(await file.text());
  } catch {
    toast('文件无法解析，请选择 .json 备份文件');
    return;
  }
  if (!data || data.type !== 'pandian-backup' || !Array.isArray(data.records)) {
    toast('不是本 APP 的备份文件');
    return;
  }
  const hasLocal = (await dbAll()).length > 0;
  let merge = false;
  if (hasLocal) {
    merge = confirm(
      `本机已有记录。\n\n点「确定」= 合并（保留本机记录，追加备份文件中的记录）\n点「取消」= 覆盖（清空本机后恢复备份）`
    );
  }
  if (!merge && hasLocal) {
    if (!confirm('覆盖会清空本机现有全部记录，确定吗？')) return;
    await dbClear();
  }
  const existKey = new Set(
    (await dbAll()).map((r) => [r.time, r.code, r.batch, r.location, r.qty, r.unit].join('|'))
  );
  let added = 0;
  for (const rec of data.records) {
    const key = [rec.time, rec.code, rec.batch, rec.location, rec.qty, rec.unit].join('|');
    if (existKey.has(key)) continue;
    existKey.add(key);
    const { id, ...rest } = rec;
    await dbAdd(rest);
    added++;
  }
  if (data.materialsCustom) localStorage.setItem('pd_materials_custom', data.materialsCustom);
  if (data.webhook) { Settings.set('webhook', data.webhook); $('#webhookInput').value = data.webhook; }
  if (data.location) { Flow.location = data.location; Settings.set('location', data.location); $('#locText').textContent = data.location; $('#locBanner').classList.add('set'); $('#locBanner').classList.remove('unset'); }
  await loadMaterials();
  renderExportStats();
  renderList();
  toast(`恢复完成：新增 ${added} 条（备份共 ${data.records.length} 条）`, 4000);
}
async function copyToClipboard() {
  const all = (await dbAll()).filter((r) => !r.voided).sort((a, b) => (a.time < b.time ? -1 : 1));
  if (!all.length) { toast('没有记录'); return; }
  const text = buildSheetRows(all).map((row) => row.join('\t')).join('\n');
  try {
    await navigator.clipboard.writeText(text);
    toast(`已复制 ${all.length} 条，打开飞书表格直接粘贴`);
  } catch {
    toast('复制失败，请用导出文件');
  }
}

/* ================= 飞书同步 ================= */
async function syncToFeishu() {
  const url = Settings.get('webhook');
  if (!url) { toast('请先在「设置」里填写 Webhook 地址'); switchView('view-settings'); return; }
  const unsynced = (await dbAll()).filter((r) => !r.synced && !r.voided);
  if (!unsynced.length) { toast('没有未同步的记录'); return; }
  toast(`正在同步 ${unsynced.length} 条…`, 4000);
  let ok = 0, fail = 0;
  for (const r of unsynced) {
    const payload = {};
    for (const [h, k] of EXPORT_COLS) payload[h] = typeof k === 'function' ? k(r) : r[k] ?? '';
    try {
      const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      if (res.ok) { r.synced = 1; await dbPut(r); ok++; }
      else fail++;
    } catch { fail++; }
  }
  toast(`同步完成：成功 ${ok} 条${fail ? `，失败 ${fail} 条` : ''}`, 3500);
  renderExportStats();
}

/* ================= 导出页统计 ================= */
async function renderExportStats() {
  const all = (await dbAll()).filter((r) => !r.voided);
  const today = new Date().toDateString();
  $('#statTotal').textContent = all.length;
  $('#statUnsync').textContent = all.filter((r) => !r.synced).length;
  $('#statToday').textContent = all.filter((r) => new Date(r.time).toDateString() === today).length;
}

/* ================= 物料表导入 ================= */
async function importMaterials(file) {
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf);
    let sheet = wb.Sheets['MARA - Basic Data'] || wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const by_sap = {}, by_acc = {};
    let count = 0;
    for (const row of rows) {
      const sap = String(row[0] ?? '').trim();
      const acc = String(row[1] ?? '').trim();
      const desc = String(row[2] ?? '').trim();
      if (!/^\d{8}$/.test(sap) || !desc) continue;
      const entry = { sap, acc, desc };
      by_sap[sap] = entry;
      for (const a of acc.split('/')) {
        const k = a.trim();
        if (k && !by_sap[k]) by_acc[k] = entry;
      }
      count++;
    }
    if (!count) { toast('导入失败：未找到有效行（需 SAP代号/ACCPAC代号/描述 三列）', 3500); return; }
    Settings.set('materials_custom', JSON.stringify({ by_sap, by_acc }));
    matSource = '自定义导入';
    bySap = new Map(Object.entries(by_sap));
    byAcc = new Map(Object.entries(by_acc).map(([k, v]) => [k.toUpperCase(), v]));
    updateMatInfo();
    toast(`已导入 ${count} 个物料`);
  } catch (e) {
    toast('导入失败：' + e.message, 3500);
  }
}

/* ================= 视图切换 ================= */
async function switchView(id) {
  await stopScanner();
  $$('.view').forEach((v) => v.classList.toggle('active', v.id === id));
  $$('.tab').forEach((t) => t.classList.toggle('active', t.dataset.view === id));
  if (id === 'view-home') renderStep();
  if (id === 'view-list') renderList();
  if (id === 'view-export') renderExportStats();
}

/* ================= 账号界面 ================= */
function showAuth() {
  $('#authOverlay').classList.remove('hidden');
  if (!Auth.list().length) showSetup(false);
  else showLogin();
}
function showSetup(isNew) {
  $('#authLogin').classList.add('hidden');
  $('#authSetup').classList.remove('hidden');
  $('#authBack').classList.toggle('hidden', !isNew);
  $('#authSub').textContent = isNew ? '新建一个录入账号' : '首次使用，请创建账号';
  $('#authName').value = '';
  $('#authPin').value = '';
  $('#authPin2').value = '';
  setTimeout(() => $('#authName').focus(), 100);
}
function showLogin() {
  const list = Auth.list();
  $('#authSetup').classList.add('hidden');
  $('#authLogin').classList.remove('hidden');
  const sel = $('#authUserSel');
  sel.innerHTML = list.map((a) => `<option value="${escapeHtml(a.name)}">${escapeHtml(a.name)}</option>`).join('');
  const last = localStorage.getItem('pd_last_user');
  if (last && list.some((a) => a.name === last)) sel.value = last;
  $('#authSub').textContent = '请输入密码继续';
  $('#authPinInput').value = '';
  setTimeout(() => $('#authPinInput').focus(), 100);
}
function unlockApp() {
  const name = Auth.current();
  $('#opBadge').textContent = name;
  $('#acctInfo').textContent = '当前账号：' + name;
  $('#authOverlay').classList.add('hidden');
  renderStep();
  renderTodayStats();
  maybeRemindBackup();
}
async function maybeRemindBackup() {
  const all = await dbAll();
  if (!all.length) return;
  const today = new Date().toDateString();
  const last = Settings.get('lastBackup');
  const reminded = Settings.get('lastRemind');
  if (last && new Date(last).toDateString() === today) return;
  if (reminded && new Date(reminded).toDateString() === today) return;
  Settings.set('lastRemind', new Date().toISOString());
  setTimeout(() => {
    toast(`今天还没备份：已有 ${all.length} 条记录，建议到「导出」页点「备份全部数据」`, 6000);
  }, 1500);
}
function bindAuth() {
  $('#authSetupOk').onclick = () => {
    const name = $('#authName').value.trim();
    const p1 = $('#authPin').value.trim();
    const p2 = $('#authPin2').value.trim();
    if (!name) { toast('请输入姓名'); return; }
    if (p1.length < 4) { toast('密码至少 4 位'); return; }
    if (p1 !== p2) { toast('两次输入的密码不一致'); return; }
    if (Auth.find(name)) { toast('该姓名已存在，请直接登录'); showLogin(); return; }
    Auth.add(name, p1);
    Auth.login(name);
    localStorage.setItem('pd_last_user', name);
    unlockApp();
    toast(`欢迎，${name}`);
  };
  $('#authBack').onclick = () => showLogin();
  $('#authLoginOk').onclick = () => {
    const name = $('#authUserSel').value;
    const pin = $('#authPinInput').value.trim();
    if (!Auth.verify(name, pin)) { toast('密码错误'); return; }
    Auth.login(name);
    Auth.touch(name);
    localStorage.setItem('pd_last_user', name);
    unlockApp();
  };
  $('#authPinInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') $('#authLoginOk').click(); });
  $('#authNewAcct').onclick = () => showSetup(true);
  $('#btnLogout').onclick = () => {
    if (!confirm('确定退出当前账号？')) return;
    Auth.logout();
    stopScanner();
    Flow.step = 'code'; Flow.current = null;
    showAuth();
  };
  $('#btnAddAcct').onclick = () => { $('#authOverlay').classList.remove('hidden'); showSetup(true); };
}

/* ================= 初始化 ================= */
async function init() {
  await openDB();
  await loadMaterials();
  await loadPresetAccounts();
  bindAuth();
  bindCamTools();

  $('#webhookInput').value = Settings.get('webhook');

  // 库位恢复
  if (Flow.location) {
    $('#locText').textContent = Flow.location;
    $('#locBanner').classList.add('set');
  } else {
    $('#locBanner').classList.add('unset');
  }

  // 事件绑定
  $('#btnScanLoc').onclick = () => {
    Flow.step = 'code'; Flow.current = null;
    Flow.location = ''; Settings.set('location', '');
    $('#locText').textContent = '未设置';
    $('#locBanner').classList.add('unset'); $('#locBanner').classList.remove('set');
    switchView('view-home');
  };
  $('#manualOk').onclick = () => {
    const v = $('#manualInput').value.trim();
    if (!v) return;
    if (!Flow.location) handleLocationScan(v);
    else if (Flow.step === 'code') handleCodeScan(v);
    else if (Flow.step === 'batch') handleBatchScan(v);
  };
  $('#manualInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') $('#manualOk').click(); });
  $$('.unit-chip').forEach((c) => (c.onclick = () => { selectUnit(c.dataset.unit); $('#unitCustom').value = ''; }));
  $('#btnSaveNext').onclick = () => saveRecord(false);
  $('#btnSaveChloc').onclick = () => saveRecord(true);
  $$('.tab').forEach((t) => (t.onclick = () => switchView(t.dataset.view)));
  $('#listFilter').addEventListener('input', applyListFilter);
  $('#editSave').onclick = saveEdit;
  $('#editCancel').onclick = () => $('#editModal').classList.add('hidden');
  $('#btnExportXlsx').onclick = () => exportFile('xlsx');
  $('#btnExportCsv').onclick = () => exportFile('csv');
  $('#btnCopy').onclick = copyToClipboard;
  $('#btnSync').onclick = syncToFeishu;
  $('#btnClear').onclick = async () => {
    const all = (await dbAll()).filter((r) => !r.voided);
    if (!all.length) { toast('没有记录'); return; }
    await backupAll(true);
    const now = Date.now();
    for (const r of all) await dbPut({ ...r, voided: 1, voidedAt: now });
    showVoided = false;
    await renderList(); renderExportStats(); renderTodayStats();
    toast('已全部移入回收站（数据未删除，可在清单页「回收站」恢复），并已自动下载备份', 6000);
  };
  $('#btnVoided').onclick = () => {
    showVoided = !showVoided;
    renderList();
  };
  $('#btnPurge').onclick = async () => {
    const bin = (await dbAll()).filter((r) => r.voided);
    if (!bin.length) { toast('回收站是空的'); return; }
    const pin = prompt(`彻底清空回收站会永久删除 ${bin.length} 条记录，无法恢复。\n\n请输入管理员密码确认：`);
    if (pin === null) return;
    if (!checkAdminPin(pin.trim())) { toast('管理员密码错误，已取消'); return; }
    await backupAll(true);
    for (const r of bin) await dbDel(r.id);
    await renderList(); renderExportStats(); renderTodayStats();
    toast('回收站已彻底清空，并已自动下载备份', 4000);
  };
  $('#btnBackup').onclick = () => backupAll(false);
  $('#btnRestore').onclick = () => $('#restoreFile').click();
  $('#restoreFile').addEventListener('change', (e) => {
    const f = e.target.files[0];
    if (f) restoreBackup(f);
    e.target.value = '';
  });
  $('#btnSaveWebhook').onclick = () => {
    Settings.set('webhook', $('#webhookInput').value.trim());
    toast('已保存');
  };
  $('#opBadge').onclick = () => switchView('view-settings');
  $('#btnImportMat').onclick = () => $('#matFile').click();
  $('#matFile').addEventListener('change', (e) => { if (e.target.files[0]) importMaterials(e.target.files[0]); });
  $('#btnResetMat').onclick = async () => {
    localStorage.removeItem('pd_materials_custom');
    await loadMaterials();
    toast('已恢复内置物料表');
  };

  if (Auth.current()) unlockApp();
  else showAuth();

  // 请求持久化存储，降低浏览器自动清理本地数据的概率
  if (navigator.storage && navigator.storage.persist) {
    navigator.storage.persist().catch(() => {});
  }

  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
}
document.addEventListener('DOMContentLoaded', init);
