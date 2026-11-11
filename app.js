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

/* ================= 盘点流程状态 ================= */
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

async function startScanner(onResult, tip) {
  await stopScanner();
  Flow.onScan = onResult;
  $('#readerWrap').classList.remove('hidden');
  $('#scanTip').textContent = tip || '将条码对准取景框';
  try {
    if (!scanner) scanner = new Html5Qrcode('reader', { verbose: false });
    await scanner.start(
      { facingMode: 'environment' },
      {
        fps: 10,
        qrbox: (w, h) => {
          const width = Math.floor(Math.min(w, h) * 0.9);
          return { width, height: Math.floor(width * 0.5) };
        },
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
  } catch (e) {
    $('#scanTip').textContent = '摄像头不可用：' + (e && e.message ? e.message : e) + '，请使用下方手动输入';
  }
}
async function stopScanner() {
  $('#readerWrap').classList.add('hidden');
  if (scanner && scannerRunning) {
    try { await scanner.stop(); } catch {}
    scannerRunning = false;
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

  if (!Flow.location) {
    card.innerHTML = `
      <p class="step-title">请先扫描库位二维码</p>
      <p class="step-sub">盘点前需要确定所在库位（如 AC102、R103）</p>`;
    manualBox.classList.remove('hidden');
    $('#manualInput').placeholder = '或手动输入库位编号';
    startScanner(handleLocationScan, '对准库位二维码');
    return;
  }

  if (Flow.step === 'code') {
    card.innerHTML = `
      <div class="step-dots"><i class="on"></i><i></i><i></i></div>
      <p class="step-title">第 1 步 · 扫描物料代号条码</p>
      <p class="step-sub">优先扫 SAP 物料代号（8位数字）；标签上没有 SAP 就扫 Accpac 代号</p>`;
    manualBox.classList.remove('hidden');
    $('#manualInput').placeholder = '或手动输入物料代号';
    $('#manualInput').value = '';
    startScanner(handleCodeScan, '对准物料代号条码');
  } else if (Flow.step === 'batch') {
    const c = Flow.current;
    card.innerHTML = `
      <div class="step-dots"><i class="on"></i><i class="on"></i><i></i></div>
      <p class="step-title">第 2 步 · 扫描批次条码</p>
      ${foundCardHtml(c)}
      <div class="btn-row">
        <button class="btn btn-outline btn-small" id="btnRescanCode">重扫代号</button>
      </div>`;
    $('#btnRescanCode').onclick = () => { Flow.step = 'code'; Flow.current = null; renderStep(); };
    manualBox.classList.remove('hidden');
    $('#manualInput').placeholder = '或手动输入批次号';
    $('#manualInput').value = '';
    startScanner(handleBatchScan, '对准批次条码');
  } else if (Flow.step === 'qty') {
    stopScanner();
    const c = Flow.current;
    card.innerHTML = `
      <div class="step-dots"><i class="on"></i><i class="on"></i><i class="on"></i></div>
      <p class="step-title">第 3 步 · 输入数量</p>`;
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
function foundCardHtml(c, warn) {
  return `
    <div class="found-card${warn ? ' warn' : ''}">
      <div class="row"><span>代号</span><b>${escapeHtml(c.code)} <span class="tag${c.codeType === 'ACCPAC' ? ' acc' : ''}">${c.codeType}</span></b></div>
      ${c.sap && c.codeType === 'ACCPAC' ? `<div class="row"><span>对应SAP</span><b>${escapeHtml(c.sap)}</b></div>` : ''}
      <div class="row"><span>描述</span><b>${c.desc ? escapeHtml(c.desc) : '<span style="color:#fa8c16">物料表中未找到，请手动补充</span>'}</b></div>
    </div>`;
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
function handleLocationScan(text) {
  const loc = text.trim().toUpperCase();
  if (!/^[A-Z]{1,4}\d{2,4}$/.test(loc)) {
    toast(`「${loc}」不像库位编码（应为字母+数字，如 AC102），已忽略`);
    return;
  }
  setLocation(loc);
  toast(`库位已切换：${loc}`);
  renderStep();
}
async function handleCodeScan(text) {
  const m = lookupMaterial(text);
  await stopScanner();
  if (m && (m.desc || m.sap)) {
    Flow.current = { ...m };
    Flow.step = 'batch';
    renderStep();
  } else if (m) {
    // SAP格式但不在表中
    Flow.current = { ...m };
    Flow.step = 'batch';
    renderStep();
    toast('该代号不在物料表中，描述为空');
  } else {
    // 完全不认识，询问是否仍作为代号使用
    Flow.current = { code: text.trim(), codeType: '未知', sap: '', acc: '', desc: '' };
    Flow.step = 'batch';
    renderStep();
    toast('未识别的代号，已按原样记录');
  }
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
  const all = await dbAll();
  const today = new Date().toDateString();
  const n = all.filter((r) => new Date(r.time).toDateString() === today).length;
  $('#todayStats').textContent = `今日已盘点 ${n} 条 · 本机共 ${all.length} 条`;
}

/* ================= 清单页 ================= */
let listCache = [];
async function renderList() {
  listCache = (await dbAll()).sort((a, b) => (a.time < b.time ? 1 : -1));
  applyListFilter();
}
function applyListFilter() {
  const kw = $('#listFilter').value.trim().toLowerCase();
  const rows = kw
    ? listCache.filter((r) =>
        [r.code, r.batch, r.location, r.desc].some((f) => (f || '').toLowerCase().includes(kw)))
    : listCache;
  $('#listCount').textContent = `共 ${rows.length} 条`;
  const box = $('#recordList');
  box.innerHTML = rows.map((r) => `
    <div class="record-card" data-id="${r.id}">
      <div class="top">
        <span class="code">${escapeHtml(r.code)}</span>
        <span class="qty">${r.qty} ${escapeHtml(r.unit)}</span>
      </div>
      <div class="mid">库位 ${escapeHtml(r.location)} · 批次 ${escapeHtml(r.batch)}</div>
      ${r.desc ? `<div class="desc">${escapeHtml(r.desc)}</div>` : ''}
      <div class="foot">
        <span><i class="sync-dot ${r.synced ? 'yes' : 'no'}"></i>${r.synced ? '已同步' : '未同步'} · ${formatTime(r.time)} · ${escapeHtml(r.operator)}</span>
        <span class="ops">
          <button class="btn btn-outline act-edit">编辑</button>
          <button class="btn btn-danger act-del">删除</button>
        </span>
      </div>
    </div>`).join('');
  box.querySelectorAll('.act-edit').forEach((b) =>
    b.addEventListener('click', (e) => openEdit(+e.target.closest('.record-card').dataset.id)));
  box.querySelectorAll('.act-del').forEach((b) =>
    b.addEventListener('click', async (e) => {
      const id = +e.target.closest('.record-card').dataset.id;
      if (confirm('确定删除这条记录？')) { await dbDel(id); renderList(); renderTodayStats(); }
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
  ['盘点时间', (r) => { const d = new Date(r.time); const p = (n) => String(n).padStart(2, '0'); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`; }],
  ['操作人', 'operator'],
];
function buildSheetRows(records) {
  const head = EXPORT_COLS.map(([h]) => h);
  const rows = records.map((r) => EXPORT_COLS.map(([, k]) => (typeof k === 'function' ? k(r) : r[k] ?? '')));
  return [head, ...rows];
}
async function exportFile(type) {
  const all = (await dbAll()).sort((a, b) => (a.time < b.time ? -1 : 1));
  if (!all.length) { toast('没有可导出的记录'); return; }
  const aoa = buildSheetRows(all);
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  const fname = `盘点_${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
  if (type === 'xlsx') {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = aoa[0].map((_, i) => ({ wch: Math.max(10, ...aoa.map((row) => String(row[i] ?? '').length)) }));
    XLSX.utils.book_append_sheet(wb, ws, '盘点记录');
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
async function copyToClipboard() {
  const all = (await dbAll()).sort((a, b) => (a.time < b.time ? -1 : 1));
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
  const unsynced = (await dbAll()).filter((r) => !r.synced);
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
  const all = await dbAll();
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

/* ================= 初始化 ================= */
async function init() {
  await openDB();
  await loadMaterials();

  // 操作人
  const op = Settings.get('operator');
  $('#opBadge').textContent = op || '点我设置操作人';
  $('#opInput').value = op;
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
    if (confirm('确定清空全部盘点记录？此操作不可恢复！')) {
      await dbClear(); renderExportStats(); renderTodayStats(); toast('已清空');
    }
  };
  $('#btnSaveOp').onclick = () => {
    const v = $('#opInput').value.trim();
    Settings.set('operator', v);
    $('#opBadge').textContent = v || '点我设置操作人';
    toast('已保存');
  };
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

  renderStep();
  renderTodayStats();

  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
}
document.addEventListener('DOMContentLoaded', init);
