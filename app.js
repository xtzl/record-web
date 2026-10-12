'use strict';

/* ============================================================================
   记录工作台 v1.0
   纯前端：无框架、无构建、无依赖、无后端。数据只存在本机浏览器 localStorage。
   ========================================================================== */

var STORE_KEY = 'record-workbench-v1';
var DRAFT_KEY = 'record-workbench-v1:draft';
var STATUSES = ['待确认', '已确认', '已处理', '已驳回'];
var TYPES = ['解绑', '绑定', '换绑', '查询', '其他'];
var STATUS_TONE = { '待确认': 'warn', '已确认': 'ok', '已处理': 'done', '已驳回': 'bad' };

/* 一句话成稿的模板：{编号}，{说明}，{确认人}确认了 */
var TEMPLATE_KEY = 'record-workbench-v1:template';
var DEFAULT_TEMPLATE = '{编号}，{说明}，{确认人}确认了';

/* 快捷句：点一下把「说明 + 确认人」填好，不用再手打 */
var PRESET_KEY = 'record-workbench-v1:presets';
var BUILTIN_PRESETS = [
  { desc: '辛苦您帮我解绑一下', owner: '运营助理' },
  { desc: '同一个连锁店下的，辛苦您帮我解绑一下', owner: '' }
];

/* ======================= 纯函数（可单独测试） ======================= */

function pad2(n) { return String(n).padStart(2, '0'); }

/** 2026-09-13 23:06 */
function stamp(ms) {
  var d = new Date(ms == null ? Date.now() : ms);
  return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) +
    ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes());
}

/** 20260913_230612 —— 用在文件名里 */
function fileStamp(ms) {
  var d = new Date(ms == null ? Date.now() : ms);
  return '' + d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) +
    '_' + pad2(d.getHours()) + pad2(d.getMinutes()) + pad2(d.getSeconds());
}

/** "23423234, 234234234 089797" → ['23423234','234234234','089797']（去重、保序） */
function splitCodes(raw) {
  if (!raw) return [];
  var seen = Object.create(null);
  var out = [];
  String(raw).split(/[\s,，、;；|/\\]+/).forEach(function (piece) {
    var v = piece.trim();
    if (!v || seen[v]) return;
    seen[v] = 1;
    out.push(v);
  });
  return out;
}

function codesText(codes) { return (codes || []).join('、'); }

function csvCell(v) {
  var s = v == null ? '' : String(v);
  return /[",\r\n\t]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

/** 加了 BOM，Excel 打开不乱码 */
function toCSV(rows) {
  return '\ufeff' + rows.map(function (r) { return r.map(csvCell).join(','); }).join('\r\n');
}

/** 简明行：制表符分隔，一行一条，粘到 Excel / 群里会自动分列 */
function composeSimple(list) {
  return list.map(function (r) {
    return [
      codesText(r.codes), r.type || '', r.desc || '', r.status || '',
      r.owner || '', r.note || '', stamp(r.createdAt)
    ].join('\t');
  }).join('\n');
}

/** 清单：带字段名的多行明细 */
function composeList(list) {
  return list.map(function (r) {
    var lines = ['【' + (r.type || '记录') + '】' + codesText(r.codes)];
    if (r.desc) lines.push('说明：' + r.desc);
    lines.push('状态：' + (r.status || '') + (r.owner ? '（' + r.owner + '）' : ''));
    if (r.note) lines.push('备注：' + r.note);
    lines.push('记录时间：' + stamp(r.createdAt));
    return lines.join('\n');
  }).join('\n\n');
}

/** 只有编号不同、其它字段都一样的记录并成一条，这样一句话能带上所有编号 */
function mergeForBrief(list) {
  var order = [];
  var bucket = Object.create(null);
  (list || []).forEach(function (r) {
    var key = [r.type, r.desc, r.status, r.owner, r.note].join('\u0000');
    if (!bucket[key]) {
      bucket[key] = {
        type: r.type, desc: r.desc, status: r.status, owner: r.owner, note: r.note,
        codes: [], createdAt: r.createdAt
      };
      order.push(key);
    }
    var b = bucket[key];
    var seen = Object.create(null);
    b.codes.forEach(function (c) { seen[c] = 1; });
    (r.codes || []).forEach(function (c) {
      if (!seen[c]) { seen[c] = 1; b.codes.push(c); }
    });
    if (r.createdAt && (!b.createdAt || r.createdAt < b.createdAt)) b.createdAt = r.createdAt;
  });
  return order.map(function (k) { return bucket[k]; });
}

/**
 * 一句话成稿：按模板拼。
 * 某一段里的占位符全是空的（比如没填确认人）→ 这一段连同它的「，」一起丢掉，
 * 所以不会出现 「123，，确认了」 这种脏文本。
 */
function renderBriefOne(r, tpl) {
  var map = {
    '编号': codesText(r.codes).split('、').join(','),
    '事项': r.type || '',
    '说明': r.desc || '',
    '状态': r.status || '',
    '确认人': r.owner || '',
    '备注': r.note || '',
    '日期': stamp(r.createdAt).slice(0, 10),
    '时间': stamp(r.createdAt).slice(11)
  };
  var parts = String(tpl == null || tpl === '' ? DEFAULT_TEMPLATE : tpl).split(/[，,]/);
  var kept = [];
  parts.forEach(function (part) {
    var hasPh = part.indexOf('{') >= 0;
    var filled = false;
    var text = part.replace(/\{([^}]*)\}/g, function (m, k) {
      if (!Object.prototype.hasOwnProperty.call(map, k)) { filled = true; return m; }
      if (String(map[k]).trim() !== '') filled = true;
      return map[k];
    });
    text = text.replace(/\s+/g, ' ').trim();
    if (!text) return;
    if (hasPh && !filled) return;
    kept.push(text);
  });
  return kept.join('，');
}

function renderBrief(list, tpl) {
  return (list || []).map(function (r) { return renderBriefOne(r, tpl); })
    .filter(Boolean).join('\n');
}

/** 快捷句按钮上显示的文字 = 它填进「说明 + 确认人」之后拼出来的样子 */
function presetLabel(p) {
  return String(p.desc || '') + (p.owner ? '，' + p.owner + '确认了' : '');
}

function filterRecords(list, opt) {
  var kw = ((opt && opt.keyword) || '').trim().toLowerCase();
  var status = (opt && opt.status) || '全部';
  var type = (opt && opt.type) || '全部';
  return list.filter(function (r) {
    if (status !== '全部' && r.status !== status) return false;
    if (type !== '全部' && r.type !== type) return false;
    if (!kw) return true;
    var hay = [codesText(r.codes), r.type, r.desc, r.status, r.owner, r.note].join(' ').toLowerCase();
    return hay.indexOf(kw) >= 0;
  });
}

function newId() {
  return 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function byNewest(a, b) { return (b.createdAt || 0) - (a.createdAt || 0); }

/** 把任意来源的数据（导入的 JSON / 老版本）整理成标准记录 */
function normalizeRecord(raw) {
  if (!raw || typeof raw !== 'object') return null;
  var codes = Array.isArray(raw.codes) ? raw.codes.map(String).filter(Boolean)
    : splitCodes(raw.code || raw.codes || '');
  var t = Number(raw.createdAt) || Date.now();
  return {
    id: (typeof raw.id === 'string' && raw.id) ? raw.id : newId(),
    codes: codes,
    type: raw.type ? String(raw.type) : '其他',
    desc: raw.desc ? String(raw.desc) : '',
    status: STATUSES.indexOf(raw.status) >= 0 ? raw.status : '待确认',
    owner: raw.owner ? String(raw.owner) : '',
    note: raw.note ? String(raw.note) : '',
    createdAt: t,
    updatedAt: Number(raw.updatedAt) || t
  };
}

function isRecord(r) { return !!r && Array.isArray(r.codes) && r.codes.length > 0; }

/* ======================= 运行时状态 ======================= */

var records = [];
var pinned = null;                 // 保存后保留的成稿（对象数组），方便马上复制
var editingId = null;
var selected = Object.create(null); // id -> true
var selectedCount = 0;
var ui = { keyword: '', status: '全部', type: '全部' };

function newRecord(codes, f) {
  var t = Date.now();
  return {
    id: newId(), codes: codes.slice(), type: f.type, desc: f.desc, status: f.status,
    owner: f.owner, note: f.note, createdAt: t, updatedAt: t
  };
}

function findRecord(id) {
  for (var i = 0; i < records.length; i++) if (records[i].id === id) return records[i];
  return null;
}

/* ======================= 小工具 ======================= */

function $(id) { return document.getElementById(id); }

function el(tag, props) {
  var node = document.createElement(tag);
  if (props) {
    Object.keys(props).forEach(function (k) {
      var v = props[k];
      if (v == null || v === false) return;
      if (k === 'class') node.className = v;
      else if (k === 'text') node.textContent = v;
      else if (k === 'dataset') Object.assign(node.dataset, v);
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') node.addEventListener(k.slice(2), v);
      else if (k in node) node[k] = v;
      else node.setAttribute(k, v);
    });
  }
  for (var i = 2; i < arguments.length; i++) {
    var kids = arguments[i];
    (Array.isArray(kids) ? kids : [kids]).forEach(function (kid) {
      if (kid == null || kid === false) return;
      node.appendChild(typeof kid === 'object' ? kid : document.createTextNode(String(kid)));
    });
  }
  return node;
}

function fillSelect(sel, items, def) {
  if (!sel) return;
  sel.textContent = '';
  items.forEach(function (v) {
    sel.appendChild(el('option', { value: v, text: v, selected: v === def }));
  });
}

var toastTimer = null;
function toast(msg, tone) {
  var box = $('toast');
  if (!box) return;
  box.textContent = msg;
  box.className = 'show' + (tone ? ' ' + tone : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { box.className = ''; }, 2600);
}

function copyText(text, okMsg) {
  if (!text) { toast('没有可复制的内容', 'warn'); return Promise.resolve(false); }
  var done = function (ok) {
    toast(ok ? (okMsg || '已复制到剪贴板') : '复制失败，请手动选中文本复制', ok ? 'ok' : 'warn');
    return ok;
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    return navigator.clipboard.writeText(text).then(function () { return done(true); },
      function () { return done(fallbackCopy(text)); });
  }
  return Promise.resolve(done(fallbackCopy(text)));
}

function fallbackCopy(text) {
  var ta = document.createElement('textarea');
  ta.value = text;
  ta.setAttribute('readonly', 'readonly');
  ta.style.position = 'fixed';
  ta.style.top = '-1000px';
  ta.style.opacity = '0';
  document.body.appendChild(ta);
  ta.select();
  ta.setSelectionRange(0, text.length);
  var ok = false;
  try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
  ta.remove();
  return ok;
}

function download(filename, text, mime) {
  var blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
  var url = URL.createObjectURL(blob);
  var a = el('a', { href: url, download: filename });
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
}

/* ======================= 存储 ======================= */

function loadStore() {
  var list = [];
  try {
    var raw = localStorage.getItem(STORE_KEY);
    if (raw) {
      var data = JSON.parse(raw);
      var arr = Array.isArray(data) ? data : (data && data.records);
      if (Array.isArray(arr)) {
        list = arr.map(normalizeRecord).filter(isRecord);
      }
    }
  } catch (e) {
    console.warn('[记录工作台] 读取本机数据失败：', e);
  }
  return list;
}

function persist() {
  try {
    localStorage.setItem(STORE_KEY, JSON.stringify({ v: 1, savedAt: Date.now(), records: records }));
    flashSaved();
    return true;
  } catch (e) {
    toast('保存失败：浏览器存储不可用或已满', 'bad');
    return false;
  }
}

function flashSaved() {
  var t = new Date();
  var st = $('saveState'), tx = $('saveText');
  if (!st || !tx) return;
  tx.textContent = '已保存 ' + pad2(t.getHours()) + ':' + pad2(t.getMinutes()) + ':' + pad2(t.getSeconds());
  st.classList.add('flash');
  setTimeout(function () { st.classList.remove('flash'); }, 700);
}

function getTemplate() {
  try {
    var v = localStorage.getItem(TEMPLATE_KEY);
    return v && v.trim() ? v : DEFAULT_TEMPLATE;
  } catch (e) { return DEFAULT_TEMPLATE; }
}

function saveTemplate(v) {
  try { localStorage.setItem(TEMPLATE_KEY, v); } catch (e) { /* 忽略 */ }
}

/* ---------- 快捷句的存取 ---------- */

var customPresets = [];

function allPresets() {
  return BUILTIN_PRESETS.map(function (p) {
    return { desc: p.desc, owner: p.owner, custom: false };
  }).concat(customPresets.map(function (p) {
    return { desc: p.desc, owner: p.owner, custom: true };
  }));
}

function loadPresets() {
  var list = [];
  try {
    var raw = localStorage.getItem(PRESET_KEY);
    var arr = raw ? JSON.parse(raw) : null;
    if (Array.isArray(arr)) {
      list = arr.map(function (p) {
        return { desc: String((p && p.desc) || '').trim(), owner: String((p && p.owner) || '').trim() };
      }).filter(function (p) { return p.desc; });
    }
  } catch (e) { list = []; }
  customPresets = list;
}

function saveCustomPresets() {
  try { localStorage.setItem(PRESET_KEY, JSON.stringify(customPresets)); }
  catch (e) { toast('快捷句保存失败（浏览器存储不可用）', 'warn'); }
}

function currentForm() {
  return {
    code: $('fCode').value, type: $('fType').value, status: $('fStatus').value,
    desc: $('fDesc').value, owner: $('fOwner').value, note: $('fNote').value,
    split: $('fSplit').checked
  };
}

function saveDraft() {
  try { localStorage.setItem(DRAFT_KEY, JSON.stringify({ editingId: editingId, form: currentForm() })); }
  catch (e) { /* 忽略：草稿存不下不影响主流程 */ }
}

function clearDraft() {
  try { localStorage.removeItem(DRAFT_KEY); } catch (e) { /* 忽略 */ }
}

function restoreDraft() {
  var d = null;
  try {
    var raw = localStorage.getItem(DRAFT_KEY);
    if (raw) d = JSON.parse(raw);
  } catch (e) { d = null; }
  if (!d || !d.form) return;
  var f = d.form;
  $('fCode').value = f.code || '';
  $('fDesc').value = f.desc || '';
  $('fOwner').value = f.owner || '';
  $('fNote').value = f.note || '';
  $('fType').value = TYPES.indexOf(f.type) >= 0 ? f.type : TYPES[0];
  $('fStatus').value = STATUSES.indexOf(f.status) >= 0 ? f.status : STATUSES[0];
  $('fSplit').checked = f.split !== false;
  if (d.editingId && findRecord(d.editingId)) enterEdit(d.editingId, true);
}

/* ======================= 渲染 ======================= */

function renderAll() {
  pruneSelection();
  renderCount();
  renderChips();
  renderPresets();
  renderList();
  updatePreview();
}

function pruneSelection() {
  Object.keys(selected).forEach(function (id) {
    if (!findRecord(id)) delete selected[id];
  });
  recountSelection();
}

function recountSelection() {
  selectedCount = Object.keys(selected).length;
}

function renderCount() {
  var hint = $('countHint');
  if (hint) hint.textContent = records.length ? '共 ' + records.length + ' 条记录' : '';
}

function renderChips() {
  var box = $('statusChips');
  if (!box) return;
  box.textContent = '';
  var scoped = filterRecords(records, { keyword: ui.keyword, type: ui.type, status: '全部' });
  var counts = { '全部': scoped.length };
  STATUSES.forEach(function (s) {
    counts[s] = scoped.filter(function (r) { return r.status === s; }).length;
  });
  ['全部'].concat(STATUSES).forEach(function (s) {
    box.appendChild(el('button', {
      class: 'chip' + (ui.status === s ? ' active' : '') + (s !== '全部' && !counts[s] ? ' zero' : ''),
      text: s + ' ' + (counts[s] || 0),
      onclick: function () { ui.status = s; renderAll(); }
    }));
  });
}

function renderList() {
  var box = $('list');
  if (!box) return;
  box.textContent = '';
  var list = filterRecords(records.slice().sort(byNewest), ui);

  var hint = $('listHint');
  if (hint) hint.textContent = records.length ? list.length + ' / ' + records.length + ' 条' : '';

  if (!list.length) {
    box.appendChild(el('div', {
      class: 'empty',
      text: records.length ? '没有匹配的记录，换个关键词或点「全部」' : '还没有记录，左边填一条试试'
    }));
    renderBulkBar();
    return;
  }

  var picked = list.filter(function (r) { return selected[r.id]; });
  var head = el('label', { class: 'list-head' },
    el('input', {
      type: 'checkbox',
      checked: picked.length === list.length,
      onchange: function (e) {
        list.forEach(function (r) {
          if (e.target.checked) selected[r.id] = true; else delete selected[r.id];
        });
        recountSelection();
        renderList();
      }
    }),
    el('span', { text: '全选当前 ' + list.length + ' 条' })
  );
  var headBox = head.firstChild;
  if (picked.length && picked.length < list.length) headBox.indeterminate = true;
  box.appendChild(head);

  list.forEach(function (r) { box.appendChild(renderItem(r)); });
  renderBulkBar();
}

function renderItem(r) {
  var pick = el('input', {
    type: 'checkbox', class: 'pick', checked: !!selected[r.id],
    onchange: function (e) {
      if (e.target.checked) selected[r.id] = true; else delete selected[r.id];
      recountSelection();
      renderList();
    }
  });

  var main = el('div', { class: 'item-main' },
    el('div', { class: 'item-top' },
      el('span', { class: 'code', text: codesText(r.codes) }),
      el('span', { class: 'badge ' + (STATUS_TONE[r.status] || ''), text: r.status || '' }),
      r.type ? el('span', { class: 'tag', text: r.type }) : null
    )
  );
  if (r.desc) main.appendChild(el('div', { class: 'item-desc', text: r.desc }));

  var meta = [];
  if (r.owner) meta.push('确认/处理人：' + r.owner);
  meta.push('创建 ' + stamp(r.createdAt));
  if (r.updatedAt - r.createdAt > 60000) meta.push('更新 ' + stamp(r.updatedAt));
  if (r.note) meta.push('备注：' + r.note);
  main.appendChild(el('div', { class: 'item-meta', text: meta.join(' · ') }));

  var statusSel = el('select', { class: 'mini', title: '改状态' },
    STATUSES.map(function (s) { return el('option', { value: s, text: s, selected: s === r.status }); })
  );
  statusSel.addEventListener('change', function () { setStatus(r.id, statusSel.value); });

  var side = el('div', { class: 'item-side' },
    statusSel,
    el('button', { class: 'btn xs', text: '复制', title: '复制这一条（一句话）', onclick: function () { copyText(renderBrief([r], getTemplate()), '已复制'); } }),
    el('button', { class: 'btn xs', text: '编辑', title: '载入表单修改', onclick: function () { enterEdit(r.id, false); } }),
    el('button', { class: 'btn xs danger', text: '删除', title: '删除这条', onclick: function () { removeRecord(r.id); } })
  );

  return el('div', { class: 'item' + (selected[r.id] ? ' picked' : ''), dataset: { id: r.id } }, pick, main, side);
}

function renderBulkBar() {
  var bar = $('bulkbar');
  if (!bar) return;
  bar.hidden = selectedCount === 0;
  var c = $('bulkCount');
  if (c) c.textContent = '已选 ' + selectedCount + ' 条';
}

/** 当前要成稿的对象：优先用刚保存的那批，否则用表单里的内容 */
function composeSource() {
  if (pinned && pinned.length) return pinned;
  var f = currentForm();
  var codes = splitCodes(f.code);
  if (!codes.length) return null;
  var base = {
    type: f.type, desc: f.desc.trim(), status: f.status,
    owner: f.owner.trim(), note: f.note.trim(), createdAt: Date.now()
  };
  if (f.split) {
    return codes.map(function (c) { return Object.assign({}, base, { codes: [c] }); });
  }
  return [Object.assign({}, base, { codes: codes })];
}

/** 一句话成稿用的对象（已按内容合并编号） */
function briefItems() {
  var items = composeSource();
  return items && items.length ? mergeForBrief(items) : null;
}

function updatePreview() {
  var box = $('preview');
  if (!box) return;
  var items = briefItems();
  if (!items) {
    box.textContent = '填好编号后，这里会显示点「复制」之后粘出去的那句话。';
    box.classList.add('empty');
    return;
  }
  box.textContent = renderBrief(items, getTemplate());
  box.classList.remove('empty');
}

/* ======================= 操作 ======================= */

function onFormInput() {
  pinned = null;
  updatePreview();
  renderPresets();
  saveDraft();
}

/* ---------- 快捷句 ---------- */

function renderPresets() {
  var box = $('presetBox');
  if (!box) return;
  box.textContent = '';
  var curDesc = $('fDesc').value.trim();
  var curOwner = $('fOwner').value.trim();

  allPresets().forEach(function (p) {
    var active = p.desc === curDesc && p.owner === curOwner;
    var tip = '点一下自动填好：说明「' + p.desc + '」' +
      (p.owner ? '、确认人「' + p.owner + '」' : '（确认人留空）');
    box.appendChild(el('div', { class: 'preset' + (active ? ' active' : '') },
      el('button', {
        class: 'preset-text', type: 'button', text: presetLabel(p), title: tip,
        onclick: function () { applyPreset(p); }
      }),
      p.custom ? el('button', {
        class: 'preset-del', type: 'button', text: '✕', title: '删除这个快捷句',
        onclick: function () { removePreset(p); }
      }) : null
    ));
  });
}

function applyPreset(p) {
  $('fDesc').value = p.desc;
  $('fOwner').value = p.owner;
  onFormInput();
  if (!splitCodes($('fCode').value).length) {
    $('fCode').focus();
    toast('已填好这句话，再填编号就能复制', '');
  } else {
    toast('已填好，点「复制」即可', 'ok');
  }
}

function removePreset(p) {
  customPresets = customPresets.filter(function (x) {
    return !(x.desc === p.desc && x.owner === p.owner);
  });
  saveCustomPresets();
  renderPresets();
  toast('已删除这个快捷句', '');
}

function saveCurrentPreset() {
  var desc = $('fDesc').value.trim();
  var owner = $('fOwner').value.trim();
  if (!desc) {
    toast('先在「说明 / 诉求」里写好这句话，再存快捷句', 'warn');
    $('fDesc').focus();
    return;
  }
  var exists = allPresets().some(function (p) { return p.desc === desc && p.owner === owner; });
  if (exists) { toast('这个快捷句已经有了', 'warn'); return; }
  customPresets.push({ desc: desc, owner: owner });
  saveCustomPresets();
  renderPresets();
  toast('已存成快捷句（存在本机，下次打开还在）', 'ok');
}

function resetFormFields() {
  editingId = null;
  $('fCode').value = '';
  $('fDesc').value = '';
  $('fOwner').value = '';
  $('fNote').value = '';
  $('fType').value = TYPES[0];
  $('fStatus').value = STATUSES[0];
  $('fSplit').checked = true;
  $('formTitle').textContent = '新增记录';
  $('btnSave').textContent = '保存记录（Ctrl + Enter）';
  $('btnCancelEdit').hidden = true;
  clearDraft();
}

function enterEdit(id, silent) {
  var r = findRecord(id);
  if (!r) return;
  editingId = id;
  pinned = null;
  $('fCode').value = codesText(r.codes).split('、').join(', ');
  $('fType').value = TYPES.indexOf(r.type) >= 0 ? r.type : '其他';
  $('fStatus').value = STATUSES.indexOf(r.status) >= 0 ? r.status : STATUSES[0];
  $('fDesc').value = r.desc || '';
  $('fOwner').value = r.owner || '';
  $('fNote').value = r.note || '';
  $('fSplit').checked = false;
  $('formTitle').textContent = '编辑记录';
  $('btnSave').textContent = '保存修改（Ctrl + Enter）';
  $('btnCancelEdit').hidden = false;
  if (!silent) {
    window.scrollTo({ top: 0, behavior: 'smooth' });
    $('fCode').focus();
    updatePreview();
    saveDraft();
    toast('已载入表单，改完点「保存修改」', '');
  }
}

function saveForm() {
  var f = currentForm();
  var codes = splitCodes(f.code);
  if (!codes.length) { toast('请先填编号', 'warn'); $('fCode').focus(); return; }

  var fields = {
    type: f.type, desc: f.desc.trim(), status: f.status,
    owner: f.owner.trim(), note: f.note.trim()
  };
  var created;

  if (editingId) {
    var r = findRecord(editingId);
    if (!r) { resetFormFields(); renderAll(); toast('这条记录已经不在了', 'warn'); return; }
    r.codes = codes;
    Object.assign(r, fields);
    r.updatedAt = Date.now();
    created = [r];
    toast('已更新 1 条', 'ok');
  } else {
    created = f.split
      ? codes.map(function (c) { return newRecord([c], fields); })
      : [newRecord(codes, fields)];
    records = records.concat(created);
    toast('已保存 ' + created.length + ' 条', 'ok');
  }

  resetFormFields();
  pinned = created;
  persist();
  renderAll();
}

function setStatus(id, status) {
  var r = findRecord(id);
  if (!r || r.status === status) return;
  r.status = status;
  r.updatedAt = Date.now();
  persist();
  renderChips();
  renderList();
  toast('状态已改为「' + status + '」', 'ok');
}

function removeRecord(id) {
  var r = findRecord(id);
  if (!r) return;
  if (!window.confirm('删除这条记录？\n\n' + codesText(r.codes) + '\n（删了没法撤销）')) return;
  records = records.filter(function (x) { return x.id !== id; });
  delete selected[id];
  recountSelection();
  persist();
  renderAll();
  toast('已删除 1 条', '');
}

function bulkSetStatus() {
  var status = $('bulkStatus').value;
  var n = 0;
  records.forEach(function (r) {
    if (selected[r.id]) { r.status = status; r.updatedAt = Date.now(); n++; }
  });
  selected = Object.create(null);
  recountSelection();
  persist();
  renderAll();
  toast('已把 ' + n + ' 条改为「' + status + '」', 'ok');
}

function bulkDelete() {
  if (!selectedCount) return;
  if (!window.confirm('删除选中的 ' + selectedCount + ' 条记录？（删了没法撤销）')) return;
  var n = selectedCount;
  records = records.filter(function (r) { return !selected[r.id]; });
  selected = Object.create(null);
  recountSelection();
  persist();
  renderAll();
  toast('已删除 ' + n + ' 条', '');
}

function exportCSV() {
  var list = filterRecords(records.slice().sort(byNewest), ui);
  if (!list.length) { toast('没有可导出的记录', 'warn'); return; }
  var rows = [['编号', '事项', '说明', '状态', '确认人', '备注', '创建时间', '更新时间']];
  list.forEach(function (r) {
    rows.push([codesText(r.codes), r.type, r.desc, r.status, r.owner, r.note, stamp(r.createdAt), stamp(r.updatedAt)]);
  });
  download('记录_' + fileStamp() + '.csv', toCSV(rows), 'text/csv;charset=utf-8');
  toast('已导出 ' + list.length + ' 条（CSV，可用 Excel 打开）', 'ok');
}

function exportJSON() {
  if (!records.length) { toast('还没有记录', 'warn'); return; }
  download('记录备份_' + fileStamp() + '.json',
    JSON.stringify({ v: 1, exportedAt: new Date().toISOString(), records: records }, null, 2),
    'application/json;charset=utf-8');
  toast('已导出备份（' + records.length + ' 条）', 'ok');
}

function importJSON(file) {
  var reader = new FileReader();
  reader.onload = function () {
    var data;
    try { data = JSON.parse(String(reader.result)); }
    catch (e) { toast('这个文件不是合法的 JSON', 'bad'); return; }
    var arr = Array.isArray(data) ? data : (data && Array.isArray(data.records) ? data.records : null);
    if (!arr) { toast('文件里没有 records 数组', 'bad'); return; }

    var have = Object.create(null);
    records.forEach(function (r) { have[r.id] = 1; });
    var added = 0, skipped = 0;
    arr.forEach(function (raw) {
      var r = normalizeRecord(raw);
      if (!isRecord(r) || have[r.id]) { skipped++; return; }
      have[r.id] = 1;
      records.push(r);
      added++;
    });
    persist();
    renderAll();
    toast('导入完成：新增 ' + added + ' 条' + (skipped ? '，跳过 ' + skipped + ' 条（重复或无效）' : ''), added ? 'ok' : 'warn');
  };
  reader.onerror = function () { toast('读取文件失败', 'bad'); };
  reader.readAsText(file);
}

function clearAll() {
  if (!records.length) { toast('本来就是空的', ''); return; }
  if (!window.confirm('清空全部 ' + records.length + ' 条记录？\n\n建议先「导出备份 JSON」。清空后无法恢复。')) return;
  records = [];
  selected = Object.create(null);
  recountSelection();
  pinned = null;
  persist();
  renderAll();
  toast('已清空', '');
}

/* ======================= 事件绑定 ======================= */

function bind() {
  ['fCode', 'fType', 'fStatus', 'fDesc', 'fOwner', 'fNote', 'fSplit'].forEach(function (id) {
    var n = $(id);
    if (!n) return;
    n.addEventListener('input', onFormInput);
    n.addEventListener('change', onFormInput);
  });

  $('btnSave').addEventListener('click', saveForm);
  $('btnReset').addEventListener('click', function () {
    resetFormFields();
    pinned = null;
    updatePreview();
    toast('已重置输入框', '');
  });
  $('btnCancelEdit').addEventListener('click', function () {
    resetFormFields();
    pinned = null;
    updatePreview();
    toast('已退出编辑', '');
  });

  /* 主按钮：复制那一句话 */
  $('btnCopyBrief').addEventListener('click', function () {
    var items = briefItems();
    if (!items) { toast('先填编号', 'warn'); $('fCode').focus(); return; }
    var n = items.reduce(function (a, r) { return a + r.codes.length; }, 0);
    copyText(renderBrief(items, getTemplate()), '已复制：' + n + ' 个编号');
  });

  /* 备用格式：表格行 / 详细清单 */
  $('btnCopyLine').addEventListener('click', function () {
    var items = composeSource();
    if (!items) { toast('先填编号', 'warn'); return; }
    copyText(composeSimple(items), '已复制表格行（可直接粘 Excel）');
  });
  $('btnCopyList').addEventListener('click', function () {
    var items = composeSource();
    if (!items) { toast('先填编号', 'warn'); return; }
    copyText(composeList(items), '已复制详细清单');
  });

  /* 复制格式模板 */
  var tplInput = $('fTpl');
  tplInput.value = getTemplate();
  tplInput.addEventListener('input', function () {
    saveTemplate(tplInput.value);
    updatePreview();
  });
  $('btnTplReset').addEventListener('click', function () {
    tplInput.value = DEFAULT_TEMPLATE;
    saveTemplate(DEFAULT_TEMPLATE);
    updatePreview();
    toast('已恢复默认格式', '');
  });

  $('btnSavePreset').addEventListener('click', saveCurrentPreset);

  $('btnCopyFiltered').addEventListener('click', function () {
    var list = filterRecords(records.slice().sort(byNewest), ui);
    if (!list.length) { toast('当前筛选结果里没有记录', 'warn'); return; }
    copyText(composeSimple(list), '已复制 ' + list.length + ' 条表格行');
  });

  var search = $('search');
  search.addEventListener('input', function () {
    ui.keyword = search.value;
    renderChips();
    renderList();
  });
  search.addEventListener('keydown', function (e) { if (e.key === 'Escape') { search.value = ''; ui.keyword = ''; renderChips(); renderList(); } });

  $('typeFilter').addEventListener('change', function (e) {
    ui.type = e.target.value;
    renderChips();
    renderList();
  });

  $('btnBulkStatus').addEventListener('click', bulkSetStatus);
  $('btnBulkDelete').addEventListener('click', bulkDelete);
  $('btnClearSel').addEventListener('click', function () {
    selected = Object.create(null);
    recountSelection();
    renderList();
  });

  $('btnExportCsv').addEventListener('click', exportCSV);
  $('btnExportJson').addEventListener('click', exportJSON);
  $('btnImport').addEventListener('click', function () { $('fileInput').click(); });
  $('fileInput').addEventListener('change', function (e) {
    var file = e.target.files && e.target.files[0];
    if (file) importJSON(file);
    e.target.value = '';
  });
  $('btnClearAll').addEventListener('click', clearAll);

  document.addEventListener('keydown', function (e) {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); saveForm(); }
  });

  window.addEventListener('beforeunload', function () { saveDraft(); });
}

/* ======================= 启动 ======================= */

function init() {
  fillSelect($('fType'), TYPES, TYPES[0]);
  fillSelect($('fStatus'), STATUSES, STATUSES[0]);
  fillSelect($('bulkStatus'), STATUSES, STATUSES[1]);
  fillSelect($('typeFilter'), ['全部'].concat(TYPES), '全部');

  records = loadStore();
  loadPresets();
  bind();
  restoreDraft();
  renderAll();

  var tx = $('saveText');
  if (tx) tx.textContent = records.length ? '已保存到本机' : '数据存在本机';
}

if (typeof document !== 'undefined') {
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
}

if (typeof window !== 'undefined') {
  window.__RW = {
    STATUSES: STATUSES,
    TYPES: TYPES,
    DEFAULT_TEMPLATE: DEFAULT_TEMPLATE,
    splitCodes: splitCodes,
    composeSimple: composeSimple,
    composeList: composeList,
    renderBrief: renderBrief,
    mergeForBrief: mergeForBrief,
    presetLabel: presetLabel,
    allPresets: allPresets,
    toCSV: toCSV,
    filterRecords: filterRecords,
    getTemplate: getTemplate,
    get records() { return records; },
    get state() { return { ui: ui, editingId: editingId, selectedCount: selectedCount }; }
  };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    splitCodes: splitCodes, codesText: codesText, composeSimple: composeSimple, composeList: composeList,
    renderBrief: renderBrief, renderBriefOne: renderBriefOne, mergeForBrief: mergeForBrief,
    toCSV: toCSV, csvCell: csvCell, filterRecords: filterRecords, normalizeRecord: normalizeRecord,
    isRecord: isRecord, stamp: stamp, fileStamp: fileStamp, DEFAULT_TEMPLATE: DEFAULT_TEMPLATE,
    presetLabel: presetLabel, BUILTIN_PRESETS: BUILTIN_PRESETS
  };
}
