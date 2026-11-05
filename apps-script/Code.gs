// Made with Claude (Anthropic) · © Homes in Georgia · Created by Irakli Kvirkvelia
// Daily Rent · Gudauri Marshall: связь сайта с Google Таблицей.
// В репозитории лежит шаблон без ключей и без данных. Рабочая копия хранится только в вашей таблице.

var KEY_EDIT = '__KEY_EDIT__';  // ссылка для команды: можно менять
var KEY_VIEW = '__KEY_VIEW__';  // ссылка для собственника: только просмотр
var SEED = /*SEED*/{ items: [], transfers: [] }/*END*/;

var T_ITEMS = 'Товары', T_BUDGET = 'Бюджет', T_LOG = 'История', T_RECEIPTS = 'Чеки';
var ITEM_COLS = [['id','ID'],['order','Порядок'],['name','Название'],['cat','Группа'],['room','Комната'],['status','Статус'],['price','Цена'],['currency','Валюта'],['paidAt','Дата оплаты'],['link','Ссылка'],['unsure','Под вопросом'],['photos','Фото (ID файлов)'],['note','Заметка'],['replaces','Вместо (ID)'],['replacedBy','Заменено на (ID)'],['updated','Изменено']];
var BUDGET_COLS = [['id','ID'],['usd','Сумма, $'],['rate','Курс, ₾ за $'],['date','Дата перевода'],['note','Заметка']];
var RECEIPT_COLS = [['id','ID'],['date','Дата'],['shop','Магазин'],['amount','Сумма'],['currency','Валюта'],['photos','Фото (ID файлов)'],['items','Товары (ID)'],['note','Заметка'],['updated','Изменено']];
var ROOMS = ['Спальня','Кухня','Ванная','Прихожая','Хозяйство','Другое'];
var LOG_COLS = [['at','Когда'],['by','Кто'],['text','Что изменилось']];
var CATS = {mebel:'Мебель',tech:'Электротовары',light:'Освещение',textile:'Текстиль и ковры',decor:'Декор',kitchen:'Посуда и кухня',home:'Ванная и быт',other:'Другое'};
var STATUSES = {need:'Нужно купить',paid:'Оплачено',replaced:'Заменено',removed:'Убрано'};
var LOG_LIMIT = 300;

function doGet() { return ContentService.createTextOutput('Daily Rent · Gudauri Marshall: скрипт работает.'); }

function doPost(e) {
  try {
    var q = JSON.parse(e.postData.contents);
    var role = q.key === KEY_EDIT ? 'edit' : (q.key === KEY_VIEW ? 'view' : null);
    if (!role || String(q.key).indexOf('__') === 0) return out({ ok: false, code: 'bad_key' });
    if (q.op === 'read') {
      if (!ss_().getSheetByName(T_RECEIPTS)) withLock_(function () { ensure_(); });
      if (needsIds_()) withLock_(function () { fixIds_(); });
      return out({ ok: true, role: role, state: state_(role) });
    }
    if (role !== 'edit') return out({ ok: false, code: 'forbidden' });
    if (q.op === 'batch') { withLock_(function () { ensure_(); fixIds_(); applyBatch_(q.ops || []); }); return out({ ok: true, role: role, state: state_(role) }); }
    if (q.op === 'upload') return out(upload_(q));
    return out({ ok: false, code: 'bad_request' });
  } catch (err) {
    return out({ ok: false, code: 'server_error', message: String(err && err.message || err).slice(0, 200) });
  }
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function withLock_(fn) { var l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }

/* ---------- Листы ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function ensure_() {
  var ss = ss_();
  if (!ss.getSheetByName(T_ITEMS)) {
    var first = ss.getSheets()[0];
    first.setName(T_ITEMS);
    var b = ss.insertSheet(T_BUDGET), l = ss.insertSheet(T_LOG);
    header_(first, ITEM_COLS); header_(b, BUDGET_COLS); header_(l, LOG_COLS);
    writeTable_(first, ITEM_COLS, SEED.items.map(itemToRow_));
    writeTable_(b, BUDGET_COLS, SEED.transfers.map(cleanTransfer_));
    first.setColumnWidth(3, 340); first.setColumnWidth(4, 150); first.setColumnWidth(6, 120);
  }
  if (!ss.getSheetByName(T_RECEIPTS)) {
    header_(ss.insertSheet(T_RECEIPTS), RECEIPT_COLS);
    dropdowns_();
  }
}
function dropdowns_() {
  var sh = ss_().getSheetByName(T_ITEMS);
  var put = function (col, list) {
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build();
    sh.getRange(2, col, 999, 1).setDataValidation(rule);
  };
  put(4, values_(CATS)); put(5, ROOMS); put(6, values_(STATUSES)); put(8, ['GEL', 'USD']); put(11, ['Да']);
}
function values_(m) { var r = []; for (var k in m) r.push(m[k]); return r; }
// Строка, добавленная вручную: достаточно названия, остальное заполнится само.
function needsIds_() {
  return readTable_(T_ITEMS, ITEM_COLS).some(function (r) { return r.id === '' && r.name !== ''; });
}
function fixIds_() {
  var sh = ss_().getSheetByName(T_ITEMS), rows = readTable_(T_ITEMS, ITEM_COLS);
  var max = rows.reduce(function (m, r) { return Math.max(m, num_(r.order)); }, 0);
  rows.forEach(function (r, i) {
    if (r.id !== '' || r.name === '') return;
    var row = i + 2;
    sh.getRange(row, 1).setValue(cell_('h' + Utilities.getUuid().slice(0, 8)));
    if (!num_(r.order)) sh.getRange(row, 2).setValue(++max);
    if (r.cat === '') sh.getRange(row, 4).setValue(cell_(CATS.other));
    if (r.room === '') sh.getRange(row, 5).setValue(cell_('Другое'));
    if (r.status === '') sh.getRange(row, 6).setValue(cell_(STATUSES.need));
    if (r.currency === '') sh.getRange(row, 8).setValue(cell_('GEL'));
  });
}
function num_(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(/\s/g, '').replace(',', '.'));
  return isFinite(n) ? n : 0;
}
function header_(sh, cols) {
  sh.getRange(1, 1, 1, cols.length).setValues([cols.map(function (c) { return c[1]; })]).setFontWeight('bold');
  sh.setFrozenRows(1);
}
function cell_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  var s = String(v);
  return s === '' ? '' : "'" + s; // апостроф: текст остаётся текстом, а не формулой или датой
}
function writeTable_(sh, cols, rows) {
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, cols.length).clearContent();
  if (!rows.length) return;
  sh.getRange(2, 1, rows.length, cols.length).setValues(rows.map(function (o) {
    return cols.map(function (c) { return cell_(o[c[0]]); });
  }));
}
function readTable_(name, cols) {
  var sh = ss_().getSheetByName(name), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols.length).getValues().map(function (r) {
    var o = {};
    cols.forEach(function (c, i) {
      var v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      o[c[0]] = v;
    });
    return o;
  });
}
function byValue_(map, v, dflt) { for (var k in map) if (map[k] === v || k === v) return k; return dflt; }
function rowToItem_(r) {
  return { id: String(r.id), order: num_(r.order), name: String(r.name), cat: byValue_(CATS, r.cat, 'other'), room: String(r.room || 'Другое'),
    status: byValue_(STATUSES, r.status, 'need'), price: num_(r.price) > 0 ? num_(r.price) : null, currency: r.currency === 'USD' ? 'USD' : 'GEL',
    paidAt: r.paidAt ? String(r.paidAt) : null, link: String(r.link || ''), unsure: r.unsure === 'Да' || r.unsure === true,
    photos: String(r.photos || '').split(',').filter(String), note: String(r.note || ''), replaces: r.replaces ? String(r.replaces) : null, replacedBy: r.replacedBy ? String(r.replacedBy) : null, updated: String(r.updated || '') };
}
function txt_(v, n) { return String(v === null || v === undefined ? '' : v).slice(0, n || 500); }
function itemToRow_(d) {
  var link = txt_(d.link, 1000).trim(); if (link && !/^https?:\/\//i.test(link)) link = /^[\w.-]+\.[a-z]{2,}/i.test(link) ? 'https://' + link : '';
  var price = Number(d.price);
  return { id: txt_(d.id, 60), order: Number(d.order) || 0, name: txt_(d.name, 300), cat: CATS[d.cat] || CATS.other, room: txt_(d.room, 60) || 'Другое',
    status: STATUSES[d.status] || STATUSES.need, price: price > 0 ? price : '', currency: d.currency === 'USD' ? 'USD' : 'GEL', paidAt: txt_(d.paidAt, 10), link: link,
    unsure: d.unsure ? 'Да' : '', photos: (d.photos || []).filter(function (x) { return /^[\w-]{10,100}$/.test(x); }).join(','), note: txt_(d.note, 1000),
    replaces: txt_(d.replaces, 60), replacedBy: txt_(d.replacedBy, 60), updated: d.updated || '' };
}
function cleanTransfer_(t) {
  var usd = Number(t.usd), rate = Number(t.rate);
  return { id: txt_(t.id, 60), usd: usd > 0 ? usd : 0, rate: rate > 0 ? rate : '', date: txt_(t.date, 10), note: txt_(t.note, 300) };
}
function receiptToRow_(d) {
  var a = Number(d.amount);
  return { id: txt_(d.id, 60), date: txt_(d.date, 10), shop: txt_(d.shop, 120), amount: a > 0 ? a : '', currency: d.currency === 'USD' ? 'USD' : 'GEL',
    photos: (d.photos || []).filter(function (x) { return /^[\w-]{10,100}$/.test(x); }).join(','),
    items: (d.items || []).filter(function (x) { return /^[\w-]{1,60}$/.test(x); }).join(','), note: txt_(d.note, 1000), updated: d.updated || '' };
}
function rowToReceipt_(r) {
  return { id: String(r.id), date: r.date ? String(r.date) : null, shop: String(r.shop || ''), amount: num_(r.amount) > 0 ? num_(r.amount) : null, currency: r.currency === 'USD' ? 'USD' : 'GEL',
    photos: String(r.photos || '').split(',').filter(String), items: String(r.items || '').split(',').filter(String), note: String(r.note || ''), updated: String(r.updated || '') };
}
function state_(role) {
  var items = readTable_(T_ITEMS, ITEM_COLS).filter(function (r) { return r.id !== ''; }).map(rowToItem_);
  var transfers = readTable_(T_BUDGET, BUDGET_COLS).filter(function (r) { return r.id !== ''; }).map(function (t) {
    return { id: String(t.id), usd: Number(t.usd) || 0, rate: Number(t.rate) > 0 ? Number(t.rate) : null, date: t.date ? String(t.date) : null, note: String(t.note || '') };
  });
  var history = readTable_(T_LOG, LOG_COLS).filter(function (r) { return r.text !== ''; }).map(function (r) { return { at: String(r.at), by: String(r.by), text: String(r.text) }; }).reverse().slice(0, 200);
  var receipts = readTable_(T_RECEIPTS, RECEIPT_COLS).filter(function (r) { return r.id !== ''; }).map(rowToReceipt_);
  if (role !== 'edit') { // собственнику отдаём только то, что входит в отчёт
    items.forEach(function (i) { i.note = ''; i.updated = ''; });
    receipts.forEach(function (r) { r.updated = ''; });
    history = [];
  }
  return { items: items, transfers: transfers, history: history, receipts: receipts };
}

/* ---------- Запись ---------- */
function applyBatch_(ops) {
  var items = readTable_(T_ITEMS, ITEM_COLS).filter(function (r) { return r.id !== ''; }).map(rowToItem_);
  var idx = {}; items.forEach(function (i) { idx[i.id] = i; });
  var itemsDirty = false, logs = [], trash = [];
  var receipts = readTable_(T_RECEIPTS, RECEIPT_COLS).filter(function (r) { return r.id !== ''; }).map(rowToReceipt_), rDirty = false;
  var stamp = new Date().toISOString().slice(0, 16).replace('T', ' ');
  ops.forEach(function (o) {
    if (o.t === 'set' || o.t === 'update') {
      var id = txt_(o.id, 60); if (!id) return;
      var cur = idx[id];
      if (o.t === 'update' && !cur) return;
      var next = {}; var base = o.t === 'update' ? cur : {};
      for (var k in base) next[k] = base[k];
      ITEM_COLS.forEach(function (c) { if (o.data && Object.prototype.hasOwnProperty.call(o.data, c[0]) && c[0] !== 'id' && c[0] !== 'updated') next[c[0]] = o.data[c[0]]; });
      next.id = id;
      next.updated = stamp + (o.data && o.data.updatedBy ? ' · ' + txt_(o.data.updatedBy, 60) : '');
      if (!cur) items.push(next);
      else items[items.indexOf(cur)] = next;
      idx[id] = next; itemsDirty = true;
    } else if (o.t === 'delete') {
      if (idx[o.id]) { items.splice(items.indexOf(idx[o.id]), 1); delete idx[o.id]; itemsDirty = true; }
    } else if (o.t === 'budget') {
      writeTable_(ss_().getSheetByName(T_BUDGET), BUDGET_COLS, (o.transfers || []).map(cleanTransfer_).filter(function (t) { return t.usd > 0; }));
    } else if (o.t === 'log') {
      logs.push({ at: txt_(o.at, 30), by: txt_(o.by, 60), text: txt_(o.text, 500) });
    } else if (o.t === 'rset') {
      var rid = txt_(o.id, 60); if (!rid) return;
      var nr = {}; RECEIPT_COLS.forEach(function (c) { if (o.data && o.data[c[0]] !== undefined) nr[c[0]] = o.data[c[0]]; });
      nr.id = rid; nr.updated = stamp + (o.data && o.data.updatedBy ? ' · ' + txt_(o.data.updatedBy, 60) : '');
      var at = -1; receipts.forEach(function (x, i) { if (x.id === rid) at = i; });
      if (at >= 0) receipts[at] = nr; else receipts.push(nr);
      rDirty = true;
    } else if (o.t === 'rdel') {
      receipts = receipts.filter(function (x) { return x.id !== o.id; }); rDirty = true;
    } else if (o.t === 'trash') { trash.push(String(o.file)); }
  });
  if (itemsDirty) writeTable_(ss_().getSheetByName(T_ITEMS), ITEM_COLS, items.map(itemToRow_));
  if (rDirty) writeTable_(ss_().getSheetByName(T_RECEIPTS), RECEIPT_COLS, receipts.map(receiptToRow_));
  if (logs.length) appendLog_(logs);
  trash.forEach(trashFile_);
}
function appendLog_(logs) {
  var sh = ss_().getSheetByName(T_LOG), start = sh.getLastRow() + 1;
  sh.getRange(start, 1, logs.length, 3).setValues(logs.map(function (l) { return [cell_(l.at), cell_(l.by), cell_(l.text)]; }));
  var extra = sh.getLastRow() - 1 - LOG_LIMIT;
  if (extra > 0) sh.deleteRows(2, extra);
}

/* ---------- Фото ---------- */
function folder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('Daily Rent · Gudauri Marshall: фото покупок');
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}
function upload_(q) {
  var mime = String(q.mime || '');
  if (['image/jpeg', 'image/png', 'image/webp', 'image/gif'].indexOf(mime) < 0) return { ok: false, code: 'unsupported_type' };
  var data = String(q.data || '');
  if (!data || data.length > 22000000) return { ok: false, code: 'too_large' };
  var blob = Utilities.newBlob(Utilities.base64Decode(data), mime, 'photo-' + Date.now());
  var file = withLock_(function () { return folder_().createFile(blob); });
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, id: file.getId() };
}
function trashFile_(id) {
  try {
    if (!/^[\w-]{10,100}$/.test(id)) return;
    var file = DriveApp.getFileById(id), folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
    var ps = file.getParents();
    while (ps.hasNext()) if (ps.next().getId() === folderId) { file.setTrashed(true); return; } // только файлы из папки покупок
  } catch (e) {}
}
