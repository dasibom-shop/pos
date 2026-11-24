/**
 * 다시봄 POS — 구글 시트 동기화 백엔드 (Google Apps Script)
 * =====================================================================
 *  쓰는 법 (약 5분)
 *  1. Google 스프레드시트를 새로 만들고 이름을 "다시봄 POS" 로 정합니다.
 *  2. 메뉴에서 [확장 프로그램] → [Apps Script] 를 엽니다.
 *  3. 이 파일 내용 전체를 Code.gs 에 붙여넣고 저장합니다.
 *  4. [배포] → [새 배포] → 유형: 웹 앱
 *       - 설명: 다시봄 POS
 *       - 다음 사용자로 실행: 나
 *       - 액세스 권한: 모든 사용자(Anyone)      ← 폰·태블릿에서 접속하려면 필수
 *  5. 배포 후 나오는 "웹 앱 URL"(.../exec)을 POS 관리자 → 동기화 화면에 붙여넣습니다.
 *
 *  시트는 처음 요청이 들어올 때 Products / Orders / SyncLog 세 개가 자동 생성됩니다.
 *  (스프레드시트에 붙이지 않고 독립 스크립트로 만들었다면, 첫 실행 시 시트를 새로 만들어
 *   그 ID를 스크립트 속성에 저장합니다.)
 *
 *  ■ 주의
 *  - Products 시트에는 매입가(회원가)가 들어 있습니다. 이 시트는 절대 공유하지 마세요.
 *  - 액세스 권한이 "모든 사용자"이므로 URL을 아는 사람은 주문을 넣을 수 있습니다.
 *    URL은 비밀번호처럼 다루고, 필요하면 Apps Script에서 배포를 새로 만들어 URL을 교체하세요.
 * =====================================================================
 */

var CONFIG = {
  // 특정 스프레드시트에 고정하려면 ID를 넣으세요. 비워 두면 활성 시트(또는 자동 생성)를 씁니다.
  SPREADSHEET_ID: '',
  PRODUCTS_SHEET: 'Products',
  ORDERS_SHEET: 'Orders',
  LOG_SHEET: 'SyncLog',
  MAX_ORDERS: 20000
};

var PRODUCT_HEADERS = [
  'id', 'name', 'category', 'summary', 'points', 'ingredients', 'usage', 'caution',
  'image', 'price', 'cost', 'costExample', 'pv', 'stock', 'visible', 'soldout',
  'sort', 'updatedAt'
];

var ORDER_HEADERS = [
  'id', 'no', 'ts', 'datetime', 'branchId', 'branch', 'channel', 'payment',
  'customerName', 'customerPhone', 'address', 'memo',
  'itemsJson', 'itemsSummary', 'qty', 'subtotal',
  'discountType', 'discountValue', 'discountAmount', 'total',
  'costTotal', 'costItemsJson', 'status', 'statusUpdatedAt',
  'device', 'manager', 'updatedAt', 'syncedAt'
];

/* ============================ 진입점 ============================ */

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    var action = p.action || 'all';
    var out = { ok: true, app: 'dashibom-pos', version: '1.0.0', serverTime: new Date().toISOString() };

    if (action === 'ping') {
      out.counts = counts_();
    } else if (action === 'products') {
      out.products = readProducts_();
    } else if (action === 'orders') {
      out.orders = readOrders_(p.since);
    } else if (action === 'all') {
      out.products = readProducts_();
      out.orders = readOrders_(p.since);
      out.counts = counts_();
    } else {
      throw new Error('알 수 없는 action: ' + action);
    }
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var body = parseBody_(e);
    var action = body.action;
    var result;

    if (action === 'products.save' || action === 'products.replace') {
      result = saveProducts_(body.products || [], action === 'products.replace');
    } else if (action === 'orders.push' || action === 'order.upsert') {
      var list = body.orders || (body.order ? [body.order] : []);
      result = saveOrders_(list);
    } else if (action === 'order.status') {
      result = updateOrderStatus_(body.id, body.status, body.updatedAt);
    } else if (action === 'order.delete') {
      result = deleteOrder_(body.id);
    } else {
      throw new Error('알 수 없는 action: ' + action);
    }
    result.ok = true;
    result.serverTime = new Date().toISOString();
    return json_(result);
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* noop */ }
  }
}

/** CORS 프리플라이트 대응(Apps Script는 OPTIONS를 직접 처리하지 않지만, 호출되면 빈 응답) */
function doOptions() {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}

/** 시트를 미리 만들어 두고 싶을 때 편집기에서 한 번 실행 */
function setup() {
  ensureProducts_();
  ensureOrders_();
  ensureLog_();
  log_('setup', '시트 초기화');
  return 'Products / Orders / SyncLog 시트를 준비했습니다.';
}

/** 브라우저에서 URL을 열었을 때 동작 확인용 (배포 후 주소창에 ?action=ping) */
function testConnection() {
  return counts_();
}

/* ============================ 응답·요청 ============================ */

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) throw new Error('요청 본문이 비어 있습니다.');
  var text = e.postData.contents;
  try {
    return JSON.parse(text);
  } catch (err) {
    throw new Error('JSON 본문을 해석할 수 없습니다.');
  }
}

/* ============================ 스프레드시트 ============================ */

function ss_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  var active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { active = null; }
  if (active) return active;

  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e2) { /* 아래에서 새로 만든다 */ }
  }
  var created = SpreadsheetApp.create('다시봄 POS 동기화');
  props.setProperty('SPREADSHEET_ID', created.getId());
  return created;
}

function sheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    var width = Math.max(headers.length, sh.getLastColumn());
    var first = sh.getRange(1, 1, 1, width).getValues()[0];
    var same = true;
    for (var i = 0; i < headers.length; i++) {
      if (String(first[i] || '') !== headers[i]) { same = false; break; }
    }
    if (!same) {
      sh.insertRowBefore(1);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function ensureProducts_() { return sheet_(CONFIG.PRODUCTS_SHEET, PRODUCT_HEADERS); }
function ensureOrders_() { return sheet_(CONFIG.ORDERS_SHEET, ORDER_HEADERS); }
function ensureLog_() { return sheet_(CONFIG.LOG_SHEET, ['at', 'action', 'detail']); }

function readRows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var width = sh.getLastColumn();
  var values = sh.getRange(2, 1, last - 1, width).getValues();
  var head = sh.getRange(1, 1, 1, width).getValues()[0].map(String);
  return values.map(function (row) {
    var o = {};
    for (var i = 0; i < head.length; i++) if (head[i]) o[head[i]] = row[i];
    return o;
  }).filter(function (o) { return String(o.id || '').trim() !== ''; });
}

function indexById_(sh) {
  var last = sh.getLastRow();
  var map = {};
  if (last < 2) return map;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i][0] || '').trim();
    if (id) map[id] = i + 2; // 실제 시트 행 번호
  }
  return map;
}

/* ============================ Products ============================ */

function readProducts_() {
  var sh = ensureProducts_();
  return readRows_(sh).map(productFromRow_);
}

function productFromRow_(r) {
  return {
    id: String(r.id || ''),
    name: String(r.name || ''),
    category: String(r.category || ''),
    summary: String(r.summary || ''),
    points: String(r.points || '').split('\n').filter(function (x) { return String(x).trim() !== ''; }),
    ingredients: String(r.ingredients || ''),
    usage: String(r.usage || ''),
    caution: String(r.caution || ''),
    image: String(r.image || ''),
    price: num_(r.price),
    cost: num_(r.cost),
    costExample: isY_(r.costExample),
    pv: r.pv === '' || r.pv === null ? '' : String(r.pv),
    stock: num_(r.stock),
    visible: !isN_(r.visible),
    soldout: isY_(r.soldout),
    sort: num_(r.sort),
    updatedAt: num_(r.updatedAt)
  };
}

function productToRow_(p, existing) {
  var prev = existing || {};
  var row = {};
  row.id = String(p.id || prev.id || '');
  row.name = p.name != null ? String(p.name) : String(prev.name || '');
  row.category = p.category != null ? String(p.category) : String(prev.category || '');
  row.summary = p.summary != null ? String(p.summary) : String(prev.summary || '');
  row.points = p.points != null
    ? (Array.isArray(p.points) ? p.points.join('\n') : String(p.points))
    : String(prev.points || '');
  row.ingredients = p.ingredients != null ? String(p.ingredients) : String(prev.ingredients || '');
  row.usage = p.usage != null ? String(p.usage) : String(prev.usage || '');
  row.caution = p.caution != null ? String(p.caution) : String(prev.caution || '');
  row.image = p.image != null ? String(p.image) : String(prev.image || '');
  row.price = p.price != null && p.price !== '' ? num_(p.price) : num_(prev.price);
  row.cost = p.cost != null && p.cost !== '' ? num_(p.cost) : num_(prev.cost);
  row.costExample = p.costExample != null
    ? (isY_(p.costExample) ? 'Y' : 'N')
    : (prev.costExample != null ? String(prev.costExample) : 'N');
  row.pv = p.pv != null ? String(p.pv) : String(prev.pv || '');
  row.stock = p.stock != null && p.stock !== '' ? num_(p.stock) : num_(prev.stock);
  row.visible = p.visible != null ? (p.visible === false || p.visible === 'N' ? 'N' : 'Y')
    : (prev.visible != null ? String(prev.visible) : 'Y');
  row.soldout = p.soldout != null ? (isY_(p.soldout) ? 'Y' : 'N')
    : (prev.soldout != null ? String(prev.soldout) : 'N');
  row.sort = p.sort != null && p.sort !== '' ? num_(p.sort) : num_(prev.sort);
  row.updatedAt = num_(p.updatedAt) || Date.now();
  return row;
}

function saveProducts_(products, replaceMissing) {
  var sh = ensureProducts_();
  var map = indexById_(sh);
  var rows = [];
  var created = 0, updated = 0, skipped = 0;

  products.forEach(function (p) {
    if (!p || !p.id) { skipped++; return; }
    var rowNum = map[String(p.id)];
    if (rowNum) {
      var existing = rowToObject_(sh, rowNum, PRODUCT_HEADERS);
      var merged = productToRow_(p, existing);
      writeRow_(sh, rowNum, merged, PRODUCT_HEADERS);
      updated++;
    } else {
      var fresh = productToRow_(p, null);
      rows.push(fresh);
      created++;
    }
  });

  if (rows.length) {
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, PRODUCT_HEADERS.length)
      .setValues(rows.map(function (r) { return headersToValues_(r, PRODUCT_HEADERS); }));
  }
  if (replaceMissing) {
    // 시트에만 있고 전달되지 않은 품목은 숨김 처리(삭제하지 않음)
    var sent = {};
    products.forEach(function (p) { if (p && p.id) sent[String(p.id)] = true; });
    readRows_(sh).forEach(function (r) {
      if (!sent[String(r.id)]) {
        var rowNum2 = map[String(r.id)];
        if (rowNum2) {
          var ex = rowToObject_(sh, rowNum2, PRODUCT_HEADERS);
          ex.visible = 'N';
          writeRow_(sh, rowNum2, ex, PRODUCT_HEADERS);
        }
      }
    });
  }
  log_('products.save', 'created=' + created + ' updated=' + updated + ' skipped=' + skipped);
  return { updated: updated, created: created, skipped: skipped, total: readRows_(sh).length };
}

/* ============================ Orders ============================ */

function readOrders_(since) {
  var sh = ensureOrders_();
  var rows = readRows_(sh);
  var sinceNum = num_(since);
  return rows.map(orderFromRow_).filter(function (o) {
    if (!sinceNum) return true;
    return num_(o.updatedAt) >= sinceNum;
  });
}

function orderFromRow_(r) {
  var items = [];
  try { items = JSON.parse(String(r.itemsJson || '[]')); } catch (e) { items = []; }
  var costItems = null;
  try { costItems = JSON.parse(String(r.costItemsJson || 'null')); } catch (e2) { costItems = null; }
  return {
    id: String(r.id || ''),
    no: String(r.no || ''),
    ts: num_(r.ts),
    datetime: String(r.datetime || ''),
    branchId: String(r.branchId || ''),
    branchName: String(r.branch || ''),
    branch: String(r.branch || ''),
    channel: String(r.channel || ''),
    payment: String(r.payment || ''),
    customer: {
      name: String(r.customerName || ''),
      phone: String(r.customerPhone || ''),
      address: String(r.address || ''),
      memo: String(r.memo || '')
    },
    items: items,
    itemsSummary: String(r.itemsSummary || ''),
    subtotal: num_(r.subtotal),
    discount: {
      type: String(r.discountType || 'none'),
      value: num_(r.discountValue),
      amount: num_(r.discountAmount)
    },
    total: num_(r.total),
    costTotal: num_(r.costTotal),
    costItems: costItems,
    status: String(r.status || 'received'),
    statusUpdatedAt: num_(r.statusUpdatedAt),
    device: String(r.device || ''),
    manager: String(r.manager || ''),
    updatedAt: num_(r.updatedAt) || num_(r.ts),
    syncedAt: num_(r.syncedAt)
  };
}

function orderToRow_(o, existing) {
  var prev = existing || {};
  var cust = o.customer || {};
  var items = Array.isArray(o.items) ? o.items : [];
  var qty = 0;
  var summary = items.map(function (it) {
    var q = Math.max(1, num_(it.qty));
    qty += q;
    return String(it.name || '') + ' x' + q;
  }).join(' | ');

  var discount = o.discount || {};
  var now = Date.now();

  return {
    id: String(o.id || prev.id || ''),
    no: String(o.no || prev.no || ''),
    ts: num_(o.ts) || num_(prev.ts) || now,
    datetime: fmtDateTime_(num_(o.ts) || num_(prev.ts) || now),
    branchId: o.branchId != null ? String(o.branchId) : String(prev.branchId || ''),
    branch: (o.branchName || o.branch) != null ? String(o.branchName || o.branch) : String(prev.branch || ''),
    channel: String(o.channel || prev.channel || ''),
    payment: String(o.payment || prev.payment || ''),
    customerName: cust.name != null ? String(cust.name) : String(prev.customerName || ''),
    customerPhone: cust.phone != null ? String(cust.phone) : String(prev.customerPhone || ''),
    address: cust.address != null ? String(cust.address) : String(prev.address || ''),
    memo: cust.memo != null ? String(cust.memo) : String(prev.memo || ''),
    itemsJson: JSON.stringify(items),
    itemsSummary: summary,
    qty: qty,
    subtotal: num_(o.subtotal != null ? o.subtotal : prev.subtotal),
    discountType: String(discount.type || prev.discountType || 'none'),
    discountValue: num_(discount.value != null ? discount.value : prev.discountValue),
    discountAmount: num_(discount.amount != null ? discount.amount : prev.discountAmount),
    total: num_(o.total != null ? o.total : prev.total),
    costTotal: num_(o.costTotal != null ? o.costTotal : prev.costTotal),
    costItemsJson: JSON.stringify(o.costItems != null ? o.costItems : (prev.costItemsJson || null)),
    status: String(o.status || prev.status || 'received'),
    statusUpdatedAt: num_(o.updatedAt) || num_(prev.statusUpdatedAt) || now,
    device: String(o.device || prev.device || ''),
    manager: String(o.manager || prev.manager || ''),
    updatedAt: num_(o.updatedAt) || now,
    syncedAt: now
  };
}

function saveOrders_(orders) {
  var sh = ensureOrders_();
  var map = indexById_(sh);
  var newRows = [];
  var created = 0, updated = 0, kept = 0, skipped = 0;

  orders.forEach(function (o) {
    if (!o || !o.id) { skipped++; return; }
    var rowNum = map[String(o.id)];
    if (rowNum) {
      var existing = orderFromRow_(rowToObject_(sh, rowNum, ORDER_HEADERS));
      // 시트의 기록이 더 최신이면 상태는 유지하고 나머지만 보완
      var incomingAt = num_(o.updatedAt) || num_(o.ts) || 0;
      if (num_(existing.updatedAt) > incomingAt) {
        o.status = existing.status;
        kept++;
      }
      var merged = orderToRow_(o, rowToObject_(sh, rowNum, ORDER_HEADERS));
      writeRow_(sh, rowNum, merged, ORDER_HEADERS);
      updated++;
    } else {
      newRows.push(orderToRow_(o, null));
      created++;
    }
  });

  if (newRows.length) {
    if (sh.getLastRow() + newRows.length > CONFIG.MAX_ORDERS + 1) {
      throw new Error('주문 행이 너무 많습니다. 시트를 정리해 주세요.');
    }
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, ORDER_HEADERS.length)
      .setValues(newRows.map(function (r) { return headersToValues_(r, ORDER_HEADERS); }));
  }
  log_('orders.push', 'created=' + created + ' updated=' + updated + ' kept=' + kept);
  return { created: created, updated: updated, kept: kept, skipped: skipped, total: readRows_(sh).length };
}

function updateOrderStatus_(id, status, updatedAt) {
  if (!id) throw new Error('주문 id가 없습니다.');
  if (!status) throw new Error('변경할 상태가 없습니다.');
  var sh = ensureOrders_();
  var map = indexById_(sh);
  var rowNum = map[String(id)];
  if (!rowNum) return { updated: 0, note: '해당 주문이 시트에 없습니다.' };
  var row = rowToObject_(sh, rowNum, ORDER_HEADERS);
  var at = num_(updatedAt) || Date.now();
  if (num_(row.statusUpdatedAt) > at) {
    return { updated: 0, note: '시트의 상태가 더 최신입니다.', status: String(row.status) };
  }
  row.status = String(status);
  row.statusUpdatedAt = at;
  row.updatedAt = Math.max(num_(row.updatedAt), at);
  writeRow_(sh, rowNum, row, ORDER_HEADERS);
  log_('order.status', id + ' → ' + status);
  return { updated: 1, id: id, status: status };
}

function deleteOrder_(id) {
  var sh = ensureOrders_();
  var map = indexById_(sh);
  var rowNum = map[String(id)];
  if (!rowNum) return { deleted: 0 };
  sh.deleteRow(rowNum);
  log_('order.delete', String(id));
  return { deleted: 1 };
}

/* ============================ 유틸 ============================ */

function rowToObject_(sh, rowNum, headers) {
  var values = sh.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  var o = {};
  for (var i = 0; i < headers.length; i++) o[headers[i]] = values[i];
  return o;
}

function writeRow_(sh, rowNum, obj, headers) {
  sh.getRange(rowNum, 1, 1, headers.length).setValues([headersToValues_(obj, headers)]);
}

function headersToValues_(obj, headers) {
  return headers.map(function (h) { return obj[h] === undefined || obj[h] === null ? '' : obj[h]; });
}

function counts_() {
  return {
    products: Math.max(0, ensureProducts_().getLastRow() - 1),
    orders: Math.max(0, ensureOrders_().getLastRow() - 1),
    spreadsheetId: ss_().getId()
  };
}

function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = String(v == null ? '' : v).replace(/,/g, '').trim();
  if (!s) return 0;
  var n = Number(s);
  return isFinite(n) ? n : 0;
}

function isY_(v) {
  if (v === true) return true;
  var s = String(v == null ? '' : v).trim().toUpperCase();
  return s === 'Y' || s === 'TRUE' || s === '1' || s === 'YES';
}

function isN_(v) {
  if (v === false) return true;
  var s = String(v == null ? '' : v).trim().toUpperCase();
  return s === 'N' || s === 'FALSE' || s === '0' || s === 'NO';
}

function fmtDateTime_(ms) {
  try {
    return Utilities.formatDate(new Date(ms), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  } catch (e) {
    return new Date(ms).toISOString();
  }
}

function log_(action, detail) {
  try {
    var sh = ensureLog_();
    sh.appendRow([new Date(), String(action), String(detail || '')]);
    var last = sh.getLastRow();
    if (last > 2000) sh.deleteRows(2, last - 2000);
  } catch (e) { /* 로그 실패는 무시 */ }
}
