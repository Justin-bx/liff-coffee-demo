const SHEET_ID = '1LxvA2Tnvw_pA8LuVjK70V6VRA-ZxBeL3Wl5mrM4Vtr0';
const LOCKERS  = Array.from({length:12}, (_,i) => 'everbrite' + String(i+1).padStart(3,'0'));

// ── 所有請求統一走 GET（避免 CORS preflight 問題）──
function doGet(e) {
  try {
    const action = e.parameter.action || 'getOrders';
    let result;

    if (action === 'getOrders') {
      result = getOrders();
    } else if (action === 'addOrder') {
      const order = JSON.parse(e.parameter.data);
      result = addOrder(order);
    } else if (action === 'updateOrder') {
      const { id, patch } = JSON.parse(e.parameter.data);
      result = updateOrder(id, patch);
    } else {
      result = { error: 'unknown action' };
    }

    return jsonResp(result);
  } catch(err) {
    return jsonResp({ error: err.message });
  }
}

// ── 讀取所有訂單（同時清除超過 15 分鐘的訂單）──
function getOrders() {
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};

  const headers  = data[0];
  const now      = Date.now();
  const orders   = {};
  const expRows  = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;

    const o = {};
    headers.forEach((h, j) => { o[h] = row[j]; });
    try { o.items = JSON.parse(o.items); } catch(_) { o.items = []; }
    o.t = Number(o.t);

    const ts = new Date(String(o.ts).length === 16 ? o.ts + ':00' : o.ts).getTime();
    if (now - ts > 15 * 60 * 1000) {
      expRows.push(i + 1);          // 記錄要刪除的列（試算表列號）
    } else {
      orders[o.id] = o;
    }
  }

  // 從下往上刪，避免列號偏移
  for (let i = expRows.length - 1; i >= 0; i--) {
    sheet.deleteRow(expRows[i]);
  }

  return orders;
}

// ── 新增訂單（含格子分配，使用 Lock 防止並發衝突）──
function addOrder(order) {
  const lock = LockService.getScriptLock();
  lock.waitLock(8000);

  try {
    const sheet = getSheet();
    ensureHeaders(sheet);

    // 讀取目前使用中的格子
    const orders = getOrders();
    const inUse  = new Set(
      Object.values(orders)
        .filter(o => o.status !== 'picked_up')
        .map(o => o.locker).filter(Boolean)
    );

    // 從上次使用的格子往後找
    const props   = PropertiesService.getScriptProperties();
    const lastIdx = parseInt(props.getProperty('last_locker_idx') || '-1', 10);
    let assignedIdx = (lastIdx + 1) % 12;

    for (let i = 1; i <= 12; i++) {
      const idx = (lastIdx + i) % 12;
      if (!inUse.has(LOCKERS[idx])) {
        assignedIdx = idx;
        break;
      }
    }

    props.setProperty('last_locker_idx', String(assignedIdx));
    order.locker = LOCKERS[assignedIdx];

    // 寫入試算表
    const headers = ['id','u','pic','uid','items','t','locker','status','ts','placedAt','pickedAt'];
    const row = headers.map(h => {
      if (h === 'items') return JSON.stringify(order.items || []);
      if (h === 't')     return order.t || 0;
      return order[h] || '';
    });
    sheet.appendRow(row);

    return order;
  } finally {
    lock.releaseLock();
  }
}

// ── 更新訂單欄位 ──
function updateOrder(id, patch) {
  const sheet   = getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol   = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) {
      Object.entries(patch).forEach(([key, val]) => {
        const col = headers.indexOf(key);
        if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(val);
      });
      return { success: true };
    }
  }
  return { error: 'not found' };
}

// ── 工具函式 ──
function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName('orders') || ss.insertSheet('orders');
}

function ensureHeaders(sheet) {
  const headers  = ['id','u','pic','uid','items','t','locker','status','ts','placedAt','pickedAt'];
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (!firstRow[0]) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function jsonResp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
