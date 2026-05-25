const SHEET_ID = '1LxvA2Tnvw_pA8LuVjK70V6VRA-ZxBeL3Wl5mrM4Vtr0';
const LOCKERS  = Array.from({length:12}, (_,i) => 'everbrite' + String(i+1).padStart(3,'0'));
const CACHE_KEY = 'orders_cache';
const CACHE_TTL = 5; // seconds

function doGet(e) {
  try {
    const action = e.parameter.action || 'getOrders';
    let result;
    if (action === 'getOrders')        result = getOrders();
    else if (action === 'addOrder')    result = addOrder(JSON.parse(e.parameter.data));
    else if (action === 'updateOrder') { const {id,patch} = JSON.parse(e.parameter.data); result = updateOrder(id, patch); }
    else if (action === 'deleteOrder') { result = deleteOrder(JSON.parse(e.parameter.data).id); }
    else if (action === 'clearOrders') { result = clearOrders(); }
    else                               result = {error:'unknown action'};
    return jsonR(result);
  } catch(err) { return jsonR({error: err.message}); }
}

function getOrders() {
  // 先查快取
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);
  if (cached) return JSON.parse(cached);

  // 讀試算表，只過濾不刪除（刪除交給 addOrder）
  const s = getSheet(), d = s.getDataRange().getValues();
  if (d.length <= 1) {
    cache.put(CACHE_KEY, '{}', CACHE_TTL);
    return {};
  }

  const h = d[0], now = Date.now(), orders = {};
  for (let i = 1; i < d.length; i++) {
    if (!d[i][0]) continue;
    const o = {}; h.forEach((k,j) => { o[k] = d[i][j]; });
    try { o.items = JSON.parse(o.items); } catch(_) { o.items = []; }
    o.t = Number(o.t);
    const ts = new Date(String(o.ts).length === 16 ? o.ts+':00' : o.ts).getTime();
    if (now - ts <= 15*60*1000) orders[o.id] = o; // 過期的直接跳過，不刪
  }

  cache.put(CACHE_KEY, JSON.stringify(orders), CACHE_TTL);
  return orders;
}

function addOrder(order) {
  const lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    const s = getSheet(); ensureH(s);

    // 清除快取，順便刪除過期列
    CacheService.getScriptCache().remove(CACHE_KEY);
    purgeExpired(s);

    const orders = getOrdersFromSheet(s);
    const inUse = new Set(Object.values(orders).filter(o=>o.status!=='picked_up').map(o=>o.locker).filter(Boolean));
    const p = PropertiesService.getScriptProperties();
    const last = parseInt(p.getProperty('li') || '-1', 10);
    let idx = (last+1) % 12;
    for (let i = 1; i <= 12; i++) { const j=(last+i)%12; if(!inUse.has(LOCKERS[j])){ idx=j; break; } }
    p.setProperty('li', String(idx));
    order.locker = LOCKERS[idx];

    const h = ['id','u','pic','uid','items','t','locker','status','ts','placedAt','pickedAt'];
    s.appendRow(h.map(k => k==='items' ? JSON.stringify(order.items||[]) : k==='t' ? order.t||0 : order[k]||''));

    CacheService.getScriptCache().remove(CACHE_KEY);
    return order;
  } finally { lock.releaseLock(); }
}

function updateOrder(id, patch) {
  const s=getSheet(), d=s.getDataRange().getValues(), h=d[0], c=h.indexOf('id');
  for (let i=1; i<d.length; i++) {
    if (String(d[i][c]) === String(id)) {
      Object.entries(patch).forEach(([k,v]) => { const j=h.indexOf(k); if(j>=0) s.getRange(i+1,j+1).setValue(v); });
      CacheService.getScriptCache().remove(CACHE_KEY);
      return {success:true};
    }
  }
  return {error:'not found'};
}

function clearOrders() {
  const s = getSheet();
  if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
  CacheService.getScriptCache().remove(CACHE_KEY);
  return {success: true};
}

function deleteOrder(id) {
  const s=getSheet(), d=s.getDataRange().getValues(), h=d[0], c=h.indexOf('id');
  for (let i=1; i<d.length; i++) {
    if (String(d[i][c]) === String(id)) {
      s.deleteRow(i+1);
      CacheService.getScriptCache().remove(CACHE_KEY);
      return {success:true};
    }
  }
  return {error:'not found'};
}

// 刪除超過 15 分鐘的列（只在 addOrder 時呼叫）
function purgeExpired(s) {
  const d = s.getDataRange().getValues();
  if (d.length <= 1) return;
  const h = d[0], now = Date.now(), exp = [];
  for (let i = 1; i < d.length; i++) {
    if (!d[i][0]) continue;
    const o = {}; h.forEach((k,j) => { o[k] = d[i][j]; });
    const ts = new Date(String(o.ts).length === 16 ? o.ts+':00' : o.ts).getTime();
    if (now - ts > 15*60*1000) exp.push(i+1);
  }
  for (let i = exp.length-1; i >= 0; i--) s.deleteRow(exp[i]);
}

// 直接讀試算表（不走快取，供 addOrder 內部用）
function getOrdersFromSheet(s) {
  const d = s.getDataRange().getValues();
  if (d.length <= 1) return {};
  const h = d[0], now = Date.now(), orders = {};
  for (let i = 1; i < d.length; i++) {
    if (!d[i][0]) continue;
    const o = {}; h.forEach((k,j) => { o[k] = d[i][j]; });
    try { o.items = JSON.parse(o.items); } catch(_) { o.items = []; }
    o.t = Number(o.t);
    const ts = new Date(String(o.ts).length === 16 ? o.ts+':00' : o.ts).getTime();
    if (now - ts <= 15*60*1000) orders[o.id] = o;
  }
  return orders;
}

function getSheet() { const ss=SpreadsheetApp.openById(SHEET_ID); return ss.getSheetByName('orders')||ss.insertSheet('orders'); }
function ensureH(s) { const h=['id','u','pic','uid','items','t','locker','status','ts','placedAt','pickedAt']; if(!s.getRange(1,1).getValue()) s.getRange(1,1,1,h.length).setValues([h]); }
function jsonR(d) { return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON); }
