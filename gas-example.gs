// Google Apps Script example (deploy as Web App with Anyone with link)
// Spreadsheet columns: orderId, customerName, customerNameNorm, startDate, warrantyDays, locked, lockReason, pricePerDay, paymentStatus, pendingDays, pendingAmount, receiptDataUrl, createdAt

function doGet(e) {
  const sheet = _sheet();
  const action = e && e.parameter && e.parameter.action;
  if (action === 'adminMagic') {
    const code = (e.parameter.code || '').trim();
    const ok = _verifyMagic(code);
    if (!ok) return _json({ ok: false, error: 'Invalid code' }, 401);
    const token = Utilities.getUuid();
    _cache().put('adm:' + token, '1', 3600); // 1h
    return _json({ ok: true, token });
  }
  if (action === 'whoami') {
    const token = (e.parameter.token || '').trim();
    const isAdmin = !!_cache().get('adm:' + token);
    return _json({ ok: true, isAdmin });
  }
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const rows = values.map(r => Object.fromEntries(headers.map((h, i) => [h, r[i]])));
  return ContentService.createTextOutput(JSON.stringify({ ok: true, data: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  const action = body.action;
  const sheet = _sheet();
  if (action === 'createOrUpdate') {
    const rec = body.record || {};
    const values = sheet.getDataRange().getValues();
    const headers = values.shift();
    let rowIndex = -1;
    for (let i = 0; i < values.length; i++) { if (values[i][0] === rec.orderId) { rowIndex = i + 2; break; } }
    // Merge with existing row if present to avoid losing fields like customerName
    let curr = {};
    if (rowIndex > 0) {
      const existing = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
      curr = Object.fromEntries(headers.map((h, i) => [h, existing[i]]));
    }
    Object.keys(rec).forEach(k => { if (typeof rec[k] !== 'undefined') curr[k] = rec[k]; });
    curr.orderId = curr.orderId || rec.orderId;
    if (typeof curr.locked === 'undefined') curr.locked = false;
    curr.lockReason = curr.lockReason || "";
    curr.pricePerDay = Number(curr.pricePerDay || 0);
    curr.paymentStatus = curr.paymentStatus || "";
    curr.pendingDays = Number(curr.pendingDays || 0);
    curr.pendingAmount = Number(curr.pendingAmount || 0);
    curr.receiptDataUrl = curr.receiptDataUrl || "";

    const outRow = headers.map((h)=> h === 'createdAt' ? new Date() : curr[h]);
    if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, headers.length).setValues([outRow]); else sheet.appendRow(outRow);
    return _json({ ok: true });
  }
  if (action === 'delete') {
    const orderId = body.orderId;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) { if (rows[i][0] === orderId) { sheet.deleteRow(i + 1); break; } }
    return _json({ ok: true });
  }
  return _json({ ok: false, error: 'Unknown action' }, 400);
}

function _sheet() {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  const sheet = ss.getSheetByName('Orders') || ss.insertSheet('Orders');
  if (sheet.getLastRow() === 0) sheet.appendRow(['orderId', 'customerName', 'customerNameNorm', 'startDate', 'warrantyDays', 'locked', 'lockReason', 'pricePerDay', 'paymentStatus', 'pendingDays', 'pendingAmount', 'receiptDataUrl', 'createdAt']);
  return sheet;
}

function _json(obj, code) {
  const out = ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  if (code) out.setResponseCode(code);
  return out;
}

function _cache() {
  return CacheService.getScriptCache();
}

function _verifyMagic(code) {
  const prop = PropertiesService.getScriptProperties();
  const secret = (prop.getProperty('MAGIC_CODE') || '').trim();
  return !!secret && code === secret;
}


