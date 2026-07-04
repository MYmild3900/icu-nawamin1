// ═══════════════════════════════════════════════════
// ICU Nawamin1 — Apps Script v3 (Items API)
// สำเนาโค้ดที่ต้อง deploy ในโปรเจกต์ "Code จาก Gemini"
// ═══════════════════════════════════════════════════

const SHEET_ID = '1LI_5pWF5XCxgMi8OhxXir42kiTpqpfUBat8egzGH4u4';

// tab ที่ไม่ใช่ชีตหมวดพัสดุ — ห้ามอ่าน/เขียนเป็นรายการพัสดุ
const NON_ITEM_SHEETS = ['ประวัติรายการ', 'Items', 'Log', 'ชีต2'];

function isItemSheet(sh) {
  if (NON_ITEM_SHEETS.indexOf(sh.getName()) !== -1) return false;
  // ชีตหมวดต้องมีหัวตาราง "รหัสพัสดุ" ที่แถว 3 คอลัมน์ A
  var a3 = (sh.getRange(3, 1).getValue() || '').toString();
  return a3.indexOf('รหัสพัสดุ') !== -1;
}

// ── GET ──
function doGet(e) {
  var action = (e.parameter && e.parameter.action) || 'getItems';
  var result;
  try {
    if (action === 'ping') {
      result = { ok: true, time: new Date().toISOString() };
    } else if (action === 'getItems') {
      result = getItems();
    } else if (action === 'getLots') {
      result = getLots();
    } else if (action === 'getLog') {
      result = getLog();
    } else {
      result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST ──
function doPost(e) {
  var result;
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    if (action === 'saveItems')        result = saveItems(data.items);
    else if (action === 'addLog')      result = addLog(data.log);
    else if (action === 'addItem')     result = addItem(data.item);
    else if (action === 'updateStock') result = updateStock(data.code, data.cur);
    else if (action === 'upsertLot')   result = upsertLot(data);
    else if (action === 'deductLot')   result = deductLot(data);
    else if (action === 'addLot')      result = { error: 'addLot ถูกปิดการใช้งานแล้ว — ใช้ upsertLot/deductLot แทน (ไม่แทรกแถวในชีตพัสดุ)' };
    else result = { error: 'Unknown action' };
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── updateStock: แก้ยอดคงเหลือ (คอลัมน์ F) ในแถวเดิมของชีตหมวด — ไม่แทรกแถวใหม่ ──
function updateStock(code, cur) {
  if (!code) return { error: 'ไม่มีรหัสพัสดุ' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();
  var codeStr = code.toString().trim();
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    if (!isItemSheet(sh)) continue;
    var lastRow = sh.getLastRow();
    if (lastRow < 4) continue;
    var colA = sh.getRange(4, 1, lastRow - 3, 1).getValues();
    for (var i = 0; i < colA.length; i++) {
      if ((colA[i][0] || '').toString().trim() === codeStr) {
        sh.getRange(i + 4, 6).setValue(Number(cur) || 0); // คอลัมน์ F = Stock คงเหลือ ★
        return { ok: true, sheet: sh.getName(), row: i + 4, cur: Number(cur) || 0 };
      }
    }
  }
  return { error: 'Item not found: ' + codeStr };
}

// ── getItems: อ่านเฉพาะชีตหมวดพัสดุจริงเท่านั้น ──
function getItems() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();
  var items = [];
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    if (!isItemSheet(sh)) continue;
    var lastRow = sh.getLastRow();
    if (lastRow < 4) continue;
    var rows = sh.getRange(4, 1, lastRow - 3, 12).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var code = (r[0] || '').toString().trim();
      if (!code || code.charAt(0) === '★' || code.length > 20) continue; // ข้ามแถวว่าง/แถวคำอธิบาย
      var cur = Number(r[5]) || 0, min = Number(r[3]) || 0;
      items.push({
        code: code, name: r[1],
        cat: sh.getName(),
        unit: r[2], cur: cur,
        min: min, max: Number(r[4]) || 0,
        loc: r[6] || '', lot: r[7] || '-',
        recv: r[8] || '-', exp: r[9] || '-',
        imgUrl: r[11] || '',
        status: cur <= 0 ? 'rd' : cur <= min ? 'wn' : 'ok',
        unitPack: '-',
        lots: [{ lot: r[7] || '-', recv: r[8] || '-', exp: r[9] || '-', qty: cur }]
      });
    }
  }
  return { ok: true, items: items, count: items.length };
}

// ── saveItems: snapshot รายการทั้งหมดลง tab "Items" (ไม่แตะชีตหมวด) ──
function saveItems(items) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Items');
  if (!sheet) {
    sheet = ss.insertSheet('Items');
    sheet.getRange(1, 1, 1, 14).setValues([[
      'code', 'name', 'cat', 'icon', 'unit', 'cur', 'min', 'max',
      'loc', 'lot', 'recv', 'exp', 'imgUrl', 'status'
    ]]);
  }
  var rows = items.map(function (item) {
    var sts = item.cur <= 0 ? 'rd' : item.cur <= item.min ? 'wn' : 'ok';
    return [
      item.code, item.name, item.cat, item.icon || '📦',
      item.unit, item.cur || 0, item.min || 0, item.max || 0,
      item.loc || '', item.lot || '—', item.recv || '—', item.exp || '—',
      item.imgUrl || '', sts
    ];
  });
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 14).clearContent();
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 14).setValues(rows);
  return { ok: true, saved: rows.length };
}

// ── addItem: เพิ่มพัสดุใหม่ (แถวใหม่ในชีตหมวด — ใช้เฉพาะกรณีสร้างรายการใหม่จริงๆ) ──
function addItem(item) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(item.cat);
  if (!sheet) {
    var catName = item.cat.replace(/^.*\d+\./, '').trim();
    sheet = ss.getSheetByName(catName);
  }
  if (!sheet) return { error: 'Sheet not found: ' + item.cat };

  // กันเพิ่มรหัสซ้ำ — ถ้ามีอยู่แล้วให้อัปเดตยอดแทน
  var lastRow = sheet.getLastRow();
  if (lastRow >= 4) {
    var colA0 = sheet.getRange(4, 1, lastRow - 3, 1).getValues();
    for (var k = 0; k < colA0.length; k++) {
      if ((colA0[k][0] || '').toString().trim() === (item.code || '').toString().trim()) {
        sheet.getRange(k + 4, 6).setValue(Number(item.cur) || 0);
        return { ok: true, updated: true, row: k + 4 };
      }
    }
  }

  var lot = (item.lots && item.lots[0]) ? item.lots[0] : {};
  // หาแถวสุดท้ายของข้อมูลจริง (ก่อนแถวว่างหรือแถว ★)
  var colA = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  var insertAt = 4;
  for (var i = 3; i < colA.length; i++) {
    var cell = (colA[i][0] || '').toString().trim();
    if (cell !== '' && cell.charAt(0) !== '★' && cell.charAt(0) !== '*') {
      insertAt = i + 2;
    } else {
      break;
    }
  }
  sheet.insertRowBefore(insertAt);
  sheet.getRange(insertAt, 1, 1, 12).setValues([[
    item.code, item.name, item.unit,
    item.min || 0, item.max || 0, item.cur || 0,
    item.loc || '', lot.lot || '', lot.recv || '', lot.exp || '', '',
    (item.imgUrl && item.imgUrl.indexOf('http') === 0) ? item.imgUrl : ''
  ]]);
  return { ok: true };
}

// ── addLog / getLog: สำรองไว้ (แอปหลักบันทึกประวัติผ่าน History API อีกโปรเจกต์) ──
function addLog(log) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Log');
  if (!sheet) {
    sheet = ss.insertSheet('Log');
    sheet.getRange(1, 1, 1, 8).setValues([[
      'datetime', 'type', 'code', 'name', 'qty', 'lot', 'user', 'newStock'
    ]]);
  }
  sheet.appendRow([
    new Date(), log.type === 'recv' ? 'รับเข้า' : 'เบิกออก',
    log.code, log.name, log.qty,
    log.lot || '—', log.user || '—', log.newStock || 0
  ]);
  return { ok: true };
}

function getLog() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Log');
  if (!sheet) return { ok: true, logs: [] };
  var rows = sheet.getDataRange().getValues();
  var logs = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    logs.push({
      date: r[0], type: r[1], code: r[2], name: r[3],
      qty: r[4], lot: r[5], user: r[6], newStock: r[7]
    });
  }
  return { ok: true, logs: logs };
}

// ═══ ทะเบียน Lot (tab "Lot") — 1 แถวต่อ 1 lot ของพัสดุแต่ละตัว ═══

function getLotSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Lot');
  if (!sheet) {
    sheet = ss.insertSheet('Lot');
    sheet.getRange(1, 1, 1, 6).setValues([['code', 'name', 'lot', 'qty', 'recv', 'exp']]);
  }
  return sheet;
}

function findLotRow_(sheet, code, lot) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var vals = sheet.getRange(2, 1, last - 1, 3).getValues();
  var c = String(code).trim(), l = String(lot || '—').trim();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === c && String(vals[i][2]).trim() === l) return i + 2;
  }
  return -1;
}

function getLots() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Lot');
  if (!sheet) return { ok: true, lots: [] };
  var rows = sheet.getDataRange().getValues();
  var lots = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    lots.push({ code: String(r[0]), name: r[1], lot: String(r[2]), qty: Number(r[3]) || 0, recv: r[4], exp: r[5] });
  }
  return { ok: true, lots: lots };
}

// รับเข้า: บวกยอด lot เดิม หรือเพิ่มแถว lot ใหม่ (เพิ่มเฉพาะใน tab Lot เท่านั้น)
function upsertLot(data) {
  if (!data.code) return { error: 'ไม่มีรหัสพัสดุ' };
  var sheet = getLotSheet_();
  var lot = String(data.lot || '—').trim() || '—';
  var row = findLotRow_(sheet, data.code, lot);
  if (row === -1) {
    sheet.appendRow([String(data.code), data.name || '', lot, Number(data.qty) || 0, data.recv || '', data.exp || '']);
    return { ok: true, created: true };
  }
  var cur = Number(sheet.getRange(row, 4).getValue()) || 0;
  sheet.getRange(row, 4).setValue(cur + (Number(data.qty) || 0));
  if (data.recv) sheet.getRange(row, 5).setValue(data.recv);
  if (data.exp)  sheet.getRange(row, 6).setValue(data.exp);
  return { ok: true, row: row, qty: cur + (Number(data.qty) || 0) };
}

// เบิกออก: ตัดยอดของ lot ที่เลือก (ต่ำสุด 0) — lot ที่ไม่อยู่ในทะเบียนให้ข้าม ไม่ถือเป็น error
function deductLot(data) {
  if (!data.code) return { error: 'ไม่มีรหัสพัสดุ' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Lot');
  if (!sheet) return { ok: true, skipped: true };
  var lot = String(data.lot || '—').trim() || '—';
  var row = findLotRow_(sheet, data.code, lot);
  if (row === -1) return { ok: true, skipped: true };
  var cur = Number(sheet.getRange(row, 4).getValue()) || 0;
  var next = Math.max(0, cur - (Number(data.qty) || 0));
  sheet.getRange(row, 4).setValue(next);
  return { ok: true, row: row, qty: next };
}

// ── ทดสอบด้วยมือใน editor ──
function testUpdateStock() {
  Logger.log(JSON.stringify(updateStock('60101009', 5)));
}
function testLots() {
  Logger.log(JSON.stringify(upsertLot({ code: '60101009', name: 'ทดสอบ lot', lot: 'L-TEST', qty: 3, recv: '4/7/69', exp: '2026-12-31' })));
  Logger.log(JSON.stringify(deductLot({ code: '60101009', lot: 'L-TEST', qty: 1 })));
  Logger.log(JSON.stringify(getLots()));
}
