// ═══════════════════════════════════════════════════
// ICU Nawamin1 — Apps Script v3 (Items API)
// สำเนาโค้ดที่ต้อง deploy ในโปรเจกต์ "Code จาก claude"
// ═══════════════════════════════════════════════════

const SHEET_ID = '1LI_5pWF5XCxgMi8OhxXir42kiTpqpfUBat8egzGH4u4';

// tab ที่ไม่ใช่ชีตหมวดพัสดุ — ห้ามอ่าน/เขียนเป็นรายการพัสดุ
const NON_ITEM_SHEETS = ['ประวัติรายการ', 'Items', 'Log', 'ชีต2', 'Lot', 'Staff'];

// อีเมลรับแจ้งเตือนประจำวัน — ใส่ได้หลายคน คั่นด้วย , (คนที่ 2 = อีเมลหอผู้ป่วย)
const DIGEST_EMAIL = 'mymild.mildmy@gmail.com,icutrauma2025@gmail.com';

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
    } else if (action === 'getStaff') {
      result = getStaff();
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
    else if (action === 'saveStaff')   result = saveStaffList(data.staff);
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

function findLotRow_(sheet, code, lot, name) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var vals = sheet.getRange(2, 1, last - 1, 3).getValues();
  var c = String(code).trim(), l = String(lot || '—').trim();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === c && String(vals[i][2]).trim() === l) return i + 2;
  }
  // แถวที่กรอกมือโดยไม่มีรหัส: จับคู่ด้วยชื่อ+ล็อตแทน แล้วเติมรหัสกลับลงแถวให้เลย (ซ่อมตัวเอง)
  if (name) {
    var norm = function (s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); };
    var n = norm(name);
    for (var j = 0; j < vals.length; j++) {
      if (!String(vals[j][0]).trim() && norm(vals[j][1]) === n && String(vals[j][2]).trim() === l) {
        sheet.getRange(j + 2, 1).setValue(c);
        return j + 2;
      }
    }
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
  var row = findLotRow_(sheet, data.code, lot, data.name);
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
  var row = findLotRow_(sheet, data.code, lot, data.name);
  if (row === -1) return { ok: true, skipped: true };
  var cur = Number(sheet.getRange(row, 4).getValue()) || 0;
  var next = Math.max(0, cur - (Number(data.qty) || 0));
  sheet.getRange(row, 4).setValue(next);
  return { ok: true, row: row, qty: next };
}

// ═══ บุคลากร (tab "Staff") — รายชื่อผู้ใช้ระบบ ═══

function getStaffSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Staff');
  if (!sheet) {
    sheet = ss.insertSheet('Staff');
    sheet.getRange(1, 1, 1, 5).setValues([['id', 'name', 'pos', 'role', 'pw']]);
  }
  // เติมหัวคอลัมน์ pw ให้ชีตเก่าที่ยังไม่มี (ซ่อมตัวเอง)
  if (String(sheet.getRange(1, 5).getValue()).trim() !== 'pw') {
    sheet.getRange(1, 5).setValue('pw');
  }
  return sheet;
}

function getStaff() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Staff');
  if (!sheet) return { ok: true, staff: [] };
  var rows = sheet.getDataRange().getValues();
  var staff = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[1]) continue;
    staff.push({ id: Number(r[0]) || i, name: String(r[1]), pos: String(r[2] || ''), role: String(r[3] || 'user'), pw: String(r[4] || '') });
  }
  return { ok: true, staff: staff };
}

// บันทึกรายชื่อทั้งชุด (snapshot แทนที่ของเดิม) — รวมรหัสผ่าน (คอลัมน์ pw)
function saveStaffList(staffArr) {
  var sheet = getStaffSheet_();
  var last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, 5).clearContent();
  var rows = (staffArr || []).map(function (s) {
    return [Number(s.id) || '', String(s.name || ''), String(s.pos || ''), String(s.role || 'user'), String(s.pw || '')];
  }).filter(function (r) { return r[1] !== ''; });
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  return { ok: true, saved: rows.length };
}

// ═══ อีเมลแจ้งเตือนประจำวัน (หมดสต็อก / ใกล้หมด / หมดอายุ / ใกล้หมดอายุ) ═══

// แปลงวันหมดอายุ → Date (รองรับ Date object, ISO, ปี พ.ศ., วว/ดด/ปปปป, ดด/ปป)
function parseExpDateGS_(exp) {
  if (exp === null || exp === undefined || exp === '' || exp === '—' || exp === '-') return null;
  if (Object.prototype.toString.call(exp) === '[object Date]') {
    if (isNaN(exp)) return null;
    var yy = exp.getFullYear();
    return (yy > 2400) ? new Date(yy - 543, exp.getMonth(), exp.getDate()) : exp;
  }
  var s = String(exp).trim();
  var beCE = function (y) { if (y < 100) y += 2500; if (y > 2400) y -= 543; return y; };
  var m3 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);   // วว/ดด/ปปปป
  if (m3) {
    var dd = parseInt(m3[1], 10), mo = parseInt(m3[2], 10), yr = beCE(parseInt(m3[3], 10));
    if (mo >= 1 && mo <= 12 && dd >= 1 && dd <= 31) return new Date(yr, mo - 1, dd);
  }
  var m2 = s.match(/^(\d{1,2})[\/\-](\d{2,4})$/);                   // ดด/ปป → สิ้นเดือน
  if (m2) {
    var mo2 = parseInt(m2[1], 10), yr2 = beCE(parseInt(m2[2], 10));
    if (mo2 >= 1 && mo2 <= 12) return new Date(yr2, mo2, 0);
  }
  var d = new Date(s);
  if (!isNaN(d)) {
    var y = d.getFullYear();
    return (y > 2400) ? new Date(y - 543, d.getMonth(), d.getDate()) : d;
  }
  return null;
}
function daysToExpGS_(exp) {
  var d = parseExpDateGS_(exp);
  if (!d) return null;
  var today = new Date(); today.setHours(0, 0, 0, 0);
  d = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.round((d - today) / 86400000);
}
function fmtD_(exp) {
  var d = parseExpDateGS_(exp);
  return d ? Utilities.formatDate(d, 'Asia/Bangkok', 'dd/MM/yyyy') : String(exp);
}

// รวบรวมข้อมูลแจ้งเตือนทั้ง 4 กลุ่ม
function buildDigestData_() {
  var items = getItems().items;
  var outStock = items.filter(function (i) { return (Number(i.cur) || 0) <= 0; });
  var low = items.filter(function (i) {
    var c = Number(i.cur) || 0, m = Number(i.min) || 0;
    return c > 0 && m > 0 && c <= m;
  });
  // แผนที่หน่วยนับ (ตามชื่อพัสดุ) เพื่อแสดง "เหลือ N หน่วย" ในอีเมล
  var unitByName = {};
  items.forEach(function (i) { if (i.name) unitByName[i.name] = i.unit || ''; });
  // วันหมดอายุ: จาก tab Lot (ราย lot ที่ยังมีของ) + จากคอลัมน์ exp ของชีตหมวด (ตัวที่ไม่มีใน Lot)
  var seen = {};
  var expired = [], soon = [];
  function consider(name, lot, exp, qty) {
    var dd = daysToExpGS_(exp);
    if (dd === null) return;
    var key = name + '|' + lot;
    if (seen[key]) return;
    seen[key] = true;
    var row = { name: name, lot: lot || '—', exp: fmtD_(exp), days: dd, qty: (Number(qty) || 0), unit: unitByName[name] || '' };
    if (dd < 0) expired.push(row);
    else if (dd <= 90) soon.push(row);
  }
  getLots().lots.forEach(function (L) {
    if ((Number(L.qty) || 0) <= 0) return;
    consider(L.name || L.code, L.lot, L.exp, L.qty);
  });
  items.forEach(function (i) {
    if ((Number(i.cur) || 0) <= 0) return;
    consider(i.name, i.lot, i.exp, i.cur);
  });
  expired.sort(function (a, b) { return a.days - b.days; });
  soon.sort(function (a, b) { return a.days - b.days; });
  return { outStock: outStock, low: low, expired: expired, soon: soon };
}

function digestHtml_(d) {
  function sec(title, color, rowsHtml) {
    if (!rowsHtml) return '';
    return '<h3 style="color:' + color + ';margin:16px 0 6px;font-size:15px">' + title + '</h3>'
      + '<table style="border-collapse:collapse;width:100%;font-size:13px">' + rowsHtml + '</table>';
  }
  function tr(cells) {
    return '<tr>' + cells.map(function (c) {
      return '<td style="border:1px solid #ddd;padding:5px 8px">' + c + '</td>';
    }).join('') + '</tr>';
  }
  var h = '';
  h += sec('🚫 ของหมดอายุแล้ว — ห้ามใช้ ต้องนำออกทันที (' + d.expired.length + ')', '#8e1b0f',
    d.expired.map(function (r) { return tr([r.name, 'Lot ' + r.lot, 'หมดอายุ ' + r.exp, 'จำนวน ' + r.qty + (r.unit ? ' ' + r.unit : ''), 'เลยกำหนด ' + (-r.days) + ' วัน']); }).join(''));
  h += sec('⏰ ใกล้ถึงวันหมดอายุ ภายใน 90 วัน (' + d.soon.length + ')', '#b26a00',
    d.soon.map(function (r) { return tr([r.name, 'Lot ' + r.lot, 'หมดอายุ ' + r.exp, 'จำนวน ' + r.qty + (r.unit ? ' ' + r.unit : ''), 'อีก ' + r.days + ' วัน']); }).join(''));
  h += sec('🔴 หมดสต็อก — คงเหลือ 0 (' + d.outStock.length + ')', '#c0392b',
    d.outStock.map(function (i) { return tr([i.name, String(i.cat || ''), 'Min ' + (i.min || 0)]); }).join(''));
  h += sec('🟡 ใกล้หมดสต็อก — ถึงจุดต่ำสุด (' + d.low.length + ')', '#b26a00',
    d.low.map(function (i) { return tr([i.name, String(i.cat || ''), 'เหลือ ' + i.cur + ' (Min ' + i.min + ')']); }).join(''));
  return '<div style="font-family:Tahoma,sans-serif;max-width:640px">'
    + '<h2 style="color:#1b2a4a;margin:0 0 4px">📦 ระบบพัสดุ ICU นวมินทร์ 1 — แจ้งเตือนประจำวัน</h2>'
    + '<div style="color:#777;font-size:12px;margin-bottom:8px">' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') + ' น. · โรงพยาบาลลำปาง</div>'
    + h
    + '<div style="color:#999;font-size:11px;margin-top:16px">เปิดระบบ: https://mymild3900.github.io/icu-nawamin1/</div></div>';
}

// ส่งอีเมลแจ้งเตือน — ส่งเฉพาะวันที่มีเรื่องต้องเตือน (trigger เรียกทุกวัน ~08:00)
function dailyDigest() {
  var d = buildDigestData_();
  var total = d.outStock.length + d.low.length + d.expired.length + d.soon.length;
  if (total === 0) return;   // ไม่มีอะไรต้องเตือน — ไม่ส่ง
  var subject = '[พัสดุ ICU1] แจ้งเตือน: หมดอายุแล้ว ' + d.expired.length
    + ' · ใกล้หมดอายุ ' + d.soon.length
    + ' · หมดสต็อก ' + d.outStock.length
    + ' · ใกล้หมด ' + d.low.length;
  MailApp.sendEmail({ to: DIGEST_EMAIL, subject: subject, htmlBody: digestHtml_(d) });
}

// ▶ กดรันครั้งเดียว: ตั้งเวลาส่งอัตโนมัติทุกวัน 08:00-09:00 น.
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyDigest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyDigest').timeBased().everyDays(1).atHour(8).create();
  Logger.log('ตั้งเวลาส่งแจ้งเตือนทุกวัน 08:00-09:00 น. เรียบร้อย');
}

// ▶ กดรันเพื่อทดสอบส่งอีเมลทันที (ส่งเสมอแม้ไม่มีเรื่องเตือน)
function sendTestDigest() {
  var d = buildDigestData_();
  var total = d.outStock.length + d.low.length + d.expired.length + d.soon.length;
  var body = total === 0
    ? '<div style="font-family:Tahoma">✅ (ทดสอบ) วันนี้ไม่มีพัสดุที่ต้องเตือนค่ะ</div>'
    : digestHtml_(d);
  MailApp.sendEmail({ to: DIGEST_EMAIL, subject: '[พัสดุ ICU1] ทดสอบระบบแจ้งเตือนอีเมล', htmlBody: body });
  Logger.log('ส่งอีเมลทดสอบไปที่ ' + DIGEST_EMAIL + ' แล้ว (รายการเตือนรวม ' + total + ')');
}

// ═══════════ สำรองข้อมูลอัตโนมัติรายสัปดาห์ ═══════════
// คัดลอกสเปรดชีตทั้งไฟล์เก็บในโฟลเดอร์ "สำรองข้อมูล พัสดุICU1" ใน Google Drive
// เก็บย้อนหลัง 8 ชุดล่าสุด — ชุดเก่ากว่านั้นย้ายลงถังขยะอัตโนมัติ

const BACKUP_FOLDER_NAME = 'สำรองข้อมูล พัสดุICU1';
const BACKUP_KEEP = 8; // จำนวนสำเนาที่เก็บไว้

function getBackupFolder_() {
  var it = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

function weeklyBackup() {
  var folder = getBackupFolder_();
  var src = DriveApp.getFileById(SHEET_ID);
  var d = new Date();
  var stamp = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');
  var copy = src.makeCopy('สำรอง พัสดุICU1 ' + stamp, folder);

  // ลบสำเนาเก่า เหลือไว้ BACKUP_KEEP ชุดล่าสุด
  var files = [];
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getName().indexOf('สำรอง พัสดุICU1') === 0) files.push(f);
  }
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  for (var i = BACKUP_KEEP; i < files.length; i++) files[i].setTrashed(true);

  Logger.log('สำรองข้อมูลแล้ว: ' + copy.getName() + ' (เก็บย้อนหลัง ' + Math.min(files.length, BACKUP_KEEP) + ' ชุด)');
  return copy.getName();
}

// ▶ กดรันครั้งเดียว: ตั้งสำรองข้อมูลอัตโนมัติทุกวันจันทร์ 07:00-08:00 น.
function setupWeeklyBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'weeklyBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('weeklyBackup').timeBased().everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  // สำรองทันที 1 ชุดเป็นชุดแรก
  var name = weeklyBackup();
  Logger.log('ตั้งสำรองข้อมูลอัตโนมัติทุกวันจันทร์ 07:00-08:00 น. เรียบร้อย (ชุดแรก: ' + name + ')');
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
