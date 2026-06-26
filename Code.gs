/**
 * ระบบ CRM Am Home Car (Backend) - Optimized Version
 * ปรับปรุงความเร็วด้วย:
 * 1. CacheService - Cache ข้อมูลลูกค้า 5 นาที
 * 2. TextFinder แทน loop หาแถว
 * 3. Batch getRange ครั้งเดียวต่อฟังก์ชัน
 * 4. SpreadsheetApp.flush() หลัง write
 * 5. ส่ง record กลับทันที ไม่รอ reload ใหม่ทั้งหมด
 */

const CACHE_TTL = 300; // 5 นาที

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('CRM Am Home Car')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ===== CACHE HELPERS =====
function invalidateCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('cust_admin');
  } catch(e) {}
}

function getCached(key) {
  try { const r = CacheService.getScriptCache().get(key); return r ? JSON.parse(r) : null; } catch(e) { return null; }
}

function setCache(key, data) {
  try {
    const str = JSON.stringify(data);
    if (str.length < 90000) CacheService.getScriptCache().put(key, str, CACHE_TTL);
  } catch(e) {}
}

// ===== DATE FORMATTER =====
function formatDate(val) {
  if (val instanceof Date) {
    return val.getFullYear() + '-' + String(val.getMonth()+1).padStart(2,'0') + '-' + String(val.getDate()).padStart(2,'0');
  }
  return val ? val.toString() : '';
}

// ===== หาแถวด้วย TextFinder (เร็วกว่า loop มาก) =====
function findRow(sheet, id) {
  const cell = sheet.createTextFinder(id).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : -1;
}

// ===== สร้าง ID ใหม่จาก array ที่มีอยู่ =====
function genID(prefix, ids) {
  let max = 0;
  ids.forEach(id => {
    if (id && id.toString().startsWith(prefix)) {
      const n = parseInt(id.toString().replace(prefix, '')) || 0;
      if (n > max) max = n;
    }
  });
  return prefix + (max + 1).toString().padStart(3, '0');
}

// ===== LOGIN =====
function login(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const last  = sheet.getLastRow();
  if (last < 2) return null;
  const data = sheet.getRange(2, 1, last - 1, 5).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString() === username && data[i][1].toString() === password) {
      return { username: data[i][0], name: data[i][2], role: data[i][3], img: data[i][4] };
    }
  }
  return null;
}

// ===== ดึงข้อมูลลูกค้า (พร้อม Cache) =====
function getCustomerData(user) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  const custSheet = ss.getSheetByName('Customers');
  const isAdmin   = user.role.toLowerCase() === 'admin';

  // userList เล็กมาก ดึงตรงเลย
  const lastU   = userSheet.getLastRow();
  const allUsers = lastU >= 2 ? userSheet.getRange(2, 1, lastU - 1, 1).getValues().flat().filter(u => u !== '') : [];

  // ลอง Cache (เฉพาะ admin cache รวม ส่วน user แต่ละคน cache แยก)
  const cacheKey = isAdmin ? 'cust_admin' : 'cust_' + user.username;
  const cached   = getCached(cacheKey);
  if (cached) return { customers: cached, userList: allUsers };

  // ดึงจาก Sheet batch เดียว
  const lastC = custSheet.getLastRow();
  if (lastC < 2) return { customers: [], userList: allUsers };

  const raw = custSheet.getRange(2, 1, lastC - 1, 7).getValues();
  const processed = raw.map(r => { r[1] = r[1].toString(); r[6] = formatDate(r[6]); return r; });
  const result    = isAdmin ? processed : processed.filter(r => r[4] === user.username);

  setCache(cacheKey, result);
  return { customers: result, userList: allUsers };
}

// ===== บันทึกข้อมูล =====
function processCustomer(action, data, role) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
  invalidateCache();

  if (action === 'add') {
    const lastRow   = sheet.getLastRow();
    const existingIds = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
    const now     = new Date();
    const prefix  = 'PD' + now.getFullYear().toString().slice(-2) + String(now.getMonth()+1).padStart(2,'0');
    const newID   = genID(prefix, existingIds);
    const newRow  = [newID, data.phone.toString(), data.name, data.status, data.assignedUser, data.remark, data.date || ''];
    sheet.appendRow(newRow);
    sheet.getRange(sheet.getLastRow(), 2).setNumberFormat('@');
    SpreadsheetApp.flush();
    writeLog('add', newID, data.assignedUser, 'เพิ่มลูกค้าใหม่', role);
    return { msg: 'เพิ่มลูกค้า ' + newID + ' สำเร็จ', newRecord: newRow };
  }

  if (action === 'delete') {
    const rowNum = findRow(sheet, data.id);
    if (rowNum === -1) return { msg: 'ไม่พบข้อมูลที่ต้องการลบ' };
    sheet.deleteRow(rowNum);
    SpreadsheetApp.flush();
    writeLog('delete', data.id, '', 'ลบลูกค้า', role);
    return { msg: 'ลบลูกค้า ' + data.id + ' เรียบร้อย', deletedId: data.id };
  }

  // update
  const rowNum = findRow(sheet, data.id);
  if (rowNum === -1) return { msg: 'ไม่พบข้อมูลที่ต้องการแก้ไข' };
  const oldStatus = sheet.getRange(rowNum, 4).getValue();

  if (role.toLowerCase() === 'admin') {
    sheet.getRange(rowNum, 2, 1, 6).setValues([[data.phone.toString(), data.name, data.status, data.assignedUser, data.remark, data.date || '']]);
    sheet.getRange(rowNum, 2).setNumberFormat('@');
  } else {
    sheet.getRange(rowNum, 4, 1, 3).setValues([[data.status, data.assignedUser, data.remark]]);
  }
  SpreadsheetApp.flush();

  const changes = [];
  if (oldStatus !== data.status) changes.push('สถานะ: ' + oldStatus + ' → ' + data.status);
  writeLog('edit', data.id, data.assignedUser, changes.length ? changes.join(', ') : 'แก้ไขข้อมูล', role);
  return { msg: 'อัปเดตข้อมูล ' + data.id + ' เรียบร้อย', updatedData: data };
}

// ===== โยนเคสไปเดือนอื่น =====
function moveToMonth(caseId, targetYY, targetMM, username) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
  invalidateCache();

  const rowNum = findRow(sheet, caseId);
  if (rowNum === -1) return { success: false, msg: 'ไม่พบเคสนี้' };

  const lastRow    = sheet.getLastRow();
  const oldData    = sheet.getRange(rowNum, 1, 1, 7).getValues()[0];
  const allIds     = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
  const newPrefix  = 'PD' + targetYY + targetMM;
  const newID      = genID(newPrefix, allIds);
  const newDate    = '20' + targetYY + '-' + targetMM + '-01';

  sheet.appendRow([newID, oldData[1]?.toString() || '', oldData[2], oldData[3], oldData[4], oldData[5], newDate]);
  sheet.getRange(sheet.getLastRow(), 2).setNumberFormat('@');
  sheet.deleteRow(rowNum);
  SpreadsheetApp.flush();

  writeLog('move_month', caseId + ' → ' + newID, username, 'โยนเคสไปเดือน ' + targetMM + '/' + targetYY, 'user');
  return { success: true, newId: newID, oldId: caseId, msg: 'โยนเคส ' + caseId + ' → ' + newID + ' สำเร็จ!' };
}

// ===== Log =====
function writeLog(action, customerId, assignedUser, detail, role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('Log');
  if (!logSheet) {
    logSheet = ss.insertSheet('Log');
    logSheet.appendRow(['วันที่/เวลา','Action','ID ลูกค้า','ผู้ดูแล','รายละเอียด','Role']);
  }
  const now = new Date();
  const d   = now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0')+'-'+String(now.getDate()).padStart(2,'0')
    +' '+String(now.getHours()).padStart(2,'0')+':'+String(now.getMinutes()).padStart(2,'0');
  logSheet.appendRow([d, action, customerId, assignedUser, detail, role]);
}

function getLogs() {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  if (!logSheet) return [];
  const last = logSheet.getLastRow();
  if (last < 2) return [];
  return logSheet.getRange(2, 1, last - 1, 6).getValues().reverse();
}

// ===== Users =====
function getUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const last  = sheet.getLastRow();
  return last >= 2 ? sheet.getRange(2, 1, last - 1, 5).getValues() : [];
}

function processUserAccount(data) {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const last   = sheet.getLastRow();
  const names  = last >= 2 ? sheet.getRange(2, 1, last - 1, 1).getValues().flat() : [];
  const isEdit = data.originalUsername && data.originalUsername !== '';

  if (!isEdit) {
    if (names.some(u => u.toString() === data.username)) return 'มี Username นี้อยู่แล้ว';
    sheet.appendRow([data.username, data.password, data.name, data.role, data.img || '']);
    return 'เพิ่ม User ' + data.username + ' สำเร็จ';
  }
  const rowNum = findRow(sheet, data.originalUsername);
  if (rowNum === -1) return 'ไม่พบ User ที่ต้องการแก้ไข';
  sheet.getRange(rowNum, 1, 1, 5).setValues([[data.username, data.password, data.name, data.role, data.img || '']]);
  return 'อัปเดต User ' + data.username + ' เรียบร้อย';
}

// ===== รายงานรายเดือน =====
function getMonthlyReport(year, month) {
  const custSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
  const last = custSheet.getLastRow();
  if (last < 2) return { report: {}, total: 0 };
  const all = custSheet.getRange(2, 1, last - 1, 7).getValues();
  const prefix   = 'PD' + year + month;
  const filtered = month === ''
    ? all.filter(r => r[0] && r[0].toString().substring(2,4) === year)
    : all.filter(r => r[0] && r[0].toString().startsWith(prefix));
  const report = {};
  filtered.forEach(r => {
    const sale = r[4] || 'ไม่ระบุ', status = r[3] || 'ไม่ระบุ';
    if (!report[sale]) report[sale] = { total:0, ไปต่อได้:0, ไปต่อไม่ได้:0, ปิดการขาย:0, ติดตาม:0, จองแล้ว:0, จัดไฟแนนซ์:0, ไม่ให้ความร่วมมือ:0 };
    report[sale].total++;
    if (report[sale][status] !== undefined) report[sale][status]++; else report[sale][status] = 1;
  });
  return { report, total: filtered.length };
}

// ===== MARKETPLACE =====
function initMarketplaceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let mp = ss.getSheetByName('Marketplace');
  if (!mp) {
    mp = ss.insertSheet('Marketplace');
    mp.appendRow(['ID ลูกค้า','เบอร์โทร','ชื่อ','Remark','วันที่เพิ่มในตลาด','ผู้เพิ่มเข้าตลาด']);
    mp.setFrozenRows(1);
  }
  return mp;
}

function getMarketplaceData() {
  initMarketplaceSheet();
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const mp   = ss.getSheetByName('Marketplace');
  const cust = ss.getSheetByName('Customers');
  const mpLast   = mp.getLastRow();
  const custLast = cust.getLastRow();
  if (mpLast < 2) return [];
  const mpData   = mp.getRange(2, 1, mpLast - 1, 6).getValues();
  const custData = custLast >= 2 ? cust.getRange(2, 1, custLast - 1, 7).getValues() : [];
  const assigned = new Set(custData.filter(r => r[4] && r[4].toString().trim() !== '').map(r => r[0]?.toString()));
  return mpData
    .filter(r => r[0] && !assigned.has(r[0].toString()))
    .map(r => ({ id: r[0]?.toString()||'', phone: r[1]?.toString()||'', name: r[2]?.toString()||'', remark: r[3]?.toString()||'', date: r[4]?.toString()||'', addedBy: r[5]?.toString()||'' }));
}

function acceptMarketplaceCase(caseId, username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initMarketplaceSheet();
  invalidateCache();
  const cust = ss.getSheetByName('Customers');
  const mp   = ss.getSheetByName('Marketplace');
  const mpRow = findRow(mp, caseId);
  if (mpRow === -1) return 'ไม่พบเคสนี้ในตลาด';
  const mpData    = mp.getRange(mpRow, 1, 1, 6).getValues()[0];
  const custRow   = findRow(cust, caseId);
  if (custRow === -1) {
    cust.appendRow([caseId, mpData[1]?.toString()||'', mpData[2], 'ติดตาม', username, mpData[3]?.toString()||'', new Date().toISOString().split('T')[0]]);
    cust.getRange(cust.getLastRow(), 2).setNumberFormat('@');
  } else {
    cust.getRange(custRow, 5).setValue(username);
  }
  mp.deleteRow(mpRow);
  SpreadsheetApp.flush();
  writeLog('marketplace_accept', caseId, username, 'รับเคสจากตลาด', 'user');
  return 'รับเคส ' + caseId + ' สำเร็จ! เคสจะไปอยู่ในของคุณแล้ว';
}

function sendToMarketplace(caseId, username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initMarketplaceSheet();
  invalidateCache();
  const cust    = ss.getSheetByName('Customers');
  const mp      = ss.getSheetByName('Marketplace');
  const custRow = findRow(cust, caseId);
  if (custRow === -1) return 'ไม่พบเคสนี้';
  if (findRow(mp, caseId) !== -1) return 'เคสนี้อยู่ในตลาดแล้ว';
  const custData = cust.getRange(custRow, 1, 1, 7).getValues()[0];
  mp.appendRow([caseId, custData[1]?.toString()||'', custData[2], custData[5]?.toString()||'', new Date().toISOString().split('T')[0], username]);
  cust.getRange(custRow, 5).setValue('');
  SpreadsheetApp.flush();
  writeLog('marketplace_send', caseId, username, 'ส่งเคสเข้าตลาด', 'user');
  return 'ส่งเคส ' + caseId + ' เข้าตลาดสำเร็จ!';
}
