/**
 * เว็บ "เปิดหาร" — แสดงรายการของที่แต่ละคนมาหาร จาก Google Sheet
 * และแนบรูปของแต่ละคน (เก็บไฟล์ใน Google Drive)
 *
 * วิธีใช้: เปิดชีต > ส่วนขยาย > Apps Script แล้ววางไฟล์ Code.gs + Index.html
 * รายละเอียดดูใน README.md
 */

const CONFIG = {
  // รหัสแอดมิน (ใช้ตอนแนบ/ลบรูป, แก้ร้าน) อย่างน้อย 4 ตัว — คนเปิดเว็บมองไม่เห็นโค้ดส่วนนี้
  // หรือเว้นว่างแล้วตั้งจากเมนู "เปิดหาร > ตั้งรหัสแอดมิน" ในชีตแทนก็ได้
  ADMIN_PIN: '',
  // เว้นว่างได้ถ้าสคริปต์ผูกกับชีต (ส่วนขยาย > Apps Script) ไม่งั้นใส่ ID จาก URL ของชีต
  SPREADSHEET_ID: '',
  TITLE: 'เปิดหาร',
  // ชีตที่จะแสดง เช่น ['MHA', 'arknights'] — [] = ทุกชีต (ยกเว้นชื่อที่ขึ้นต้นด้วย _)
  SHEETS: [],
  HEADER_ROWS: 1,
  // เลขคอลัมน์ในชีต (A = 1, B = 2, ...)
  COL: {
    item: 1,       // A  ของที่หาร
    owner: 2,      // B  เจ้าของ
    pay: 3,        // C  สถานะจ่าย
    contact: 4,    // D  ช่องทางติดต่อ
    contact2: 5,   // E  ช่องทางติดต่อสำรอง
    shipping: 6,   // F  ส่งแบบ
    slip: 7,       // G  สลิป
    sent: 8,       // H  ส่ง
    note: 9,       // I  note
    code: 10,      // J  โค้ดส่วนตัว
    group: 11,     // K  ชื่อรอบหาร (ผสานเซลล์คร่อมแถวของรอบนั้น)
    // คอลัมน์ทางขวาของ K ทั้งหมด (L, M, N, O, ...) อ่านเป็นข้อมูลของรอบหาร:
    //   ลิงก์ -> ลิงก์ของรอบ, ข้อความ -> รายละเอียดรอบ, รูป -> รูปของรอบ
    // ลิงก์ที่ใส่ในคอลัมน์ K (ใต้ชื่อรอบ) ก็นับเป็นลิงก์ของรอบนั้น
  },
  // ข้อความในคอลัมน์ "ของ" ที่ถือว่าเป็นแถวคั่นรอบ
  SEPARATORS: ['z'],
  // ค่าในคอลัมน์ "เจ้าของ" ที่แปลว่ายังว่าง
  EMPTY_OWNER: ['', 'ว่าง', 'เจ้าของ'],
  // ทุกคนเห็นทุกอย่าง ยกเว้นสลิป (คอลัมน์ G) ที่เห็นเฉพาะแอดมิน
  PUBLIC_NOTE: true,
  PHOTO_SHEET: '_Photos',
  ORDER_SHEET: '_Orders',
  FOLDER_NAME: 'เปิดหาร - รูปภาพ',
  ROW_KEY: 'harnRowId',
  MAX_UPLOAD_BYTES: 10 * 1024 * 1024,
};

const PHOTO_HEADERS = ['photoId', 'ชีต', 'rowId', 'fileId', 'คำอธิบาย', 'อัปโหลดเมื่อ', 'ของ', 'เจ้าของ'];
const ORDER_HEADERS = ['ชีต', 'รอบหาร', 'สั่งที่ร้าน', 'ลิงก์', 'แก้ไขล่าสุด',
  'สถานะ (0-4)', 'วันที่สั่งแล้ว', 'วันที่ของกำลังมา', 'วันที่ถึงไทย', 'วันที่ถึงมือ'];
// ขั้นติดตามสถานะ (ลำดับต้องตรงกับ STEPS ใน Index.html)
const ORDER_STEPS = ['สั่งแล้ว', 'ของกำลังมา', 'ถึงไทย', 'ถึงมือ'];

/* ============================================================
 * Web app + เมนูในชีต
 * ============================================================ */

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.title = CONFIG.TITLE;
  return t.evaluate()
    .setTitle(CONFIG.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('เปิดหาร')
    .addItem('ตั้งรหัสแอดมิน (สำหรับแนบรูป)', 'setAdminPinPrompt')
    .addItem('ลิงก์หน้าเว็บ', 'showWebUrl')
    .addToUi();
}

function setAdminPinPrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('ตั้งรหัสแอดมิน', 'ใช้ตอนแนบ/ลบรูปบนหน้าเว็บ (อย่างน้อย 4 ตัว)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const pin = res.getResponseText().trim();
  if (pin.length < 4) { ui.alert('รหัสต้องมีอย่างน้อย 4 ตัว'); return; }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', pin);
  setup_();
  ui.alert('ตั้งรหัสเรียบร้อย');
}

function showWebUrl() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url || 'ยังไม่ได้ Deploy เป็น Web app (การทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > เว็บแอป)');
}

/** สร้างชีตเก็บรูปและโฟลเดอร์ใน Drive (เรียกซ้ำได้) */
function setup_() {
  photoSheet_();
  orderSheet_();
  photoFolder_();
}

/* ============================================================
 * API ที่หน้าเว็บเรียก
 * ============================================================ */

/** ตรวจรหัสแอดมิน */
function login(pin) {
  return checkPin_(pin);
}

/** ข้อมูลทั้งหมดสำหรับหน้าเว็บ (ใส่รหัสแอดมินมาด้วยจะได้ข้อมูลส่วนตัวเพิ่ม) */
function getData(pin) {
  const isAdmin = pin ? checkPin_(pin) : false;
  const ss = spreadsheet_();
  const photos = readPhotos_();
  const orders = readOrders_();
  const sheets = ss.getSheets()
    .filter(sh => CONFIG.SHEETS.length
      ? CONFIG.SHEETS.indexOf(sh.getName()) >= 0
      : sh.getName().charAt(0) !== '_' && !sh.isSheetHidden())
    .map(sh => readSheet_(sh, isAdmin, photos, orders));
  return { title: CONFIG.TITLE, isAdmin: isAdmin, sheets: sheets };
}

/**
 * อัปโหลดรูป
 * req: { sheet, dataUrl, caption, rowId?, row?, item?, group? }
 *   - แนบให้คน: ส่ง rowId (ถ้ามี) หรือ row + item (ใช้ตรวจว่าแถวไม่เลื่อน)
 *   - แนบเป็นรูปของรอบหาร: ส่ง group
 *   - แนบหลักฐานการสั่งของรอบหาร: ส่ง group + kind: 'order'
 */
function uploadPhoto(pin, req) {
  requirePin_(pin);
  const m = String(req.dataUrl || '').match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
  if (!m) throw new Error('ไฟล์ต้องเป็นรูปภาพ');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > CONFIG.MAX_UPLOAD_BYTES) throw new Error('รูปใหญ่เกิน 10MB');

  const sheet = spreadsheet_().getSheetByName(req.sheet);
  if (!sheet) throw new Error('ไม่พบชีต ' + req.sheet);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let rowId, item = '', owner = '';
    if (req.group) {
      rowId = (req.kind === 'order' ? 'order:' : 'group:') + req.group;
      item = (req.kind === 'order' ? 'หลักฐานการสั่ง ' : '') + req.group;
    } else {
      rowId = req.rowId || ensureRowId_(sheet, Number(req.row), req.item);
      const row = findRowById_(sheet, rowId);
      if (row) {
        item = sheet.getRange(row, CONFIG.COL.item).getDisplayValue();
        owner = sheet.getRange(row, CONFIG.COL.owner).getDisplayValue();
      }
    }

    const ext = m[1].split('/')[1].replace('jpeg', 'jpg').replace(/\W.*$/, '');
    const name = [req.sheet, item, owner].filter(String).join(' - ') + ' ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '.' + ext;
    const file = photoFolder_().createFile(Utilities.newBlob(bytes, m[1], name));
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      // บางโดเมน (Workspace) ห้ามแชร์แบบลิงก์ — รูปจะเห็นเฉพาะคนที่มีสิทธิ์ในไฟล์
    }

    const photo = {
      id: Utilities.getUuid(),
      fileId: file.getId(),
      caption: String(req.caption || '').slice(0, 200),
      at: new Date().toISOString(),
    };
    photoSheet_().appendRow([photo.id, req.sheet, rowId, photo.fileId, photo.caption, new Date(), item, owner]);
    return { rowId: rowId, photo: photo };
  } finally {
    lock.releaseLock();
  }
}

/**
 * บันทึกข้อมูลการสั่งของรอบหาร
 * req: { sheet, group, shop, url }
 */
function saveOrder(pin, req) {
  requirePin_(pin);
  const shop = String(req.shop || '').trim().slice(0, 200);
  const url = String(req.url || '').trim().slice(0, 1000);
  if (url && !/^https?:\/\//i.test(url)) throw new Error('ลิงก์ต้องขึ้นต้นด้วย http:// หรือ https://');
  if (!req.sheet || !req.group) throw new Error('ไม่พบรอบหาร');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = orderSheet_();
    const row = findOrderRow_(sh, req.sheet, req.group);
    sh.getRange(row, 3, 1, 3).setValues([[shop, url, new Date()]]);
    return { shop: shop, url: url };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ตั้งสถานะติดตามของรอบหาร
 * req: { sheet, group, status } — status 0 = ยังไม่เริ่ม, 1-4 = ตาม ORDER_STEPS
 * บันทึกวันที่ของขั้นที่เพิ่งถึง และล้างวันที่ของขั้นที่ถอยกลับ
 */
function saveOrderStatus(pin, req) {
  requirePin_(pin);
  const status = Math.floor(Number(req.status));
  if (!(status >= 0 && status <= ORDER_STEPS.length)) throw new Error('สถานะไม่ถูกต้อง');
  if (!req.sheet || !req.group) throw new Error('ไม่พบรอบหาร');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = orderSheet_();
    const row = findOrderRow_(sh, req.sheet, req.group);
    const dates = sh.getRange(row, 7, 1, ORDER_STEPS.length).getValues()[0];
    const now = new Date();
    for (let k = 0; k < ORDER_STEPS.length; k++) {
      if (k < status) { if (!(dates[k] instanceof Date)) dates[k] = now; }
      else dates[k] = '';
    }
    sh.getRange(row, 6, 1, 1 + ORDER_STEPS.length).setValues([[status].concat(dates)]);
    return { status: status, dates: dates.map(d => (d instanceof Date ? d.toISOString() : '')) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * แก้วันที่ของขั้นสถานะ (กรณีลืมกดตอนนั้น หรือสั่งไว้นานแล้ว)
 * req: { sheet, group, step (0-3), date: 'yyyy-mm-dd' }
 */
function saveOrderDate(pin, req) {
  requirePin_(pin);
  const step = Math.floor(Number(req.step));
  if (!(step >= 0 && step < ORDER_STEPS.length)) throw new Error('ขั้นไม่ถูกต้อง');
  const m = String(req.date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('วันที่ไม่ถูกต้อง');
  // เที่ยงวัน เพื่อไม่ให้วันเลื่อนเพราะโซนเวลา
  const date = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12);
  if (!req.sheet || !req.group) throw new Error('ไม่พบรอบหาร');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = orderSheet_();
    const row = findOrderRow_(sh, req.sheet, req.group);
    const status = Number(sh.getRange(row, 6).getValue()) || 0;
    if (step >= status) throw new Error('ขั้นนี้ยังไม่ถึง ตั้งสถานะก่อนแล้วค่อยแก้วันที่');
    sh.getRange(row, 7 + step).setValue(date);
    const dates = sh.getRange(row, 7, 1, ORDER_STEPS.length).getValues()[0];
    return { dates: dates.map(d => (d instanceof Date ? d.toISOString() : '')) };
  } finally {
    lock.releaseLock();
  }
}

/** แถวของรอบหารในชีต _Orders (ไม่มีจะสร้างให้) — เรียกภายใต้ lock */
function findOrderRow_(sh, sheetName, group) {
  const last = sh.getLastRow();
  const keys = last >= 2 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) === sheetName && String(keys[i][1]) === group) return i + 2;
  }
  sh.appendRow([sheetName, group, '', '', new Date()]);
  return sh.getLastRow();
}

function deletePhoto(pin, photoId) {
  requirePin_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = photoSheet_();
    const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 2).getValues();
    for (let i = ids.length - 1; i >= 1; i--) {
      if (ids[i][0] !== photoId) continue;
      const fileId = sh.getRange(i + 1, 4).getValue();
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* ไฟล์ถูกลบไปแล้ว */ }
      sh.deleteRow(i + 1);
      return true;
    }
    return false;
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * อ่านชีต
 * ============================================================ */

function readSheet_(sheet, isAdmin, photos, orders) {
  const C = CONFIG.COL;
  const name = sheet.getName();
  const out = { name: name, groups: [] };
  const lastRow = sheet.getLastRow();
  const first = CONFIG.HEADER_ROWS + 1;
  if (lastRow < first) return out;

  const n = lastRow - first + 1;
  const width = Math.min(
    Math.max(sheet.getLastColumn(), Math.max.apply(null, Object.keys(C).map(k => C[k]))),
    sheet.getMaxColumns());
  const range = sheet.getRange(first, 1, n, width);
  const disp = range.getDisplayValues();
  const vals = range.getValues();
  let rich = null;
  try { rich = range.getRichTextValues(); } catch (e) { /* ปล่อยว่าง */ }

  const cell = (i, col) => (col <= width ? String(disp[i][col - 1] || '').trim() : '');
  const link = (i, col) => {
    if (col > width) return '';
    const r = rich && rich[i][col - 1];
    let url = '';
    if (r) {
      url = r.getLinkUrl() || '';
      if (!url) {
        const runs = r.getRuns();
        for (let k = 0; k < runs.length && !url; k++) url = runs[k].getLinkUrl() || '';
      }
    }
    const text = cell(i, col);
    if (!url && /^https?:\/\//i.test(text)) url = text;
    return url;
  };

  // ชื่อรอบหารจากเซลล์ที่ผสานในคอลัมน์ group
  const groupOfRow = {};
  if (C.group <= width) {
    sheet.getRange(first, C.group, n, 1).getMergedRanges().forEach(mr => {
      const g = mr.getDisplayValue().trim();
      for (let r = mr.getRow(); r <= mr.getLastRow(); r++) groupOfRow[r] = g;
    });
  }

  // รูปของรอบหาร: รูปที่วางในเซลล์ (in-cell) หรือรูปลอย (over-grid) ที่ยึดในคอลัมน์ทางขวาของ K
  const coverByRow = {};
  const addCover = (row, url) => { if (url) (coverByRow[row] = coverByRow[row] || []).push(url); };
  try {
    sheet.getImages().forEach(img => {
      const a = img.getAnchorCell();
      if (a.getColumn() <= C.group) return;
      let url = '';
      try { url = img.getUrl() || ''; } catch (e) { /* รูปที่อัปโหลดเองมักไม่มี URL */ }
      addCover(a.getRow(), url);
    });
  } catch (e) { /* ปล่อยว่าง */ }

  const rowIds = rowIdMap_(sheet);
  const isSep = s => CONFIG.SEPARATORS.indexOf(s.toLowerCase()) >= 0;
  const key = s => s.toLowerCase().replace(/\s+/g, '');
  const groups = [];
  let current = null;

  const groupFor = gName => {
    const id = gName || '(ไม่มีชื่อรอบ)';
    if (current && current.name === id) return current;
    // ชื่อซ้ำแต่ถูกคั่นด้วยแถว z ให้แยกเป็นคนละรอบ
    const g = {
      name: id, covers: [], links: [], info: [], rows: [],
      photos: photos[name + '|group:' + id] || [],
      order: Object.assign({ shop: '', url: '', status: 0, dates: [] }, orders[name + '|' + id], {
        photos: photos[name + '|order:' + id] || [],
      }),
    };
    groups.push(g);
    current = g;
    return g;
  };

  for (let i = 0; i < n; i++) {
    const row = first + i;
    const item = cell(i, C.item);
    const ownerRaw = cell(i, C.owner);

    if (isSep(item)) { current = null; continue; }

    // ลิงก์ที่อยู่ในคอลัมน์ K ไม่ใช่ชื่อรอบใหม่ แต่เป็นลิงก์ของรอบปัจจุบัน
    const groupCell = cell(i, C.group);
    const groupCellIsLink = !groupOfRow[row] && !!(link(i, C.group) || /^https?:\/\//i.test(groupCell));
    const gName = groupOfRow[row] || (groupCellIsLink ? '' : groupCell) || (current ? current.name : '');

    // ข้อมูลรอบหาร (ลิงก์/ข้อความ/รูป) จากคอลัมน์ K (ถ้าเป็นลิงก์) และทุกคอลัมน์ทางขวา
    const extras = { links: [], info: [] };
    for (let col = groupCellIsLink ? C.group : C.group + 1; col <= width; col++) {
      const url = link(i, col);
      const text = cell(i, col);
      if (url) extras.links.push(url);
      else if (text) extras.info.push(text);
      addCover(row, imageUrl_(vals[i][col - 1]));
    }

    const hasAnything = item || ownerRaw || groupCell || extras.links.length || extras.info.length || coverByRow[row];
    if (!hasAnything) continue;
    const g = groupFor(gName);
    extras.links.forEach(u => { if (g.links.indexOf(u) < 0) g.links.push(u); });
    extras.info.forEach(t => { if (g.info.indexOf(t) < 0) g.info.push(t); });
    (coverByRow[row] || []).forEach(u => { if (g.covers.indexOf(u) < 0) g.covers.push(u); });

    // แถวหัวข้อ (ชื่อของ = ชื่อรอบ และไม่มีเจ้าของ) ไม่นับเป็นคน
    if (!item || (key(item) === key(g.name) && !ownerRaw)) continue;

    // ว่าง = ไม่มีเจ้าของ และยังไม่มีการจ่าย/ส่ง (ถ้าจ่ายแล้วแต่ลืมใส่ชื่อ ไม่นับว่าว่าง)
    const noOwner = CONFIG.EMPTY_OWNER.indexOf(ownerRaw) >= 0;
    const available = noOwner && !cell(i, C.pay) && !cell(i, C.sent);
    const rowId = rowIds[row] || '';
    const r = {
      row: row,
      rowId: rowId,
      item: item,
      owner: noOwner ? '' : ownerRaw,
      available: available,
      pay: cell(i, C.pay),
      shipping: cell(i, C.shipping),
      sent: cell(i, C.sent),
      note: CONFIG.PUBLIC_NOTE || isAdmin ? cell(i, C.note) : '',
      photos: rowId ? photos[name + '|' + rowId] || [] : [],
    };
    r.contacts = [C.contact, C.contact2]
      .map(col => ({ text: cell(i, col), url: link(i, col) }))
      .filter(c => c.text || c.url);
    r.code = cell(i, C.code);
    // สลิปเห็นเฉพาะแอดมิน
    if (isAdmin) {
      r.slipImage = C.slip <= width ? imageUrl_(vals[i][C.slip - 1]) : '';
      r.slipText = r.slipImage ? '' : cell(i, C.slip);
    }
    g.rows.push(r);
  }

  // ข้ามรอบที่ไม่มีอะไรเลย
  out.groups = groups.filter(g => g.rows.length || g.covers.length || g.photos.length ||
    g.order.shop || g.order.url || g.order.status || g.order.photos.length);
  return out;
}

/** URL ของรูปที่วางในเซลล์ (CellImage) */
function imageUrl_(v) {
  if (v && typeof v === 'object' && typeof v.getContentUrl === 'function') {
    try { return v.getContentUrl() || ''; } catch (e) { return ''; }
  }
  return '';
}

/* ============================================================
 * รหัสแถว — ผูกกับแถวด้วย Developer Metadata
 * แทรก/ลบ/ย้ายแถวในชีตแล้วรูปยังติดอยู่กับคนเดิม
 * ============================================================ */

function rowIdMap_(sheet) {
  const map = {};
  sheet.createDeveloperMetadataFinder()
    .withKey(CONFIG.ROW_KEY)
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.ROW)
    .find()
    .forEach(md => { map[md.getLocation().getRow().getRow()] = md.getValue(); });
  return map;
}

function findRowById_(sheet, rowId) {
  const map = rowIdMap_(sheet);
  for (const row in map) if (map[row] === rowId) return Number(row);
  return 0;
}

function ensureRowId_(sheet, row, expectedItem) {
  if (!(row > CONFIG.HEADER_ROWS)) throw new Error('แถวไม่ถูกต้อง');
  const actual = sheet.getRange(row, CONFIG.COL.item).getDisplayValue().trim();
  if (actual !== String(expectedItem || '').trim()) {
    throw new Error('ข้อมูลในชีตเปลี่ยนไปแล้ว กรุณารีเฟรชหน้าเว็บแล้วลองใหม่');
  }
  const existing = rowIdMap_(sheet)[row];
  if (existing) return existing;
  const id = Utilities.getUuid().slice(0, 12);
  sheet.getRange(row + ':' + row).addDeveloperMetadata(CONFIG.ROW_KEY, id);
  return id;
}

/* ============================================================
 * ชีตเก็บรูป / โฟลเดอร์ Drive
 * ============================================================ */

function spreadsheet_() {
  return CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function photoSheet_() {
  const ss = spreadsheet_();
  let sh = ss.getSheetByName(CONFIG.PHOTO_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.PHOTO_SHEET);
    sh.getRange(1, 1, 1, PHOTO_HEADERS.length).setValues([PHOTO_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function orderSheet_() {
  const ss = spreadsheet_();
  let sh = ss.getSheetByName(CONFIG.ORDER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.ORDER_SHEET);
    sh.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.hideSheet();
  } else if (sh.getRange(1, ORDER_HEADERS.length).getValue() === '') {
    // ชีตจากเวอร์ชันก่อน: เติมหัวคอลัมน์สถานะที่เพิ่มมา
    sh.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS]).setFontWeight('bold');
  }
  return sh;
}

/** { 'ชีต|รอบหาร': {shop, url, status, dates} } */
function readOrders_() {
  const sh = spreadsheet_().getSheetByName(CONFIG.ORDER_SHEET);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  const width = Math.min(ORDER_HEADERS.length, sh.getMaxColumns());
  sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues().forEach(v => {
    if (!v[0] || !v[1]) return;
    const status = Math.max(0, Math.min(ORDER_STEPS.length, Math.floor(Number(v[5]) || 0)));
    const dates = [];
    for (let k = 0; k < ORDER_STEPS.length; k++) {
      const d = v[6 + k];
      dates.push(d instanceof Date ? d.toISOString() : '');
    }
    map[v[0] + '|' + v[1]] = { shop: String(v[2] || ''), url: String(v[3] || ''), status: status, dates: dates };
  });
  return map;
}

/** { 'ชีต|rowId': [ {id, fileId, caption, at} ] } */
function readPhotos_() {
  const sh = spreadsheet_().getSheetByName(CONFIG.PHOTO_SHEET);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(v => {
    if (!v[0] || !v[3]) return;
    const k = v[1] + '|' + v[2];
    (map[k] = map[k] || []).push({
      id: String(v[0]),
      fileId: String(v[3]),
      caption: String(v[4] || ''),
      at: v[5] instanceof Date ? v[5].toISOString() : String(v[5] || ''),
    });
  });
  return map;
}

function photoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* โฟลเดอร์ถูกลบ สร้างใหม่ */ }
  }
  const folder = DriveApp.createFolder(CONFIG.FOLDER_NAME);
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

/* ============================================================
 * รหัสแอดมิน (กันเดารหัส: ผิดเกิน 10 ครั้งใน 15 นาที ล็อก)
 * ============================================================ */

function checkPin_(pin) {
  const real = String(CONFIG.ADMIN_PIN || '').trim() ||
    PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!real || real.length < 4) {
    throw new Error('ยังไม่ได้ตั้งรหัสแอดมิน: ใส่ใน ADMIN_PIN บรรทัดบนสุดของ Code.gs แล้ว Deploy เวอร์ชันใหม่');
  }
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinFails') || 0);
  if (fails >= 10) throw new Error('ใส่รหัสผิดหลายครั้ง กรุณารอ 15 นาที');
  if (String(pin || '') === real) return true;
  cache.put('pinFails', String(fails + 1), 15 * 60);
  return false;
}

function requirePin_(pin) {
  if (!checkPin_(pin)) throw new Error('รหัสแอดมินไม่ถูกต้อง');
}
