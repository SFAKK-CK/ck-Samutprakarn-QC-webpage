// ============================================================
// CK QC System — Factory 2 (Samut Prakarn) Backend
// Google Apps Script — doPost HTTP router
// ============================================================

const CENTRAL_FOLDER_ID = '1-Js1TzO2tTdxHkAM3cods3iWyibybHEr';
const GM_PASSWORD = '1971';
const SPREADSHEET_NAME = 'CK_SP_QC_Log';

// ===== EXECUTION-LEVEL CACHE =====
let _cachedSS     = null;
let _cachedFolder = null;

// ===== HELPER: User-friendly lock acquisition =====
function tryWaitLock(lock, ms) {
  try {
    lock.waitLock(ms);
    return { ok: true };
  } catch(e) {
    Logger.log('Lock acquisition failed: ' + e);
    return {
      ok: false,
      message: '⏳ ระบบกำลังบันทึกข้อมูลของผู้อื่นอยู่ — กรุณารอสักครู่แล้วลองใหม่\nSystem busy with another save — please wait and retry'
    };
  }
}

// ===== HELPER: Idempotency cache (10-minute TTL) =====
function checkClientNonce(nonce) {
  if (!nonce) return null;
  try {
    const cached = CacheService.getScriptCache().get('nonce_' + nonce);
    return cached ? JSON.parse(cached) : null;
  } catch(e) { return null; }
}

function storeClientNonce(nonce, result) {
  if (!nonce || !result || !result.success) return;
  try {
    CacheService.getScriptCache().put('nonce_' + nonce, JSON.stringify(result), 600);
  } catch(e) { Logger.log('Nonce store failed: ' + e); }
}

// ===== HELPER: Rollback saved photos when sheet update fails =====
function cleanupSavedFiles(savedFiles) {
  if (!savedFiles || !savedFiles.length) return;
  savedFiles.forEach(function(f) {
    try { f.setTrashed(true); } catch(e) { Logger.log('Cleanup failed: ' + e); }
  });
}

const CONFIG = {
  TIMEZONE: 'Asia/Bangkok',
  PRODUCT_TYPES: [
    'HILSA','ROHU','CATLA','MRIGAL','BACHA','SINGI','TANGRA','TENGRA',
    'PABDA','KORAL','BATA','BARRAMUNDI','PANGUS','MAGUR','SHINGHI',
    'BUAL','BOAL','SOL','SOLE','RITA','AIR','AYRE','CHITAL','TILAPIA',
    'PUTI','SCAMPI','PRAWN','SHRIMP','BLACK TIGER','VANNAMEI','SQUID',
    'CUTTLEFISH','OCTOPUS','RUHI DRIED','CATLA DRIED','TENGRA DRIED',
    'LOITTA DRIED','CHINGRI DRIED','SURMA DRIED','CHAPILA DRIED',
    'SILVER CARP','GRASS CARP','COMMON CARP','MIRROR CARP','CARFU','DRY SHRIMP','DRY KESKI','KHOLSHA',
    'GREEN  CHILLI','RED CHILLI','E-MOLA','OTHER'
  ],
  SIZES: [
    '100G','200G','240G','250G','100-200G','200-300G','300-500G','500G','400-600G','600-800G','800G',
    '1KG','1.5KG','2KG','3KG','4KG','5KG',
    '6KG','7KG','8KG','9KG','10KG','11KG','12KG','13KG','14KG',
    '15KG','15KG UP','OTHER'
  ],
  PROCESSING_TYPES: [
    'WR','WS','G','GS','GGS','CLN','FC','BLK','H/L','H/O',
    'SO','BG','TRAY','IQF BLK','H/L IQF','RFC',
    'STEAK','STEAK WITH SCALE','FILLET','DRY','EEG','IQF','CHILLI','Other'
  ],
  SHEETS: {
    ORDER_MASTER:     'Order Master',
    PACKING:          'Packing Log',
    COLD_STORAGE:     'Cold Storage Log',
    LOADING_MASTER:   'Loading Master',
    LOADING_LINES:    'Loading Lines',
    EMPLOYEES:        'Employees',
    CONTAINER_PHOTOS: 'Container Photos',
  }
};

// ── Utility ──────────────────────────────────────────────────

function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return HtmlService.createHtmlOutput(
    '<div style="display:flex;align-items:center;justify-content:center;min-height:100vh;font-family:Sarabun,sans-serif;text-align:center;padding:20px">' +
    '<div>' +
    '<p style="font-size:20px;font-weight:bold;margin-bottom:16px">🐟 CK Frozen Fish and Food</p>' +
    '<p style="font-size:16px;margin-bottom:20px">ระบบได้ย้ายไปที่ลิงก์ใหม่แล้ว<br>กรุณาเข้าใช้งานที่ลิงก์ด้านล่าง</p>' +
    '<a href="https://sfakk-ck.github.io/ck-Samutprakarn-QC-webpage/" target="_top" style="font-size:18px;color:#2563eb;font-weight:bold">👉 เข้าสู่ระบบ QC คลิกที่นี่</a>' +
    '</div></div>'
  ).setTitle('CK QC System');
}
// ── Config ────────────────────────────────────────────────────

function getConfig() {
  return {
    productTypes:    CONFIG.PRODUCT_TYPES,
    processingTypes: CONFIG.PROCESSING_TYPES,
    sizes:           CONFIG.SIZES
  };
}

// ── Spreadsheet helpers ───────────────────────────────────────

function makeAccessible(item) {
  try {
    item.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return true;
  } catch(e) {
    Logger.log('⚠️ makeAccessible FAILED for ' + item.getName() + ': ' + e);
    return false;
  }
}

function getRootFolder() {
  if (_cachedFolder) return _cachedFolder;
  const folder = DriveApp.getFolderById(CENTRAL_FOLDER_ID);
  makeAccessible(folder);
  _cachedFolder = folder;
  return _cachedFolder;
}

function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  const f = parent.createFolder(name);
  makeAccessible(f);
  return f;
}

function savePhoto(folder, photoData, fileName) {
  const b64 = photoData.split(',')[1];
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', fileName);
  const file = folder.createFile(blob);
  makeAccessible(file);
  return file;
}

function getOrCreateSpreadsheet() {
  if (_cachedSS) return _cachedSS;
  const root = getRootFolder();
  const files = root.getFilesByName(SPREADSHEET_NAME);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    const ssFile = DriveApp.getFileById(ss.getId());
    root.addFile(ssFile);
    DriveApp.getRootFolder().removeFile(ssFile);
    const sheetNames = Object.values(CONFIG.SHEETS);
    sheetNames.forEach((name, i) => {
      if (i === 0) ss.getActiveSheet().setName(name);
      else ss.insertSheet(name);
    });
  }
  _cachedSS = ss;
  return ss;
}

function getSheet(sheetName) {
  return getOrCreateSpreadsheet().getSheetByName(sheetName);
}

// ── Date helpers ──────────────────────────────────────────────

function getTodayBangkok() {
  const now = new Date();
  const bkk = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  return bkk;
}

function getYearMonth() {
  const now = new Date();
  return {
    year:  Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy'),
    month: Utilities.formatDate(now, CONFIG.TIMEZONE, 'MMMM')
  };
}

// ── Employee validation ───────────────────────────────────────

function validateStaff(staffId) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    if (!sheet || sheet.getLastRow() < 2) {
      return { valid: false, message: '❌ ไม่พบข้อมูลพนักงาน — กรุณาติดต่อผู้ดูแลระบบ\nEmployees sheet not found — contact admin' };
    }
    const data  = sheet.getDataRange().getValues();
    const idKey = String(staffId).trim();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === idKey) {
        const name       = String(data[i][1] || '').trim();
        const department = String(data[i][2] || '').trim();
        return {
          valid: true, staffId: idKey, name: name, department: department,
          message: 'ยินดีต้อนรับ ' + name + '! 👋'
        };
      }
    }
    return { valid: false, message: '❌ ไม่พบรหัสพนักงาน ' + staffId + '\nStaff ID not found — please contact your supervisor' };
  } catch(e) {
    Logger.log('validateStaff error: ' + e);
    return { valid: false, message: '❌ เกิดข้อผิดพลาด: ' + e.toString() };
  }
}

// ── Order folder helpers ──────────────────────────────────────

function getOrCreateOrderFolder(orderName) {
  const root = getRootFolder();
  const { year, month } = getYearMonth();
  const yearFolder  = getOrCreateFolder(root, year);
  const monthFolder = getOrCreateFolder(yearFolder, month);
  return getOrCreateFolder(monthFolder, orderName);
}

function getOrCreateStepFolder(orderFolder, stepName) {
  return getOrCreateFolder(orderFolder, stepName);
}

// ── Step 3 folder layout ──────────────────────────────────────
// Current layout — three step folders per order:
//     <order>/1-Packing
//     <order>/2-Cold Storage QC
//     <order>/3-Loading/Loading Lines
//     <order>/3-Loading/Container Photos
//
// Legacy layout — orders that already existed before this change:
//     <order>/Step3
//     <order>/Container
//
// An order that already has a top-level 'Container' or 'Step3' folder keeps the
// legacy layout for the rest of its life. That way an order which was part-way
// through loading when this version was deployed never ends up with its photos
// split across two different folder shapes. Brand-new orders always get 3-Loading.
//
// Once every legacy order is closed out, isLegacyLoadingLayout() can be deleted
// and both helpers reduced to the 3-Loading branch.
function isLegacyLoadingLayout(orderFolder) {
  return orderFolder.getFoldersByName('Container').hasNext() ||
         orderFolder.getFoldersByName('Step3').hasNext();
}

// The '3-Loading' wrapper. Returns null for legacy orders, which have no wrapper.
function getLoadingRoot(orderFolder) {
  if (isLegacyLoadingLayout(orderFolder)) return null;
  return getOrCreateFolder(orderFolder, '3-Loading');
}

// Destination for the Line-* folders moved out of _Pending on Complete.
function getLinesFolder(orderFolder) {
  const root = getLoadingRoot(orderFolder);
  return root ? getOrCreateFolder(root, 'Loading Lines')
              : getOrCreateFolder(orderFolder, 'Step3');
}

// Where container photos are written. Unchanged behaviour: still saved straight
// to the order folder at capture time, NOT staged through _Pending.
function getContainerFolder(orderFolder) {
  const root = getLoadingRoot(orderFolder);
  return root ? getOrCreateFolder(root, 'Container Photos')
              : getOrCreateFolder(orderFolder, 'Container');
}

// ── Month consolidation on Complete ───────────────────────────
// An order is often packed in one month and loaded in the next: Steps 1 & 2 are
// month-locked to the order's Original Month (see uploadItemStep12), while Step 3
// follows today's date. That used to leave one order sitting in two month folders.
//
// On Complete the WHOLE order folder is now moved into the month the officer
// pressed Complete, so 1-Packing, 2-Cold Storage QC and 3-Loading always sit
// together. Order Master's Original Month/Year is rewritten to match, so any
// Step 1/2 item added afterwards lands in the same place instead of re-splitting.
//
// The OLDEST folder is the one kept and moved — never copied — so its Drive ID,
// and therefore every photo URL already written into the spreadsheets and the
// Folder URL in Order Master column H, stays valid.

// Look up a child folder without creating it. Returns null if absent.
function findFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

// Move everything out of src into dst. Sub-folders present on both sides are
// merged recursively rather than duplicated. src is left empty.
function mergeFolderInto(src, dst) {
  const subs = src.getFolders();
  const subList = [];
  while (subs.hasNext()) subList.push(subs.next());
  subList.forEach(function(sub) {
    const twin = findFolder(dst, sub.getName());
    if (twin) {
      mergeFolderInto(sub, twin);   // same name both sides — merge contents
      sub.setTrashed(true);         // now-empty duplicate goes to the bin
    } else {
      dst.addFolder(sub);
      src.removeFolder(sub);
    }
  });
  const files = src.getFiles();
  const fileList = [];
  while (files.hasNext()) fileList.push(files.next());
  fileList.forEach(function(f) {
    dst.addFile(f);
    src.removeFile(f);
  });
}

// Every folder carrying this order name, across all <year>/<month> folders.
function findAllOrderFolders(orderName) {
  const root  = getRootFolder();
  const found = [];
  const years = root.getFolders();
  while (years.hasNext()) {
    const y = years.next();
    if (!/^\d{4}$/.test(y.getName())) continue;   // skips _Pending and anything else
    const months = y.getFolders();
    while (months.hasNext()) {
      const f = findFolder(months.next(), orderName);
      if (f) found.push(f);
    }
  }
  return found;
}

// Bring every folder for this order together in targetYear/targetMonth and
// return the surviving folder.
function consolidateOrderFolder(orderName, targetYear, targetMonth) {
  const root         = getRootFolder();
  const targetMonthF = getOrCreateFolder(getOrCreateFolder(root, targetYear), targetMonth);

  const all = findAllOrderFolders(orderName);
  if (!all.length) return getOrCreateFolder(targetMonthF, orderName);

  // Oldest wins — that is the folder whose URL was recorded in Order Master.
  all.sort(function(a, b) { return a.getDateCreated() - b.getDateCreated(); });
  const keep = all[0];

  for (let i = 1; i < all.length; i++) {
    mergeFolderInto(all[i], keep);
    all[i].setTrashed(true);
  }

  // Move the survivor into the Complete month if it is not already there.
  const parents = [];
  const pit = keep.getParents();
  while (pit.hasNext()) parents.push(pit.next());
  const alreadyThere = parents.some(function(p) { return p.getId() === targetMonthF.getId(); });
  if (!alreadyThere) {
    targetMonthF.addFolder(keep);
    parents.forEach(function(p) { p.removeFolder(keep); });
  }
  return keep;
}

// Rewrite Order Master C=Original Month, D=Original Year after a consolidation,
// so later Step 1/2 items follow the order instead of recreating the old month.
function setOrderMasterMonth(orderName, year, month) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORDER_MASTER);
    if (!sheet || sheet.getLastRow() < 2) return;
    const key  = String(orderName).trim().toUpperCase();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === key) {
        sheet.getRange(i + 1, 3).setValue(month);
        sheet.getRange(i + 1, 4).setValue(year);
        return;
      }
    }
  } catch(e) {
    Logger.log('setOrderMasterMonth failed: ' + e);
  }
}

// Use the order's existing folder wherever it lives (oldest first, matching what
// consolidation keeps); only create a fresh month folder when the order has none.
// This keeps container photos with the rest of the order even when the save
// happens in a later month than Steps 1 & 2, and lets the legacy-layout check see
// the real folder rather than an empty shell created under today's month.
function getExistingOrCreateOrderFolder(orderName) {
  const all = findAllOrderFolders(orderName);
  if (all.length) {
    all.sort(function(a, b) { return a.getDateCreated() - b.getDateCreated(); });
    return all[0];
  }
  return getOrCreateOrderFolder(orderName);
}

// ── Step 1 & 2 (Packing + Cold Storage) ──────────────────────

function checkOrderStep12(orderName) {
  try {
    const ss = getOrCreateSpreadsheet();
    const nameKey = String(orderName).trim().toUpperCase();

    // Count items directly from log sheets — most reliable source
    function countItems(sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return 0;
      const data = sheet.getDataRange().getValues();
      let count = 0;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toUpperCase() === nameKey) count++;
      }
      return count;
    }

    const step1Count = countItems(CONFIG.SHEETS.PACKING);
    const step2Count = countItems(CONFIG.SHEETS.COLD_STORAGE);

    // Check Order Master for month-lock info
    let originalMonth = '', originalYear = '';
    const masterSheet = ss.getSheetByName(CONFIG.SHEETS.ORDER_MASTER);
    if (masterSheet && masterSheet.getLastRow() > 1) {
      const masterData = masterSheet.getDataRange().getValues();
      for (let i = 1; i < masterData.length; i++) {
        if (String(masterData[i][0]).trim().toUpperCase() === nameKey) {
          originalMonth = masterData[i][2] || '';
          originalYear  = masterData[i][3] || '';
          break;
        }
      }
    }

    const exists = step1Count > 0 || step2Count > 0 || originalMonth !== '';
    return {
      exists: exists, locked: false,
      step1Count: step1Count,
      step2Count: step2Count,
      originalMonth: originalMonth,
      originalYear: originalYear
    };
  } catch(e) {
    Logger.log('checkOrderStep12 error: ' + e);
    return { exists: false, locked: false, step1Count: 0, step2Count: 0 };
  }
}

function uploadItemStep12(body) {
  const nonceCached = checkClientNonce(body.clientNonce);
  if (nonceCached) return nonceCached;

  const lock = LockService.getScriptLock();
  const lockResult = tryWaitLock(lock, 15000);
  if (!lockResult.ok) return { success: false, message: lockResult.message };

  const savedFiles = [];
  var stage = 'init';
  try {
    const orderName      = body.orderName;
    const step           = body.step;
    const productType    = body.productType;
    const processingType = body.processingType;
    const size           = body.size;
    const staffId        = body.staffId;
    const staffName      = body.staffName || '';
    const notes          = body.notes || '';
    const clientNonce    = body.clientNonce;

    if (!orderName || !productType || !processingType || !size || !staffId) {
      return { success: false, message: '❌ ข้อมูลไม่ครบ — Order/สินค้า/การแปรรูป/ขนาด/รหัสพนักงาน' };
    }

    stage = 'check_order';
    const ss = getOrCreateSpreadsheet();
    const orderCheck = checkOrderStep12(orderName);
    let year, month;
    if (orderCheck.exists && orderCheck.originalMonth && orderCheck.originalYear) {
      year  = String(orderCheck.originalYear);
      month = String(orderCheck.originalMonth);
    } else {
      const now = new Date();
      year  = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy');
      month = Utilities.formatDate(now, CONFIG.TIMEZONE, 'MMMM');
    }

    stage = 'compute_item_number';
    const itemNumber = (step === 1 ? orderCheck.step1Count : orderCheck.step2Count) + 1;

    stage = 'create_folders';
    const root        = getRootFolder();
    const yearFolder  = getOrCreateFolder(root, year);
    const monthFolder = getOrCreateFolder(yearFolder, month);
    const orderFolder = getOrCreateFolder(monthFolder, orderName);
    const stepFolder  = getOrCreateFolder(orderFolder, step === 1 ? '1-Packing' : '2-Cold Storage QC');
    const itemFolder  = getOrCreateFolder(stepFolder, productType + '/' + processingType + ' ' + size);

    stage = 'save_photos';
    const photoLabels = ['1-รูปปลา','2-ฉลากอินเนอร์','3-รูปปลาพร้อมฉลาก','4-ถุงบรรจุ','5-กล่องพร้อมฉลาก','6-รหัสกล่อง','7-ปั้มกระสอบ','8-กระสอบ'];
    const ts = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss');
    const photoUrls = {};
    let savedCount = 0;
    for (let i = 1; i <= 8; i++) {
      const data = body['photo' + i];
      if (data && data.length > 100) {
        try {
          const file = savePhoto(itemFolder, data, photoLabels[i-1] + '_' + ts + '.jpg');
          savedFiles.push(file);
          photoUrls['photo' + i] = file.getUrl();
          savedCount++;
        } catch(photoErr) {
          throw new Error('รูปที่ ' + i + ' บันทึกไม่ได้: ' + photoErr.toString());
        }
      } else {
        photoUrls['photo' + i] = '';
      }
    }

    stage = 'update_master';
    const masterSheet = ss.getSheetByName(CONFIG.SHEETS.ORDER_MASTER) || getSheet(CONFIG.SHEETS.ORDER_MASTER);
    const masterData  = masterSheet.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < masterData.length; i++) {
      if (String(masterData[i][0]).trim().toUpperCase() === String(orderName).trim().toUpperCase()) {
        rowIdx = i + 1; break;
      }
    }
    if (rowIdx > 0) {
      masterSheet.getRange(rowIdx, step === 1 ? 5 : 6).setValue(itemNumber);
      masterSheet.getRange(rowIdx, 7).setValue(new Date());
    } else {
      masterSheet.appendRow([
        orderName, new Date(), month, year,
        step === 1 ? itemNumber : 0,
        step === 2 ? itemNumber : 0,
        new Date(), orderFolder.getUrl()
      ]);
    }

    stage = 'log_item';
    const logSheetName = step === 1 ? CONFIG.SHEETS.PACKING : CONFIG.SHEETS.COLD_STORAGE;
    const logSheet = ss.getSheetByName(logSheetName) || getSheet(logSheetName);
    logSheet.appendRow([
      orderName, itemNumber, productType, processingType, size,
      staffId, staffName, new Date(), notes, itemFolder.getUrl(),
      photoUrls.photo1, photoUrls.photo2, photoUrls.photo3, photoUrls.photo4,
      photoUrls.photo5, photoUrls.photo6, photoUrls.photo7, photoUrls.photo8
    ]);

    savedFiles.length = 0; // success — don't rollback

    const result = {
      success: true,
      message: '✅ บันทึก Item #' + itemNumber + ' สำเร็จ (' + savedCount + ' รูป)',
      itemNumber: itemNumber,
      nextItemNumber: itemNumber + 1,
      photoCount: savedCount
    };
    storeClientNonce(clientNonce, result);
    return result;

  } catch(e) {
    Logger.log('uploadItemStep12 [' + stage + '] error: ' + e);
    if (savedFiles.length > 0) cleanupSavedFiles(savedFiles);
    return { success: false, message: '❌ ขั้นตอน: ' + stage + '\n' + e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// ── Step 3 (Container Loading) ────────────────────────────────

function getRecentStep3Orders() {
  // Reads from LOADING_MASTER (Step 3 orders), not ORDER_MASTER (Step 1/2 orders).
  // Returns { name, status } objects so the frontend chip can display them correctly.
  const sheet = getSheet(CONFIG.SHEETS.LOADING_MASTER);
  if (!sheet) return { orders: [] };
  const data = sheet.getDataRange().getValues();
  const orders = [];
  // Loading Master column layout:
  // A(0)=Order Name, B(1)=First Upload Date, C(2)=Status,
  // D(3)=Complete Date, E(4)=Staff ID, F(5)=Completed By Name,
  // G(6)=Total Lines, H(7)=Final Folder URL
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    if (name) orders.push({ name: name, status: String(data[i][2] || '').trim() });
  }
  return { orders: orders.slice(-30).reverse() };
}

function checkOrderStep3(orderName) {
  const masterSheet = getSheet(CONFIG.SHEETS.LOADING_MASTER);
  if (!masterSheet) return { exists: false, status: 'not_started', lines: [] };
  const masterData = masterSheet.getDataRange().getValues();
  let found = false;
  let orderStatus = 'not_started';
  let orderInfo = {};

  for (let i = 1; i < masterData.length; i++) {
    if (String(masterData[i][0]).trim() === String(orderName).trim()) {
      found = true;
      // A(0)=Order Name, B(1)=First Upload Date, C(2)=Status,
      // D(3)=Complete Date, E(4)=Staff ID, F(5)=Completed By Name,
      // G(6)=Total Lines, H(7)=Final Folder URL
      orderStatus = String(masterData[i][2] || 'in_progress').trim();
      orderInfo = {
        staffId:         String(masterData[i][4] || '').trim(),
        staffName:       String(masterData[i][5] || '').trim(), // Completed By Name used as staff name
        startTime:       masterData[i][1] ? new Date(masterData[i][1]).toISOString() : null,
        completedByName: String(masterData[i][5] || '').trim(),
        totalLines:      Number(masterData[i][6] || 0),
        folderUrl:       String(masterData[i][7] || '').trim()
      };
      break;
    }
  }

  if (!found) return { exists: false, status: 'not_started', lines: [] };

  // Fetch individual line details from LOADING_LINES so the frontend
  // can populate line chips and the gallery on re-entry.
  const lines = [];
  const linesSheet = getSheet(CONFIG.SHEETS.LOADING_LINES);
  if (linesSheet) {
    const linesData = linesSheet.getDataRange().getValues();
    for (let i = 1; i < linesData.length; i++) {
      if (String(linesData[i][0]).trim() !== String(orderName).trim()) continue;
      const lineNumber = Number(linesData[i][1]);
      const photoUrls = [];
      // Col 5 (F) = photo count; col 6+ (G+) = actual URLs.
      // Skip col 5 so the count number is never treated as a photo URL.
      for (let c = 6; c < linesData[i].length; c++) {
        const u = String(linesData[i][c] || '').trim();
        if (u) photoUrls.push(u);
      }
      lines.push({
        lineNumber:  lineNumber,
        staffId:     String(linesData[i][2] || '').trim(),
        staffName:   String(linesData[i][3] || '').trim(),
        photosCount: photoUrls.length,
        photos:      photoUrls
      });
    }
  }

  return Object.assign({ exists: true, status: orderStatus, lines: lines }, orderInfo);
}

function uploadLineStep3(body) {
  const nonceCached = checkClientNonce(body.clientNonce);
  if (nonceCached) return nonceCached;

  const lock = LockService.getScriptLock();
  const lockResult = tryWaitLock(lock, 12000);
  if (!lockResult.ok) return { success: false, message: lockResult.message };

  const savedFiles = [];
  try {
    const { orderName, lineNumber, staffId, staffName, photos } = body;

    const masterSheet = getSheet(CONFIG.SHEETS.LOADING_MASTER);
    const masterData  = masterSheet.getDataRange().getValues();
    let orderRow = -1;
    for (let i = 1; i < masterData.length; i++) {
      if (String(masterData[i][0]).trim() === String(orderName).trim()) {
        orderRow = i + 1; break;
      }
    }
    if (orderRow === -1) {
      // A=Order Name, B=First Upload Date, C=Status, D=Complete Date,
      // E=Staff ID, F=Completed By Name, G=Total Lines, H=Final Folder URL
      masterSheet.appendRow([orderName, new Date(), 'in_progress', '', staffId, '', 0, '']);
      orderRow = masterSheet.getLastRow();
    }

    // Save photos to _Pending folder
    const root         = getRootFolder();
    const pendingRoot  = getOrCreateFolder(root, '_Pending');
    const orderPending = getOrCreateFolder(pendingRoot, orderName);
    const lineFolder   = getOrCreateFolder(orderPending, 'Line-' + lineNumber);

    const photoUrls = [];
    (photos || []).forEach((b64, idx) => {
      if (!b64) return; // skip empty slots — do not write blank cells to sheet
      const fileName = orderName + '_Line' + lineNumber + '_' + (idx + 1) + '.jpg';
      const file = savePhoto(lineFolder, b64, fileName);
      savedFiles.push(file);
      photoUrls.push(file.getUrl());
    });

    // Format: [orderName, lineNumber, staffId, staffName, date, photoCount, url1, url2, ...]
    // photoCount in col F, URLs in col G onward — matches the pre-existing sheet structure.
    const linesSheet = getSheet(CONFIG.SHEETS.LOADING_LINES);
    linesSheet.appendRow([orderName, lineNumber, staffId, staffName, new Date(), photoUrls.length, ...photoUrls]);

    const currentLines = Number(masterSheet.getRange(orderRow, 7).getValue() || 0);
    masterSheet.getRange(orderRow, 7).setValue(currentLines + 1);

    const result = { success: true };
    storeClientNonce(body.clientNonce, result);
    return result;
  } catch(e) {
    cleanupSavedFiles(savedFiles);
    Logger.log('uploadLineStep3 error: ' + e);
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function completeStep3Order(body) {
  const nonceCached = checkClientNonce(body.clientNonce);
  if (nonceCached) return nonceCached;

  const lock = LockService.getScriptLock();
  const lockResult = tryWaitLock(lock, 20000);
  if (!lockResult.ok) return { success: false, message: lockResult.message };

  try {
    const { orderName, staffId, staffName } = body;

    const masterSheet = getSheet(CONFIG.SHEETS.LOADING_MASTER);
    const masterData  = masterSheet.getDataRange().getValues();
    let orderRow = -1;
    for (let i = 1; i < masterData.length; i++) {
      if (String(masterData[i][0]).trim() === String(orderName).trim()) {
        orderRow = i + 1; break;
      }
    }
    if (orderRow === -1) return { success: false, message: 'Order not found in Loading Master' };

    // Pull the whole order into the month Complete was pressed, then move
    // _Pending/orderName → <that month>/orderName/3-Loading/Loading Lines
    const root        = getRootFolder();
    const completeYM  = getYearMonth();
    const orderFolder = consolidateOrderFolder(orderName, completeYM.year, completeYM.month);
    const step3Folder = getLinesFolder(orderFolder);

    const pendingRoot = getOrCreateFolder(root, '_Pending');
    const pendingIter = pendingRoot.getFoldersByName(orderName);
    const folderUrl   = step3Folder.getUrl();

    if (pendingIter.hasNext()) {
      const pendingOrder = pendingIter.next();
      const subIter = pendingOrder.getFolders();
      while (subIter.hasNext()) {
        const sub = subIter.next();
        step3Folder.addFolder(sub);
        pendingOrder.removeFolder(sub);
      }
      const fileIter = pendingOrder.getFiles();
      while (fileIter.hasNext()) {
        const f = fileIter.next();
        step3Folder.addFile(f);
        pendingOrder.removeFile(f);
      }
      pendingRoot.removeFolder(pendingOrder);
    }

    // C(3)=Status → completed, D(4)=Complete Date, F(6)=Completed By Name, H(8)=Folder URL
    masterSheet.getRange(orderRow, 3).setValue('completed');
    masterSheet.getRange(orderRow, 4).setValue(new Date());
    masterSheet.getRange(orderRow, 6).setValue(staffName);
    masterSheet.getRange(orderRow, 8).setValue(folderUrl);

    // Keep Order Master's month in step with where the folder now lives.
    setOrderMasterMonth(orderName, completeYM.year, completeYM.month);

    const result = { success: true, folderUrl: folderUrl };
    storeClientNonce(body.clientNonce, result);
    return result;
  } catch(e) {
    Logger.log('completeStep3Order error: ' + e);
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function revertStep3Order(body) {
  const lock = LockService.getScriptLock();
  const lockResult = tryWaitLock(lock, 10000);
  if (!lockResult.ok) return { success: false, message: lockResult.message };

  try {
    const { orderName } = body;
    // v2 fix: the page sends `password`, v1 read only `gmPassword` — so GM
    // revert always answered "Incorrect", even with the right password.
    // Accept both names.
    const gmPassword = body.gmPassword || body.password;
    if (gmPassword !== GM_PASSWORD) return { success: false, message: 'Invalid GM password' };

    const masterSheet = getSheet(CONFIG.SHEETS.LOADING_MASTER);
    const masterData  = masterSheet.getDataRange().getValues();
    for (let i = 1; i < masterData.length; i++) {
      if (String(masterData[i][0]).trim() === String(orderName).trim()) {
        // C(3)=Status, D(4)=Complete Date cleared, F(6)=Completed By cleared, H(8)=Folder URL cleared
        masterSheet.getRange(i + 1, 3).setValue('in_progress');
        masterSheet.getRange(i + 1, 4).setValue('');
        masterSheet.getRange(i + 1, 6).setValue('');
        masterSheet.getRange(i + 1, 8).setValue('');
        return { success: true };
      }
    }
    return { success: false, message: 'Order not found' };
  } finally {
    lock.releaseLock();
  }
}

// ── Container Photos ──────────────────────────────────────────

// Sheet layout — Container Photos (one row per order):
// Col A (0): Order Name
// Col B (1): Staff ID  (most recent saver)
// Col C (2): Staff Name (most recent saver)
// Col D (3): Last Save Time
//
// v2 — 11 slots. Storage order is NOT shooting order; the page holds the
// display mapping. Slots 0–5 are unchanged from v1 and never move, which is
// what makes the migration safe. Slots 6–10 are new columns K–O.
//
// Names below are the confirmed sheet headers (Container Photos, 5 Aug 2026),
// which match CONTAINER_LABELS_EN in index.html (minus ' URL'; slots 9/10 are
// shortened on the tile). The code addresses columns by position — headers
// are for humans.
//
// Col E (4):  Slot 0  URL — Open Door with fog          (on screen: 5)   redefined
// Col F (5):  Slot 1  URL — อุณหภูมิก่อนเปิดตู้           (on screen: 4)
// Col G (6):  Slot 2  URL — Temp middle with green light (on screen: 6)  redefined
// Col H (7):  Slot 3  URL — Temp rear with green light  (on screen: 8)   redefined
// Col I (8):  Slot 4  URL — เลข Seal                    (on screen: 10)
// Col J (9):  Slot 5  URL — Closed Door with seal       (on screen: 11)  redefined
// Col K (10): Slot 6  URL — Exterior side A             (on screen: 1)   NEW
// Col L (11): Slot 7  URL — Exterior side B             (on screen: 2)   NEW
// Col M (12): Slot 8  URL — Back Door before opening    (on screen: 3)   NEW
// Col N (13): Slot 9  URL — Temp middle position with Line Number (on screen: 7) NEW
// Col O (14): Slot 10 URL — Temp rear position with Line Number  (on screen: 9) NEW
// Col P (15): Photos Saved count                        (was col K in v1)

// ── Container slot geometry — single source of truth ──────────
// To change the slot count in future, change CONTAINER_SLOT_COUNT only.
// Everything below derives from it. The page must agree (CONTAINER_LABELS_*
// and CONTAINER_DISPLAY_ORDER in index.html).
const CONTAINER_SLOT_COUNT   = 11;
const CONTAINER_COL_OFFSET   = 4;                                  // col E, 0-based
const CONTAINER_COUNT_COL    = CONTAINER_COL_OFFSET + CONTAINER_SLOT_COUNT; // 15 → col P
const CONTAINER_ROW_WIDTH    = CONTAINER_COUNT_COL + 1;            // 16 columns, A–P

// v1 stored one row per slot with the slot index in col B. Those rows are still
// read for history. Col B now holds a 7-digit Staff ID, so it can never collide
// with a slot index at any plausible slot count.
function isLegacySlotRow(colBVal) {
  const n = Number(colBVal);
  return n >= 0 && n < CONTAINER_SLOT_COUNT && String(colBVal) === String(n);
}

function saveContainerPhoto(body) {
  // v2 fix: the page sends `nonce`; v1 read only `clientNonce`, so the
  // idempotency cache never hit and retries genuinely re-uploaded to Drive.
  // Accept both so this works against old and new front ends during rollout.
  const clientNonce = body.nonce || body.clientNonce;
  const nonceCached = checkClientNonce(clientNonce);
  if (nonceCached) return nonceCached;

  const lock = LockService.getScriptLock();
  const lockResult = tryWaitLock(lock, 10000);
  if (!lockResult.ok) return { success: false, message: lockResult.message };

  const savedFiles = [];
  try {
    const { orderName, slotIndex, staffId, staffName, photo, slotLabel } = body;
    if (!photo) return { success: false, message: 'No photo data' };
    const idx = Number(slotIndex);
    if (idx < 0 || idx >= CONTAINER_SLOT_COUNT) return { success: false, message: 'Invalid slot index' };

    // Save photo to Drive
    const orderFolder     = getExistingOrCreateOrderFolder(orderName);
    const containerFolder = getContainerFolder(orderFolder);
    const fileName = orderName + '_Container_Slot' + (idx + 1) + '_' + (slotLabel || '') + '.jpg';
    const file = savePhoto(containerFolder, photo, fileName);
    savedFiles.push(file);
    const fileUrl = file.getUrl();

    // Ensure order exists in Loading Master (create if first save for this order)
    const masterSheet = getSheet(CONFIG.SHEETS.LOADING_MASTER);
    if (masterSheet) {
      const masterData = masterSheet.getDataRange().getValues();
      let foundInMaster = false;
      for (let i = 1; i < masterData.length; i++) {
        if (String(masterData[i][0]).trim() === String(orderName).trim()) {
          foundInMaster = true; break;
        }
      }
      if (!foundInMaster) {
        // A=Order Name, B=First Upload Date, C=Status, D=Complete Date,
        // E=Staff ID, F=Completed By Name, G=Total Lines, H=Final Folder URL
        masterSheet.appendRow([orderName, new Date(), 'in_progress', '', staffId || '', '', 0, '']);
      }
    }

    // Upsert: one row per order
    const sheet = getSheet(CONFIG.SHEETS.CONTAINER_PHOTOS);
    const data  = sheet.getDataRange().getValues();
    const PHOTO_COL_OFFSET = CONTAINER_COL_OFFSET; // slots live at 0-based 4 … 14 (cols E–O)
    const COUNT_COL        = CONTAINER_COUNT_COL;  // col P (0-based 15)

    // Find existing single-row entry for this order
    let existingRowNum = -1; // 1-based sheet row
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== String(orderName).trim()) continue;
      // Only treat as the new one-row-per-order format when col B is NOT a slot
      // index. Old multi-row entries have a slot index in col B.
      if (!isLegacySlotRow(data[i][1])) {
        existingRowNum = i + 1; // convert 0-based array index → 1-based row
        break;
      }
    }

    if (existingRowNum === -1) {
      // No one-row entry exists yet — create one, width-derived so the row
      // always matches CONTAINER_SLOT_COUNT.
      const newRow = [orderName, staffId || '', staffName || '', new Date()];
      for (let s = 0; s < CONTAINER_SLOT_COUNT; s++) newRow.push('');
      newRow.push(0);
      newRow[PHOTO_COL_OFFSET + idx] = fileUrl;
      newRow[COUNT_COL] = 1;
      sheet.appendRow(newRow);
    } else {
      // Update the existing row in place
      const rowRange = sheet.getRange(existingRowNum, 1, 1, CONTAINER_ROW_WIDTH);
      const rowVals  = rowRange.getValues()[0];
      rowVals[1] = staffId   || rowVals[1]; // keep previous staff if not provided
      rowVals[2] = staffName || rowVals[2];
      rowVals[3] = new Date();
      rowVals[PHOTO_COL_OFFSET + idx] = fileUrl;
      // Recount non-empty URL slots
      let count = 0;
      for (let c = PHOTO_COL_OFFSET; c < PHOTO_COL_OFFSET + CONTAINER_SLOT_COUNT; c++) {
        if (String(rowVals[c] || '').trim()) count++;
      }
      rowVals[COUNT_COL] = count;
      rowRange.setValues([rowVals]);
    }

    const result = { success: true, url: fileUrl };
    storeClientNonce(clientNonce, result);
    return result;
  } catch(e) {
    cleanupSavedFiles(savedFiles);
    Logger.log('saveContainerPhoto error: ' + e);
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getContainerPhotos(orderName) {
  // Returns a flat array of CONTAINER_SLOT_COUNT URL strings indexed by
  // STORAGE slot. The page maps these to shooting order for display.
  // Handles both the one-row-per-order format AND the v1 per-slot rows
  // (where col B held the slot index and the URL was at col G / index 6).
  const urls = new Array(CONTAINER_SLOT_COUNT).fill('');
  const sheet = getSheet(CONFIG.SHEETS.CONTAINER_PHOTOS);
  if (!sheet) return { photos: urls, savedCount: 0 };
  const data = sheet.getDataRange().getValues();

  let foundOrderRow = false; // set to true once we locate a one-row entry

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== String(orderName).trim()) continue;

    const colBVal = data[i][1];

    if (isLegacySlotRow(colBVal)) {
      // v1 multi-row format: URL at index 6 (col G)
      const url = String(data[i][6] || '').trim();
      if (url) urls[Number(colBVal)] = url;
    } else {
      // One-row-per-order format: URLs at indices 4 … 4+SLOT_COUNT-1
      for (let s = 0; s < CONTAINER_SLOT_COUNT; s++) {
        const url = String(data[i][CONTAINER_COL_OFFSET + s] || '').trim();
        if (url) urls[s] = url;
      }
      foundOrderRow = true;
      break; // only one such row per order
    }
  }

  const savedCount = urls.filter(function(u){ return !!u; }).length;
  return { photos: urls, savedCount: savedCount };
}

function clearContainerSlot(body) {
  const lock = LockService.getScriptLock();
  const lockResult = tryWaitLock(lock, 10000);
  if (!lockResult.ok) return { success: false, message: lockResult.message };

  try {
    const { orderName, slotIndex, gmPassword } = body;
    if (gmPassword !== GM_PASSWORD) return { success: false, message: 'Invalid GM password' };

    const sheet = getSheet(CONFIG.SHEETS.CONTAINER_PHOTOS);
    if (!sheet) return { success: false, message: 'Container Photos sheet not found' };
    const data = sheet.getDataRange().getValues();
    const idx  = Number(slotIndex);
    if (idx < 0 || idx >= CONTAINER_SLOT_COUNT) return { success: false, message: 'Invalid slot index' };
    const PHOTO_COL_OFFSET = CONTAINER_COL_OFFSET;
    const COUNT_COL        = CONTAINER_COUNT_COL;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== String(orderName).trim()) continue;

      const colBVal = data[i][1];

      if (isLegacySlotRow(colBVal)) {
        // v1 format: delete the row for this slot
        if (Number(colBVal) === idx) {
          sheet.deleteRow(i + 1);
          return { success: true };
        }
      } else {
        // One-row format: clear the slot cell and update count
        const rowNum   = i + 1; // 1-based
        const slotCol  = PHOTO_COL_OFFSET + idx + 1; // 1-based column number
        sheet.getRange(rowNum, slotCol).clearContent();
        // Recount after clearing
        const rowVals = sheet.getRange(rowNum, 1, 1, CONTAINER_ROW_WIDTH).getValues()[0];
        let count = 0;
        for (let c = PHOTO_COL_OFFSET; c < PHOTO_COL_OFFSET + CONTAINER_SLOT_COUNT; c++) {
          if (String(rowVals[c] || '').trim()) count++;
        }
        sheet.getRange(rowNum, COUNT_COL + 1).setValue(count); // 1-based
        return { success: true };
      }
    }
    return { success: false, message: 'Slot not found' };
  } catch(e) {
    Logger.log('clearContainerSlot error: ' + e);
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ── Daily Stats ───────────────────────────────────────────────

function getDailyStats() {
  const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');

  function countToday(sheetName, dateColIndex) {
    const sheet = getSheet(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return 0;
    const lastRow = sheet.getLastRow();
    const startRow = Math.max(2, lastRow - 49); // only scan last 50 rows
    const numRows = lastRow - startRow + 1;
    const data = sheet.getRange(startRow, dateColIndex + 1, numRows, 1).getValues();
    let count = 0;
    for (let i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      try {
        if (Utilities.formatDate(new Date(data[i][0]), CONFIG.TIMEZONE, 'yyyy-MM-dd') === todayStr) count++;
      } catch(e) {}
    }
    return count;
  }

  return {
    packed:     countToday(CONFIG.SHEETS.PACKING,       7),
    qcd:        countToday(CONFIG.SHEETS.COLD_STORAGE,  7),
    containers: countToday(CONFIG.SHEETS.LOADING_LINES, 4),
  };
}

// ── GM Password Verify ────────────────────────────────────────

function verifyGMPassword(body) {
  return { valid: (body.password === GM_PASSWORD) };
}

// ── Request body validator ────────────────────────────────────

function validatePostBody(body, requiredFields) {
  for (const field of requiredFields) {
    if (body[field] === undefined || body[field] === null || body[field] === '') {
      return { valid: false, missing: field };
    }
  }
  return { valid: true };
}

// ── doPost router ─────────────────────────────────────────────

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    switch (action) {

      case 'getConfig':
        return _jsonResponse(getConfig());

      case 'validateStaff': {
        const v = validatePostBody(body, ['staffId']);
        if (!v.valid) return _jsonResponse({ success: false, error: 'Missing: ' + v.missing });
        return _jsonResponse(validateStaff(body.staffId));
      }

      case 'checkOrderStep12': {
        const v = validatePostBody(body, ['orderName']);
        if (!v.valid) return _jsonResponse({ success: false, error: 'Missing: ' + v.missing });
        return _jsonResponse(checkOrderStep12(body.orderName));
      }

      case 'uploadItemStep12': {
        const v = validatePostBody(body, ['orderName', 'step', 'staffId']);
        if (!v.valid) return _jsonResponse({ success: false, error: 'Missing: ' + v.missing });
        return _jsonResponse(uploadItemStep12(body));
      }

      case 'getRecentStep3Orders':
        return _jsonResponse(getRecentStep3Orders());

      case 'checkOrderStep3': {
        const v = validatePostBody(body, ['orderName']);
        if (!v.valid) return _jsonResponse({ success: false, error: 'Missing: ' + v.missing });
        return _jsonResponse(checkOrderStep3(body.orderName));
      }

      case 'uploadLineStep3': {
        const v = validatePostBody(body, ['orderName', 'lineNumber', 'staffId']);
        if (!v.valid) return _jsonResponse({ success: false, error: 'Missing: ' + v.missing });
        return _jsonResponse(uploadLineStep3(body));
      }

      case 'completeStep3Order': {
        const v = validatePostBody(body, ['orderName', 'staffId']);
        if (!v.valid) return _jsonResponse({ success: false, error: 'Missing: ' + v.missing });
        return _jsonResponse(completeStep3Order(body));
      }

      case 'saveContainerPhoto': {
        const v = validatePostBody(body, ['orderName', 'slotIndex', 'staffId', 'photo']);
        if (!v.valid) return _jsonResponse({ success: false, error: 'Missing: ' + v.missing });
        return _jsonResponse(saveContainerPhoto(body));
      }

      case 'getContainerPhotos': {
        const v = validatePostBody(body, ['orderName']);
        if (!v.valid) return _jsonResponse({ success: false, error: 'Missing: ' + v.missing });
        return _jsonResponse(getContainerPhotos(body.orderName));
      }

      case 'verifyGMPassword': {
        const v = validatePostBody(body, ['password']);
        if (!v.valid) return _jsonResponse({ success: false, error: 'Missing: ' + v.missing });
        return _jsonResponse(verifyGMPassword(body));
      }

      case 'revertStep3Order': {
        const v = validatePostBody(body, ['orderName', 'gmPassword']);
        if (!v.valid) return _jsonResponse({ success: false, error: 'Missing: ' + v.missing });
        return _jsonResponse(revertStep3Order(body));
      }

      case 'clearContainerSlot': {
        const v = validatePostBody(body, ['orderName', 'slotIndex', 'gmPassword']);
        if (!v.valid) return _jsonResponse({ success: false, error: 'Missing: ' + v.missing });
        return _jsonResponse(clearContainerSlot(body));
      }

      case 'getDailyStats':
        return _jsonResponse(getDailyStats());

      default:
        return _jsonResponse({ success: false, error: 'Unknown action: ' + action });
    }
  } catch(err) {
    return _jsonResponse({ success: false, error: err.toString() });
  }
}