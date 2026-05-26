/*******************************************************
 * ระบบบริหารงานพัสดุ Google Apps Script
 * ใช้ Google Sheets เป็นฐานข้อมูล + Web App
 * เวอร์ชันปรับปรุง: session 2 ชั่วโมง, โปรไฟล์, กลุ่มงาน,
 * เบิกพัสดุ, ติดตามคำขอ, อนุมัติคำขอเบิกพัสดุ
 *******************************************************/

const APP_NAME = 'ระบบบริหารงานพัสดุ';
const DEFAULT_ORG_NAME = 'ชื่อหน่วยงาน';
const TZ = Session.getScriptTimeZone() || 'Asia/Bangkok';
const SESSION_SECONDS = 7200;
const SHEETS = {
  USERS: 'Users',
  CATEGORIES: 'Categories',
  ITEMS: 'Items',
  STOCK_IN: 'StockIn',
  STOCK_OUT: 'StockOut',
  MOVEMENTS: 'StockMovements',
  REQUESTS: 'StockRequests',
  DEPARTMENTS: 'Departments',
  SETTINGS: 'Settings'
};

const HEADERS = {
  Users: ['user_id','username','password','fullname','role','department','status','profile_photo','phone','approval_note'],
  Categories: ['category_id','category_name','description','status'],
  Items: ['item_id','item_code','item_name','category_id','unit','unit_price','opening_balance','current_balance','min_stock','status','note'],
  StockIn: ['stockin_id','stockin_no','stockin_date','supplier','document_no','item_id','quantity','unit_price','total_amount','balance_before','balance_after','created_by','note'],
  StockOut: ['stockout_id','stockout_no','stockout_date','requester_name','department','purpose','item_id','quantity','unit_price','total_amount','balance_before','balance_after','approved_by','created_by','note'],
  StockMovements: ['movement_id','movement_date','item_id','movement_type','ref_no','quantity_in','quantity_out','balance_before','balance_after','unit_price','total_amount','created_by'],
  StockRequests: ['request_id','request_no','request_date','requester_user_id','requester_name','department','purpose','item_id','quantity','unit_price','total_amount','status','approved_by','approved_date','approval_note','created_by','created_at','note','stockout_no'],
  Departments: ['department_id','department_name','status'],
  Settings: ['key','value']
};

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.appName = APP_NAME;
  return t.evaluate()
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clear();
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
    sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold').setBackground('#d9ead3');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, HEADERS[name].length);
  });

  appendRows(SHEETS.USERS, [
    [uid('USR'), 'admin', 'admin123', 'ผู้ดูแลระบบ', 'Admin', 'งานพัสดุ', 'active', '', '', ''],
    [uid('USR'), 'user', 'user123', 'ผู้ใช้งานทั่วไป', 'User', 'ทั่วไป', 'active', '', '', '']
  ]);
  appendRows(SHEETS.CATEGORIES, [
    [uid('CAT'), 'วัสดุสำนักงาน', 'กระดาษ ปากกา แฟ้ม และอุปกรณ์สำนักงาน', 'active'],
    [uid('CAT'), 'วัสดุงานบ้านงานครัว', 'อุปกรณ์ทำความสะอาดและครัว', 'active'],
    [uid('CAT'), 'วัสดุคอมพิวเตอร์', 'อุปกรณ์คอมพิวเตอร์และไอที', 'active'],
    [uid('CAT'), 'วัสดุไฟฟ้า', 'หลอดไฟ สายไฟ และอุปกรณ์ไฟฟ้า', 'active'],
    [uid('CAT'), 'อื่น ๆ', 'ประเภทวัสดุอื่น ๆ', 'active']
  ]);
  appendRows(SHEETS.DEPARTMENTS, [
    [uid('DEP'), 'งานบริหารทั่วไป', 'active'],
    [uid('DEP'), 'งานพัสดุ', 'active'],
    [uid('DEP'), 'งานการเงิน', 'active'],
    [uid('DEP'), 'งานสารสนเทศ', 'active'],
    [uid('DEP'), 'อื่น ๆ', 'active']
  ]);
  appendRows(SHEETS.SETTINGS, [['app_name', APP_NAME], ['organization_name', DEFAULT_ORG_NAME], ['running_stockin', '0'], ['running_stockout', '0'], ['running_request', '0']]);
  return ok('สร้างฐานข้อมูลและข้อมูลตั้งต้นเรียบร้อย');
}

function migrateDatabase() {
  Object.keys(HEADERS).forEach(name => ensureSheetExists(name, HEADERS[name]));
  ensureSetting('app_name', APP_NAME);
  ensureSetting('organization_name', DEFAULT_ORG_NAME);
  ensureSetting('running_stockin', '0');
  ensureSetting('running_stockout', '0');
  ensureSetting('running_request', '0');
  return ok('ปรับปรุงฐานข้อมูลเรียบร้อย');
}

function login(username, password) {
  ensureSheetExists(SHEETS.USERS, HEADERS.Users);
  const users = getRows(SHEETS.USERS);
  const u = users.find(x => String(x.username).trim() === String(username).trim() && String(x.password) === String(password));
  if (!u) return fail('Username หรือ Password ไม่ถูกต้อง');
  const st = String(u.status || '').toLowerCase();
  if (st === 'pending') return fail('บัญชีของคุณรอการอนุมัติจากผู้ดูแลระบบ');
  if (st === 'rejected') return fail('บัญชีของคุณไม่ผ่านการอนุมัติ กรุณาติดต่อผู้ดูแลระบบ');
  if (st !== 'active') return fail('บัญชีถูกปิดใช้งาน');
  const token = Utilities.getUuid();
  const expiresAt = Date.now() + (SESSION_SECONDS * 1000);
  const session = Object.assign({}, u, {expiresAt});
  CacheService.getScriptCache().put('session_' + token, JSON.stringify(session), SESSION_SECONDS);
  return { success: true, token, expiresAt, user: safeUser(session) };
}

function getSession(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('session_' + token);
  if (!raw) return null;
  const user = JSON.parse(raw);
  if (user.expiresAt && Date.now() > Number(user.expiresAt)) {
    CacheService.getScriptCache().remove('session_' + token);
    return null;
  }
  return user;
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('session_' + token);
  return ok('ออกจากระบบแล้ว');
}

function requireLogin(token) {
  const user = getSession(token);
  if (!user) throw new Error('Session หมดอายุ กรุณาเข้าสู่ระบบใหม่');
  return user;
}
function requireAdmin(token) {
  const user = requireLogin(token);
  if (user.role !== 'Admin') throw new Error('สิทธิ์ไม่เพียงพอ');
  return user;
}
function safeUser(u) { const c = Object.assign({}, u); delete c.password; return c; }

function getInitialData(token) {
  migrateDatabase();
  ensureSheetExists(SHEETS.USERS, HEADERS.Users);
  ensureSheetExists(SHEETS.CATEGORIES, HEADERS.Categories);
  ensureSheetExists(SHEETS.ITEMS, HEADERS.Items);
  ensureSheetExists(SHEETS.DEPARTMENTS, HEADERS.Departments);
  ensureSheetExists(SHEETS.REQUESTS, HEADERS.StockRequests);
  const user = requireLogin(token);
  return {
    user: safeUser(findById(SHEETS.USERS, 'user_id', user.user_id) || user),
    expiresAt: user.expiresAt,
    categories: getActiveRows(SHEETS.CATEGORIES),
    items: enrichItems(getRows(SHEETS.ITEMS).filter(r => String(r.status || 'active').toLowerCase() === 'active')),
    departments: getDepartmentsSafe(),
    users: [],
    pendingCount: user.role === 'Admin' ? getPendingRequestsCount() : 0,
    notifications: getNotificationCounts(token),
    settings: getAppSettingsInternal()
  };
}



function ensureSetting(key, defaultValue) {
  ensureSheetExists(SHEETS.SETTINGS, HEADERS.Settings);
  const rows = getRows(SHEETS.SETTINGS);
  if (!rows.some(r => String(r.key) === String(key))) {
    SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS).appendRow([key, defaultValue]);
  }
}
function getSetting(key, defaultValue) {
  ensureSetting(key, defaultValue);
  const row = getRows(SHEETS.SETTINGS).find(r => String(r.key) === String(key));
  return row ? row.value : defaultValue;
}
function getAppSettingsInternal() {
  return {
    app_name: APP_NAME,
    organization_name: getSetting('organization_name', DEFAULT_ORG_NAME)
  };
}
function getAppSettings(token) {
  requireLogin(token);
  return getAppSettingsInternal();
}
function saveAppSettings(token, data) {
  requireAdmin(token);
  data = data || {};
  const org = String(data.organization_name || '').trim() || DEFAULT_ORG_NAME;
  ensureSetting('organization_name', DEFAULT_ORG_NAME);
  updateById(SHEETS.SETTINGS, 'key', 'organization_name', {value: org});
  return {success:true, message:'บันทึกตั้งค่าระบบเรียบร้อย', settings:getAppSettingsInternal()};
}

function registerUser(data) {
  migrateDatabase();
  data = data || {};
  const username = String(data.username || '').trim();
  const password = String(data.password || '');
  const fullname = String(data.fullname || '').trim();
  const phone = String(data.phone || '').trim();
  const department = String(data.department || '').trim();
  if (!username || !password || !fullname || !phone || !department) return fail('กรุณากรอกข้อมูลลงทะเบียนให้ครบถ้วน');
  const users = getRows(SHEETS.USERS);
  if (users.some(u => String(u.username).trim() === username)) return fail('Username นี้มีอยู่แล้ว');
  appendRows(SHEETS.USERS, [[uid('USR'), username, password, fullname, 'User', department, 'pending', '', phone, '']]);
  return ok('ส่งคำขอลงทะเบียนเรียบร้อย กรุณารอผู้ดูแลระบบอนุมัติ');
}

function approveUser(token, userId) {
  requireAdmin(token);
  updateById(SHEETS.USERS, 'user_id', userId, {status:'active', approval_note:''});
  return ok('อนุมัติผู้ใช้เรียบร้อย');
}

function rejectUser(token, userId, note) {
  requireAdmin(token);
  updateById(SHEETS.USERS, 'user_id', userId, {status:'rejected', approval_note: note || 'ไม่อนุมัติ'});
  return ok('ไม่อนุมัติผู้ใช้เรียบร้อย');
}

function updateMyProfile(token, data) {
  const user = requireLogin(token);
  data = data || {};
  const fullname = String(data.fullname || '').trim();
  if (!fullname) return fail('กรุณาระบุชื่อ-นามสกุล');
  updateById(SHEETS.USERS, 'user_id', user.user_id, {
    fullname,
    department: data.department || '',
    phone: data.phone || ''
  });
  const updated = findById(SHEETS.USERS, 'user_id', user.user_id);
  const session = Object.assign({}, updated, {expiresAt: user.expiresAt});
  CacheService.getScriptCache().put('session_' + token, JSON.stringify(session), Math.max(1, Math.floor((Number(user.expiresAt) - Date.now()) / 1000)));
  return {success:true, message:'บันทึกโปรไฟล์เรียบร้อย', user:safeUser(session)};
}

function getDepartmentsSafe() {
  ensureSheetExists(SHEETS.DEPARTMENTS, HEADERS.Departments);
  let deps = getActiveRows(SHEETS.DEPARTMENTS);
  if (deps.length === 0) seedDepartments();
  return getActiveRows(SHEETS.DEPARTMENTS);
}

function ensureSheetExists(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#d9ead3');
    sh.setFrozenRows(1);
    return;
  }
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].filter(String);
  headers.forEach(h => {
    if (!current.includes(h)) {
      sh.getRange(1, sh.getLastColumn()+1).setValue(h).setFontWeight('bold').setBackground('#d9ead3');
    }
  });
}

function listPublicDepartments() {
  migrateDatabase();
  return getDepartmentsSafe();
}

function seedDepartments() {
  ensureSheetExists(SHEETS.DEPARTMENTS, HEADERS.Departments);
  const existing = getRows(SHEETS.DEPARTMENTS).map(d => String(d.department_name).trim());
  const defaults = ['งานบริหารทั่วไป','งานพัสดุ','งานการเงิน','งานสารสนเทศ','อื่น ๆ'];
  const rows = defaults.filter(name => !existing.includes(name)).map(name => [uid('DEP'), name, 'active']);
  if (rows.length) appendRows(SHEETS.DEPARTMENTS, rows);
  return ok('เพิ่มข้อมูลกลุ่มงานตั้งต้นเรียบร้อย');
}

function getDashboard(token, filter) {
  const user = requireLogin(token);
  filter = filter || {type:'month'};
  const range = resolveRange(filter);
  const selectedCategory = filter.category_id || '';
  let items = enrichItems(getRows(SHEETS.ITEMS).filter(r => r.status === 'active'));
  if (selectedCategory) items = items.filter(i => String(i.category_id) === String(selectedCategory));
  const itemMap = mapBy(items, 'item_id');
  const itemIds = new Set(items.map(i => String(i.item_id)));
  const categories = getRows(SHEETS.CATEGORIES);
  const catMap = mapBy(categories, 'category_id');
  const stockIn = getRows(SHEETS.STOCK_IN).filter(r => inDateRange(r.stockin_date, range) && (!selectedCategory || itemIds.has(String(r.item_id))));
  let stockOut = getRows(SHEETS.STOCK_OUT).filter(r => inDateRange(r.stockout_date, range) && (!selectedCategory || itemIds.has(String(r.item_id))));
  const allReq = getRows(SHEETS.REQUESTS).filter(r => inDateRange(r.request_date, range));
  const myReq = allReq.filter(r => String(r.requester_user_id) === String(user.user_id));
  if (user.role !== 'Admin') stockOut = stockOut.filter(r => String(r.created_by) === String(user.fullname) || String(r.requester_name) === String(user.fullname));

  const totalItems = items.length;
  const totalStockValue = sum(items.map(i => n(i.current_balance) * n(i.unit_price)));
  const stockInQty = user.role === 'Admin' ? sum(stockIn.map(x => n(x.quantity))) : 0;
  const stockOutQty = sum(stockOut.map(x => n(x.quantity)));
  const stockOutAmount = sum(stockOut.map(x => n(x.total_amount)));
  const lowStock = items.filter(i => n(i.current_balance) <= n(i.min_stock));

  const latestIn = user.role === 'Admin' ? stockIn.sort((a,b) => dateVal(b.stockin_date)-dateVal(a.stockin_date)).slice(0,10).map(r => ({...r, item_name: itemMap[r.item_id]?.item_name || ''})) : [];
  const latestOut = stockOut.sort((a,b) => dateVal(b.stockout_date)-dateVal(a.stockout_date)).slice(0,10).map(r => ({...r, item_name: itemMap[r.item_id]?.item_name || ''}));

  const topMap = {};
  stockOut.forEach(r => {
    topMap[r.item_id] = topMap[r.item_id] || {item_id:r.item_id, item_name:itemMap[r.item_id]?.item_name || '', quantity:0, amount:0};
    topMap[r.item_id].quantity += n(r.quantity);
    topMap[r.item_id].amount += n(r.total_amount);
  });
  const topIssued = Object.values(topMap).sort((a,b)=>b.quantity-a.quantity).slice(0,10);

  const byCategoryMap = {};
  stockOut.forEach(r => {
    const item = itemMap[r.item_id] || {};
    const catId = item.category_id || '';
    const catName = catMap[catId]?.category_name || 'ไม่ระบุ';
    byCategoryMap[catName] = byCategoryMap[catName] || {category_name:catName, quantity:0, amount:0};
    byCategoryMap[catName].quantity += n(r.quantity);
    byCategoryMap[catName].amount += n(r.total_amount);
  });
  return {
    totalItems, totalStockValue, stockInQty, stockOutQty, stockOutAmount,
    lowStock, latestIn, latestOut, topIssued, byCategory: Object.values(byCategoryMap),
    range, category_id:selectedCategory,
    myPending: myReq.filter(r=>r.status==='pending').length,
    myApproved: myReq.filter(r=>r.status==='approved').length,
    myRejected: myReq.filter(r=>r.status==='rejected').length,
    pendingCount: getPendingRequestsCount(),
    role:user.role
  };
}

function listUsersInternal() { return getRows(SHEETS.USERS).map(safeUser); }
function listUsers(token) { requireAdmin(token); return listUsersInternal(); }
function saveUser(token, data) {
  requireAdmin(token);
  data = data || {};
  if (!data.username || !data.fullname) return fail('กรุณากรอก username และชื่อ-สกุล');
  const role = data.role === 'Admin' ? 'Admin' : 'User';
  const status = data.status || 'active';
  const users = getRows(SHEETS.USERS);
  const duplicate = users.find(u => String(u.username).trim() === String(data.username).trim() && String(u.user_id) !== String(data.user_id || ''));
  if (duplicate) return fail('username นี้มีอยู่แล้ว');
  if (data.user_id) {
    const payload = {username:data.username, fullname:data.fullname, role, department:data.department || '', status, phone:data.phone || '', approval_note:data.approval_note || ''};
    if (data.password) payload.password = data.password;
    updateById(SHEETS.USERS, 'user_id', data.user_id, payload);
    return ok('แก้ไขผู้ใช้เรียบร้อย');
  }
  if (!data.password) return fail('กรุณากำหนดรหัสผ่าน');
  appendRows(SHEETS.USERS, [[uid('USR'), data.username, data.password, data.fullname, role, data.department || '', status, data.profile_photo || '', data.phone || '', data.approval_note || '']]);
  return ok('เพิ่มผู้ใช้เรียบร้อย');
}
function deleteUser(token, userId) {
  const admin = requireAdmin(token);
  if (String(admin.user_id) === String(userId)) return fail('ไม่สามารถปิดใช้งานบัญชีของตนเองได้');
  updateById(SHEETS.USERS, 'user_id', userId, {status:'inactive'});
  return ok('ปิดใช้งานผู้ใช้แล้ว');
}

function listDepartments(token) { requireAdmin(token); return getRows(SHEETS.DEPARTMENTS); }
function saveDepartment(token, data) {
  requireAdmin(token);
  ensureSheetExists(SHEETS.DEPARTMENTS, HEADERS.Departments);
  data = data || {};
  const name = String(data.department_name || '').trim();
  if (!name) return fail('กรุณาระบุชื่อกลุ่มงาน');
  const status = data.status || 'active';
  const deps = getRows(SHEETS.DEPARTMENTS);
  const duplicate = deps.find(d => String(d.department_name).trim() === name && String(d.department_id) !== String(data.department_id || ''));
  if (duplicate) return fail('ชื่อกลุ่มงานนี้มีอยู่แล้ว');
  if (data.department_id) {
    updateById(SHEETS.DEPARTMENTS, 'department_id', data.department_id, {department_name:name, status});
    return ok('แก้ไขกลุ่มงานเรียบร้อย');
  }
  appendRows(SHEETS.DEPARTMENTS, [[uid('DEP'), name, status]]);
  return ok('เพิ่มกลุ่มงานเรียบร้อย');
}
function deleteDepartment(token, departmentId) {
  requireAdmin(token);
  updateById(SHEETS.DEPARTMENTS, 'department_id', departmentId, {status:'inactive'});
  return ok('ปิดใช้งานกลุ่มงานแล้ว');
}

function saveCategory(token, data) {
  requireAdmin(token);
  data = data || {};
  if (!data.category_name) return fail('กรุณาระบุชื่อประเภทวัสดุ');
  if (data.category_id) {
    updateById(SHEETS.CATEGORIES, 'category_id', data.category_id, {category_name: data.category_name, description: data.description || '', status: data.status || 'active'});
    return ok('แก้ไขประเภทวัสดุเรียบร้อย');
  }
  appendRows(SHEETS.CATEGORIES, [[uid('CAT'), data.category_name, data.description || '', data.status || 'active']]);
  return ok('เพิ่มประเภทวัสดุเรียบร้อย');
}
function deleteCategory(token, categoryId) {
  requireAdmin(token);
  categoryId = String(categoryId || '').trim();
  if (!categoryId) return fail('ไม่พบประเภทวัสดุที่ต้องการลบ');

  // ห้ามลบประเภทวัสดุที่ยังมีรายการพัสดุใช้งานอยู่
  const activeItems = getRows(SHEETS.ITEMS).filter(i =>
    String(i.category_id) === categoryId &&
    String(i.status || 'active').toLowerCase() === 'active'
  );
  if (activeItems.length > 0) {
    return fail('ไม่สามารถลบได้ เนื่องจากยังมีรายการพัสดุที่ใช้ประเภทนี้อยู่');
  }

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.CATEGORIES);
  if (!sh || sh.getLastRow() < 2) return fail('ไม่พบประเภทวัสดุที่ต้องการลบ');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idIndex = headers.indexOf('category_id');
  if (idIndex < 0) return fail('ไม่พบหัวตาราง category_id');

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][idIndex]) === categoryId) {
      sh.deleteRow(i + 2);
      return ok('ลบประเภทวัสดุเรียบร้อย');
    }
  }
  return fail('ไม่พบประเภทวัสดุที่ต้องการลบ');
}
function listCategories(token) { requireLogin(token); return getRows(SHEETS.CATEGORIES); }

function saveItem(token, data) {
  requireAdmin(token);
  data = data || {};
  if (!data.item_name) return fail('กรุณาระบุชื่อวัสดุ');
  const currentBalance = Math.max(0, Math.floor(n(data.current_balance)));
  const existing = data.item_id ? findById(SHEETS.ITEMS, 'item_id', data.item_id) : null;
  const payload = {
    item_code: data.item_code || autoCode('ITM'), item_name: data.item_name, category_id: data.category_id || '', unit: data.unit || '',
    unit_price: Math.max(0, Math.floor(n(data.unit_price))), opening_balance: existing ? Math.max(0, Math.floor(n(existing.opening_balance))) : currentBalance,
    current_balance: currentBalance, min_stock: Math.max(0, Math.floor(n(data.min_stock))), status: data.status || 'active', note: data.note || ''
  };
  if (data.item_id) { updateById(SHEETS.ITEMS, 'item_id', data.item_id, payload); return ok('แก้ไขรายการวัสดุเรียบร้อย'); }
  appendRows(SHEETS.ITEMS, [[uid('ITM'), payload.item_code, payload.item_name, payload.category_id, payload.unit, payload.unit_price, payload.opening_balance, payload.current_balance, payload.min_stock, payload.status, payload.note]]);
  return ok('เพิ่มรายการวัสดุเรียบร้อย');
}
function deleteItem(token, itemId) { requireAdmin(token); updateById(SHEETS.ITEMS, 'item_id', itemId, {status:'inactive'}); return ok('ปิดใช้งานรายการวัสดุแล้ว'); }
function listItems(token) { requireLogin(token); return enrichItems(getRows(SHEETS.ITEMS).filter(r => String(r.status || 'active').toLowerCase() === 'active')); }

function qtyInt(value) { const x = Number(value); if (!Number.isInteger(x) || x <= 0) throw new Error('จำนวนต้องเป็นเลขจำนวนเต็มมากกว่า 0 เท่านั้น'); return x; }

function createStockIn(token, data) {
  const user = requireAdmin(token);
  data = data || {};
  if (!data.item_id || n(data.quantity) <= 0) return fail('กรุณาเลือกรายการวัสดุและระบุจำนวนรับเข้า');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const item = findById(SHEETS.ITEMS, 'item_id', data.item_id);
    if (!item) return fail('ไม่พบรายการวัสดุ');
    const qty = qtyInt(data.quantity);
    const price = Math.max(0, Math.floor(n(data.unit_price || item.unit_price)));
    const before = n(item.current_balance);
    const after = before + qty;
    const total = qty * price;
    const no = nextNo('IN');
    updateById(SHEETS.ITEMS, 'item_id', data.item_id, {current_balance: after, unit_price: price});
    appendRows(SHEETS.STOCK_IN, [[uid('SIN'), no, toDateText(data.stockin_date || new Date()), data.supplier || '', data.document_no || '', data.item_id, qty, price, total, before, after, user.fullname, data.note || '']]);
    appendRows(SHEETS.MOVEMENTS, [[uid('MOV'), toDateText(data.stockin_date || new Date()), data.item_id, 'IN', no, qty, 0, before, after, price, total, user.fullname]]);
    return ok('บันทึกรับเข้าวัสดุเรียบร้อย');
  } finally { lock.releaseLock(); }
}

function createStockOut(token, data) {
  const user = requireLogin(token);
  data = data || {};
  if (!data.item_id || n(data.quantity) <= 0) return fail('กรุณาเลือกรายการวัสดุและระบุจำนวนเบิก');
  const item = findById(SHEETS.ITEMS, 'item_id', data.item_id);
  if (!item) return fail('ไม่พบรายการวัสดุ');
  const qty = qtyInt(data.quantity);
  if (qty > n(item.current_balance)) return fail('จำนวนเบิกเกินจำนวนคงเหลือปัจจุบัน');
  const price = Math.max(0, Math.floor(n(item.unit_price)));
  const total = qty * price;
  const no = nextNo('REQ');
  appendRows(SHEETS.REQUESTS, [[uid('REQ'), no, toDateText(data.stockout_date || new Date()), user.user_id, user.fullname, user.department || '', data.purpose || '', data.item_id, qty, price, total, 'pending', '', '', '', user.fullname, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), data.note || '', '']]);
  return ok('ส่งคำขอเบิกพัสดุเรียบร้อย กรุณารอการอนุมัติ');
}

function listMyRequests(token) {
  const user = requireLogin(token);
  const itemMap = mapBy(enrichItems(getRows(SHEETS.ITEMS)), 'item_id');
  return getRows(SHEETS.REQUESTS)
    .filter(r => user.role === 'Admin' || String(r.requester_user_id) === String(user.user_id))
    .sort((a,b)=>dateVal(b.created_at || b.request_date)-dateVal(a.created_at || a.request_date))
    .map(r => ({...r, item_name:itemMap[r.item_id]?.item_name||'', category_name:itemMap[r.item_id]?.category_name||'', unit:itemMap[r.item_id]?.unit||''}));
}

function listPendingRequests(token) {
  requireAdmin(token);
  return listMyRequests(token).filter(r => r.status === 'pending');
}
function getPendingRequestsCount() { return getRows(SHEETS.REQUESTS).filter(r => r.status === 'pending').length; }

function getNotificationCounts(token) {
  const user = requireLogin(token);
  const requests = getRows(SHEETS.REQUESTS);
  const counts = {
    pendingApprovals: 0,
    pendingUsers: 0,
    myUpdatedRequests: 0
  };
  if (user.role === 'Admin') {
    counts.pendingApprovals = requests.filter(r => String(r.status) === 'pending').length;
    counts.pendingUsers = getRows(SHEETS.USERS).filter(u => String(u.status) === 'pending').length;
  } else {
    counts.myUpdatedRequests = requests.filter(r =>
      String(r.requester_user_id) === String(user.user_id) &&
      ['approved','rejected'].includes(String(r.status))
    ).length;
  }
  return counts;
}

function approveRequest(token, requestId, note) {
  const admin = requireAdmin(token);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const req = findById(SHEETS.REQUESTS, 'request_id', requestId);
    if (!req) return fail('ไม่พบคำขอเบิกพัสดุ');
    if (req.status !== 'pending') return fail('คำขอนี้ถูกดำเนินการแล้ว');
    const item = findById(SHEETS.ITEMS, 'item_id', req.item_id);
    if (!item) return fail('ไม่พบรายการวัสดุ');
    const qty = qtyInt(req.quantity);
    const before = n(item.current_balance);
    if (qty > before) return fail('จำนวนคงเหลือไม่เพียงพอ ไม่สามารถอนุมัติได้');
    const price = Math.max(0, Math.floor(n(req.unit_price || item.unit_price)));
    const after = before - qty;
    const total = qty * price;
    const outNo = nextNo('OUT');
    updateById(SHEETS.ITEMS, 'item_id', req.item_id, {current_balance: after});
    appendRows(SHEETS.STOCK_OUT, [[uid('SOUT'), outNo, toDateText(req.request_date || new Date()), req.requester_name || '', req.department || '', req.purpose || '', req.item_id, qty, price, total, before, after, admin.fullname, req.created_by || req.requester_name || '', req.note || '']]);
    appendRows(SHEETS.MOVEMENTS, [[uid('MOV'), toDateText(req.request_date || new Date()), req.item_id, 'OUT', outNo, 0, qty, before, after, price, total, admin.fullname]]);
    updateById(SHEETS.REQUESTS, 'request_id', requestId, {status:'approved', approved_by:admin.fullname, approved_date:Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), approval_note:note || '', stockout_no:outNo});
    return ok('อนุมัติและตัดสต็อกเรียบร้อย');
  } finally { lock.releaseLock(); }
}

function rejectRequest(token, requestId, note) {
  const admin = requireAdmin(token);
  const req = findById(SHEETS.REQUESTS, 'request_id', requestId);
  if (!req) return fail('ไม่พบคำขอเบิกพัสดุ');
  if (req.status !== 'pending') return fail('คำขอนี้ถูกดำเนินการแล้ว');
  updateById(SHEETS.REQUESTS, 'request_id', requestId, {status:'rejected', approved_by:admin.fullname, approved_date:Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), approval_note:note || 'ไม่อนุมัติ'});
  return ok('บันทึกไม่อนุมัติเรียบร้อย');
}

function getReport(token, type, filter) {
  const user = requireLogin(token);
  if (user.role !== 'Admin') type = 'balance';
  filter = filter || {type:'month'};
  const range = resolveRange(filter);
  const selectedCategory = String(filter.category_id || '');
  const items = enrichItems(getRows(SHEETS.ITEMS));
  const itemMap = mapBy(items, 'item_id');
  const inSelectedCategory = itemId => !selectedCategory || String(itemMap[itemId]?.category_id || '') === selectedCategory;
  const addSeq = rows => rows.map((r,i) => Object.assign({seq:i+1}, r));
  let rows = [];
  if (type === 'balance' || type === 'value') {
    rows = items
      .filter(i => i.status === 'active' && (!selectedCategory || String(i.category_id) === selectedCategory))
      .map(i => ({item_code:i.item_code, item_name:i.item_name, category_name:i.category_name, current_balance:n(i.current_balance), unit:i.unit, unit_price:n(i.unit_price), total_amount:n(i.current_balance)*n(i.unit_price), min_stock:n(i.min_stock)}));
  } else if (type === 'low') {
    rows = items
      .filter(i => i.status === 'active' && (!selectedCategory || String(i.category_id) === selectedCategory) && n(i.current_balance) <= n(i.min_stock))
      .map(i => ({item_code:i.item_code,item_name:i.item_name,category_name:i.category_name,current_balance:n(i.current_balance),unit:i.unit,min_stock:n(i.min_stock)}));
  } else if (type === 'requests') {
    rows = getRows(SHEETS.REQUESTS)
      .filter(r=>inDateRange(r.request_date, range) && inSelectedCategory(r.item_id) && (user.role==='Admin' || String(r.requester_user_id)===String(user.user_id)))
      .map(r => ({request_date:r.request_date, request_no:r.request_no, requester_name:r.requester_name, department:r.department, item_code:itemMap[r.item_id]?.item_code||'', item_name:itemMap[r.item_id]?.item_name||'', category_name:itemMap[r.item_id]?.category_name||'', quantity:n(r.quantity), unit:itemMap[r.item_id]?.unit||'', unit_price:n(r.unit_price), total_amount:n(r.total_amount), status:r.status, approved_by:r.approved_by||'', approval_note:r.approval_note||''}));
  } else if (type === 'in') {
    if (user.role !== 'Admin') return {rows:[], range};
    rows = getRows(SHEETS.STOCK_IN)
      .filter(r=>inDateRange(r.stockin_date, range) && inSelectedCategory(r.item_id))
      .map(r => ({stockin_date:r.stockin_date, stockin_no:r.stockin_no, item_code:itemMap[r.item_id]?.item_code||'', item_name:itemMap[r.item_id]?.item_name||'', category_name:itemMap[r.item_id]?.category_name||'', quantity:n(r.quantity), unit:itemMap[r.item_id]?.unit||'', unit_price:n(r.unit_price), total_amount:n(r.total_amount), created_by:r.created_by, note:r.note||''}));
  } else if (type === 'out') {
    rows = getRows(SHEETS.STOCK_OUT)
      .filter(r=>inDateRange(r.stockout_date, range) && inSelectedCategory(r.item_id) && (user.role==='Admin' || String(r.requester_name)===String(user.fullname)))
      .map(r => ({stockout_date:r.stockout_date, stockout_no:r.stockout_no, requester_name:r.requester_name, department:r.department, item_code:itemMap[r.item_id]?.item_code||'', item_name:itemMap[r.item_id]?.item_name||'', category_name:itemMap[r.item_id]?.category_name||'', quantity:n(r.quantity), unit:itemMap[r.item_id]?.unit||'', unit_price:n(r.unit_price), total_amount:n(r.total_amount), status:'approved', approved_by:r.approved_by||'', note:r.note||''}));
  } else if (type === 'out_category') {
    const m = {};
    getRows(SHEETS.STOCK_OUT).filter(r=>inDateRange(r.stockout_date, range) && inSelectedCategory(r.item_id) && user.role==='Admin').forEach(r=>{
      const item = itemMap[r.item_id] || {}; const name = item.category_name || 'ไม่ระบุ';
      m[name] = m[name] || {category_name:name, quantity:0, total_amount:0};
      m[name].quantity += n(r.quantity); m[name].total_amount += n(r.total_amount);
    }); rows = Object.values(m);
  } else if (type === 'out_department') {
    const m = {};
    getRows(SHEETS.STOCK_OUT).filter(r=>inDateRange(r.stockout_date, range) && inSelectedCategory(r.item_id) && user.role==='Admin').forEach(r=>{
      const name = r.department || 'ไม่ระบุ';
      m[name] = m[name] || {department:name, quantity:0, total_amount:0};
      m[name].quantity += n(r.quantity); m[name].total_amount += n(r.total_amount);
    }); rows = Object.values(m);
  }
  return {rows:addSeq(rows), range};
}
function exportReportData(token, type, filter) { return getReport(token, type, filter).rows; }

function getRows(sheetName) {
  ensureSheetExists(sheetName, HEADERS[sheetName] || []);
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  return sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().map(row => {
    const obj = {};
    headers.forEach((h,i)=> obj[h] = row[i] instanceof Date ? toDateText(row[i]) : row[i]);
    return obj;
  });
}
function getActiveRows(sheetName) { return getRows(sheetName).filter(r => String(r.status || 'active').toLowerCase() === 'active'); }
function appendRows(sheetName, rows) { const sh=SpreadsheetApp.getActive().getSheetByName(sheetName); sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows); }
function findById(sheetName, idField, id) { return getRows(sheetName).find(r => String(r[idField]) === String(id)); }
function updateById(sheetName, idField, id, data) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const values = sh.getRange(2,1,Math.max(sh.getLastRow()-1,0),sh.getLastColumn()).getValues();
  const idIndex = headers.indexOf(idField);
  for (let i=0; i<values.length; i++) {
    if (String(values[i][idIndex]) === String(id)) {
      Object.keys(data).forEach(k => { const col = headers.indexOf(k); if (col >= 0) sh.getRange(i+2, col+1).setValue(data[k]); });
      return true;
    }
  }
  throw new Error('ไม่พบข้อมูลที่ต้องการแก้ไข');
}
function mapBy(rows, key) { const m={}; rows.forEach(r=>m[r[key]]=r); return m; }
function enrichItems(items) { const cats = mapBy(getRows(SHEETS.CATEGORIES), 'category_id'); return items.map(i => ({...i, category_name: cats[i.category_id]?.category_name || ''})); }
function uid(prefix) { return prefix + '-' + Utilities.getUuid().slice(0,8).toUpperCase(); }
function autoCode(prefix) { return prefix + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss'); }
function n(v) { const x = Number(v); return isNaN(x) ? 0 : x; }
function sum(arr) { return arr.reduce((a,b)=>a+n(b),0); }
function ok(message) { return {success:true, message}; }
function fail(message) { return {success:false, message}; }
function toDateText(d) { return Utilities.formatDate(new Date(d), TZ, 'yyyy-MM-dd'); }
function dateVal(d) { return new Date(d).getTime() || 0; }
function inDateRange(d, range) { const x = new Date(d); return x >= new Date(range.start + 'T00:00:00') && x <= new Date(range.end + 'T23:59:59'); }
function resolveRange(filter) {
  const now = new Date(); const today = toDateText(now); let start = today, end = today; const type = filter.type || 'month';
  if (type === 'today') {}
  else if (type === 'week') { const d = new Date(now); const day = d.getDay() || 7; d.setDate(d.getDate() - day + 1); start = toDateText(d); end = today; }
  else if (type === 'month') { start = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), 1), TZ, 'yyyy-MM-dd'); }
  else if (type === 'year') { start = Utilities.formatDate(new Date(now.getFullYear(), 0, 1), TZ, 'yyyy-MM-dd'); }
  else if (type === 'custom') { start = filter.start || today; end = filter.end || today; }
  return {start, end, type};
}
function nextNo(kind) {
  const prefix = kind === 'IN' ? 'RI' : kind === 'OUT' ? 'BO' : 'REQ';
  const key = kind === 'IN' ? 'running_stockin' : kind === 'OUT' ? 'running_stockout' : 'running_request';
  const rows = getRows(SHEETS.SETTINGS); let row = rows.find(r => r.key === key); let next = row ? n(row.value)+1 : 1;
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);
  if (row) updateById(SHEETS.SETTINGS, 'key', key, {value: next}); else sh.appendRow([key, next]);
  return prefix + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd') + '-' + ('0000'+next).slice(-4);
}
