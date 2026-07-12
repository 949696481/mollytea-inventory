/**
 * 库存登记 —— Google Apps Script 后端(多门店共用一个 Sheet + 一次部署)
 *
 * 跟旧版不一样:以前是每个门店一个独立 Google Sheet + 独立部署;现在**整个
 * 生意共用一张表格、一次部署**——门店/分类/物品/记录都在这一张表格里,员工
 * 和管理员登录同一个网址,靠账号角色区分权限,而不是靠"知道哪个门店的链接"。
 *
 * 部署步骤(只需要做一次):
 * 1. 新建一个 Google Sheet(专门给这个 App 用,不要跟别的表格混用)
 * 2. 菜单栏 扩展程序 (Extensions) -> Apps Script
 * 3. 把这个文件的全部内容粘贴进去,覆盖默认代码
 * 4. 右上角"部署" (Deploy) -> "新建部署" (New deployment)
 *    - 类型选 "网页应用" (Web app)
 *    - 执行身份 (Execute as): 我 (Me)
 *    - 谁可以访问 (Who has access): 任何人 (Anyone)
 * 5. 部署后会得到一个网址,形如 https://script.google.com/macros/s/xxxx/exec
 *    把这个网址填进 App 前端的 config.js 里的 API_BASE_URL
 * 6. (建议,一次性)在 Apps Script 编辑器里选中 installDailyCleanupTrigger
 *    函数,点"运行" —— 装一个每天清理过期登录会话的定时任务,免得
 *    sessions 表越滚越大。只需要点这一次,以后每天自动跑。
 * 7. 第一次打开 App 网址会看到"设置管理员账号"——那就是你自己的登录账号;
 *    之后在 App 里"账号管理"页面给员工建账号,员工绑定固定门店,只能录入
 *    当天数据和翻历史,改字段/物品/结算货币这些配置类操作都需要管理员账号。
 *
 * 安全说明:这份部署对整个互联网可访问(网址知道就能连上),所以跟桌面版
 * (数据全在本机、无网络访问)不同,这里每一次会修改数据的请求都要带着登录
 * 时发的 token,后端会重新核实身份和角色——不能只靠前端把按钮藏起来。
 * 密码用加盐迭代 HMAC-SHA256(见下方 PBKDF2_ITERATIONS 注释),强度比桌面版
 * 本地存储用的 200,000 轮 PBKDF2 弱一些——Apps Script 没有原生 PBKDF2,轮数
 * 太高会让登录明显变慢,这里的强度对"几个店员"这种规模是合理的取舍。
 */

// ---------- 表结构 ----------
const STORES_SHEET = "stores";
const CATEGORIES_SHEET = "categories";
const USERS_SHEET = "users";
const SESSIONS_SHEET = "sessions";

const STORES_HEADERS = ["id", "name", "created_at"];
const CATEGORIES_HEADERS = ["id", "store_id", "name", "created_at", "settlement_currency"];
const USERS_HEADERS = ["id", "username", "password_hash", "salt", "role", "display_name", "store_id"];
const SESSIONS_HEADERS = ["token", "user_id", "created_at", "expires_at"];
const FIELDS_HEADERS = ["id", "name", "type", "role", "order"];
const ITEMS_HEADERS = ["id", "name", "unit", "price", "currency", "locked"];
const LOG_HEADERS = ["date", "item_id", "item_name", "values_json", "usage", "cost", "edited_at", "edited_by"];
const HISTORY_HEADERS = ["date", "item_id", "item_name", "values_json", "usage", "cost", "edited_at", "edited_by"];

const DEFAULT_CURRENCY = "CNY";
const SESSION_TTL_DAYS = 14;
const CACHE_TTL_SECONDS = 21600; // CacheService 上限约 6 小时,更长的有效期靠 sessions 表兜底
const PBKDF2_ITERATIONS = 8000;  // 手搓的迭代 HMAC-SHA256,不是真 PBKDF2——见文件头说明

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function fieldsSheetName(categoryId) { return "fields_" + categoryId; }
function itemsSheetName(categoryId) { return "items_" + categoryId; }
function logSheetName(categoryId) { return "log_" + categoryId; }
function historySheetName(categoryId) { return "history_" + categoryId; }

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function formatDate(d) {
  // 日期字符串(如客户端传来的 "2026-07-12")按 ECMA 规范会被当成 UTC 零点解析,
  // 但 Sheet 里已存的日期格是按表格时区自动转换的本地零点——两者一比就可能差一天,
  // 导致同一天的 upsert 找不到已有行、变成重复插入。字符串直接按字面 Y-M-D 截取,
  // 完全绕开时区解析,避免这个不一致。
  if (typeof d === "string") {
    const m = d.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return m[1] + "-" + m[2] + "-" + m[3];
  }
  const date = new Date(d);
  return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0") + "-" + String(date.getDate()).padStart(2, "0");
}

function isSameDate(a, b) { return formatDate(a) === formatDate(b); }

function parseValues(raw) {
  if (raw === undefined || raw === null || raw === "") return {};
  try { return JSON.parse(raw); } catch (e) { return {}; }
}

function slugify(name) {
  const slug = name.trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  return slug || "x";
}

// ---------- 密码哈希(加盐迭代 HMAC-SHA256) ----------
function bytesToHex(bytes) {
  return bytes.map(function (b) {
    const v = b < 0 ? b + 256 : b;
    const h = v.toString(16);
    return h.length === 1 ? "0" + h : h;
  }).join("");
}

function genSalt() {
  return Utilities.getUuid().replace(/-/g, "");
}

function hashPassword(password, saltHex) {
  let hex = saltHex;
  for (let i = 0; i < PBKDF2_ITERATIONS; i++) {
    const bytes = Utilities.computeHmacSha256Signature(hex + password, saltHex);
    hex = bytesToHex(bytes);
  }
  return hex;
}

// ---------- 登录会话 ----------
function createSession(userId) {
  const token = Utilities.getUuid();
  const now = new Date();
  const expires = new Date(now.getTime() + SESSION_TTL_DAYS * 24 * 60 * 60 * 1000);
  const sheet = getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
  sheet.appendRow([token, userId, now.toISOString(), expires.toISOString()]);
  cacheSession(token, userId, expires);
  return token;
}

function cacheSession(token, userId, expiresDate) {
  const ttl = Math.min(CACHE_TTL_SECONDS, Math.max(1, Math.floor((expiresDate.getTime() - Date.now()) / 1000)));
  CacheService.getScriptCache().put("sess_" + token, userId + "|" + expiresDate.toISOString(), ttl);
}

function resolveSession(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const cached = cache.get("sess_" + token);
  if (cached) {
    const parts = cached.split("|");
    if (new Date(parts[1]).getTime() > Date.now()) return getUserById(parts[0]);
    return null;
  }
  const sheet = getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === token) {
      const expiresIso = rows[i][3];
      if (new Date(expiresIso).getTime() > Date.now()) {
        cacheSession(token, rows[i][1], new Date(expiresIso));
        return getUserById(rows[i][1]);
      }
      return null;
    }
  }
  return null;
}

function deleteSession(token) {
  CacheService.getScriptCache().remove("sess_" + token);
  const sheet = getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === token) { sheet.deleteRow(i + 1); break; }
  }
}

function pruneExpiredSessions() {
  const sheet = getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const now = Date.now();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (new Date(rows[i][3]).getTime() <= now) sheet.deleteRow(i + 1);
  }
}

function installDailyCleanupTrigger() {
  ScriptApp.newTrigger("pruneExpiredSessions").timeBased().everyDays(1).create();
}

// ---------- 账号 ----------
function rowToUser(row) {
  return { id: row[0], username: row[1], password_hash: row[2], salt: row[3], role: row[4], display_name: row[5], store_id: row[6] };
}

function publicUser(u) {
  return { id: u.id, username: u.username, displayName: u.display_name, role: u.role, storeId: u.store_id || null };
}

function getUserById(userId) {
  const sheet = getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(userId)) return rowToUser(rows[i]);
  }
  return null;
}

function findUserByUsername(username) {
  const sheet = getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(username)) return rowToUser(rows[i]);
  }
  return null;
}

function hasAdmin() {
  const sheet = getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][4] === "admin") return true;
  }
  return false;
}

function listUsers() {
  const sheet = getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push(publicUser(rowToUser(rows[i])));
  }
  return out;
}

function registerAdmin(username, password, displayName) {
  username = (username || "").trim();
  if (!username || !password) throw new Error("用户名和密码不能为空。");
  if (hasAdmin()) throw new Error("管理员账号已经存在了。");
  if (findUserByUsername(username)) throw new Error("用户名「" + username + "」已经被使用。");
  const salt = genSalt();
  const id = "user_" + new Date().getTime();
  const name = (displayName || "").trim() || username;
  const sheet = getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
  sheet.appendRow([id, username, hashPassword(password, salt), salt, "admin", name, ""]);
  const token = createSession(id);
  return { user: { id: id, username: username, displayName: name, role: "admin", storeId: null }, token: token };
}

function createEmployee(username, password, displayName, storeId) {
  username = (username || "").trim();
  if (!username || !password) throw new Error("用户名和密码不能为空。");
  if (!storeId) throw new Error("请给这个员工账号选一个门店。");
  if (findUserByUsername(username)) throw new Error("用户名「" + username + "」已经被使用。");
  const salt = genSalt();
  const id = "user_" + new Date().getTime();
  const name = (displayName || "").trim() || username;
  const sheet = getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
  sheet.appendRow([id, username, hashPassword(password, salt), salt, "employee", name, storeId]);
  return { id: id, username: username, displayName: name, role: "employee", storeId: storeId };
}

function deleteUserById(userId) {
  const sheet = getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(userId)) {
      if (rows[i][4] === "admin") throw new Error("不能删除管理员账号。");
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

function resetPassword(userId, newPassword) {
  if (!newPassword) throw new Error("新密码不能为空。");
  const sheet = getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(userId)) {
      const salt = genSalt();
      sheet.getRange(i + 1, 3, 1, 2).setValues([[hashPassword(newPassword, salt), salt]]);
      return;
    }
  }
  throw new Error("找不到这个账号。");
}

function login(username, password) {
  const user = findUserByUsername(username);
  if (!user || hashPassword(password, user.salt) !== user.password_hash) {
    throw new Error("用户名或密码不对。");
  }
  const token = createSession(user.id);
  const result = publicUser(user);
  if (user.role === "employee" && user.store_id) {
    const store = getStoreById(user.store_id);
    result.storeName = store ? store.name : null;
  }
  return { user: result, token: token };
}

// ---------- 权限校验 ----------
function requireAdmin(session) {
  if (!session || session.role !== "admin") throw new Error("只有管理员账号能做这个操作。");
}

function requireStoreForCategory(session, categoryId) {
  if (session.role === "employee") {
    const storeId = getCategoryStoreId(categoryId);
    if (String(storeId) !== String(session.store_id)) throw new Error("没有权限操作这个门店的数据。");
  }
}

// ---------- 门店 ----------
function listStores() {
  const sheet = getOrCreateSheet(STORES_SHEET, STORES_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push({ id: rows[i][0], name: rows[i][1] });
  }
  return out;
}

function getStoreById(storeId) {
  const stores = listStores();
  for (let i = 0; i < stores.length; i++) {
    if (String(stores[i].id) === String(storeId)) return stores[i];
  }
  return null;
}

function createStore(name) {
  name = (name || "").trim();
  if (!name) throw new Error("门店名称不能为空。");
  const stores = listStores();
  for (let i = 0; i < stores.length; i++) {
    if (String(stores[i].name) === String(name)) throw new Error("已经有一个叫「" + name + "」的门店了。");
  }
  const baseSlug = slugify(name);
  const existingIds = stores.map(function (s) { return s.id; });
  let slug = baseSlug, n = 2;
  while (existingIds.indexOf(slug) >= 0) { slug = baseSlug + "_" + n; n++; }
  const sheet = getOrCreateSheet(STORES_SHEET, STORES_HEADERS);
  sheet.appendRow([slug, name, new Date().toISOString()]);
  return { id: slug, name: name };
}

function deleteStoreById(storeId) {
  const sheet = getOrCreateSheet(STORES_SHEET, STORES_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(storeId)) { sheet.deleteRow(i + 1); break; }
  }
  const categoriesSheet = getOrCreateSheet(CATEGORIES_SHEET, CATEGORIES_HEADERS);
  const catRows = categoriesSheet.getDataRange().getValues();
  const toDelete = [];
  for (let i = 1; i < catRows.length; i++) {
    if (String(catRows[i][1]) === String(storeId)) toDelete.push(catRows[i][0]);
  }
  toDelete.forEach(function (catId) { deleteCategoryById(catId); });
}

// ---------- 分类 ----------
function listCategories(storeId) {
  const sheet = getOrCreateSheet(CATEGORIES_SHEET, CATEGORIES_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const id = rows[i][0], sId = rows[i][1], name = rows[i][2], settlementCurrency = rows[i][4];
    if (!id || String(sId) !== String(storeId)) continue;
    out.push({ id: id, name: name, settlementCurrency: settlementCurrency || DEFAULT_CURRENCY });
  }
  return out;
}

function getCategoryRow(categoryId) {
  const sheet = getOrCreateSheet(CATEGORIES_SHEET, CATEGORIES_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(categoryId)) return { rowIndex: i + 1, values: rows[i] };
  }
  return null;
}

function getCategoryStoreId(categoryId) {
  const row = getCategoryRow(categoryId);
  return row ? row.values[1] : null;
}

function createCategory(storeId, name) {
  name = (name || "").trim();
  if (!name) throw new Error("分类名称不能为空。");
  const id = "cat_" + new Date().getTime();
  const sheet = getOrCreateSheet(CATEGORIES_SHEET, CATEGORIES_HEADERS);
  sheet.appendRow([id, storeId, name, new Date().toISOString(), DEFAULT_CURRENCY]);
  getOrCreateSheet(fieldsSheetName(id), FIELDS_HEADERS);
  getOrCreateSheet(itemsSheetName(id), ITEMS_HEADERS);
  getOrCreateSheet(logSheetName(id), LOG_HEADERS);
  return { id: id, name: name };
}

function deleteCategoryById(categoryId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(CATEGORIES_SHEET, CATEGORIES_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(categoryId)) { sheet.deleteRow(i + 1); break; }
  }
  [fieldsSheetName(categoryId), itemsSheetName(categoryId), logSheetName(categoryId), historySheetName(categoryId)].forEach(function (name) {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });
}

function getSettlementCurrency(categoryId) {
  const row = getCategoryRow(categoryId);
  return row ? (row.values[4] || DEFAULT_CURRENCY) : DEFAULT_CURRENCY;
}

function setSettlementCurrency(categoryId, currency) {
  const row = getCategoryRow(categoryId);
  if (!row) throw new Error("找不到这个分类。");
  const sheet = getOrCreateSheet(CATEGORIES_SHEET, CATEGORIES_HEADERS);
  sheet.getRange(row.rowIndex, 5).setValue(currency);
}

// ---------- 每日录入字段(每个分类自己配置) ----------
function listFields(categoryId) {
  const sheet = getOrCreateSheet(fieldsSheetName(categoryId), FIELDS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push({ id: rows[i][0], name: rows[i][1], type: rows[i][2], role: rows[i][3], order: Number(rows[i][4]) });
  }
  out.sort(function (a, b) { return a.order - b.order; });
  return out;
}

function stockCheckFieldId(categoryId) {
  const fields = listFields(categoryId);
  for (let i = 0; i < fields.length; i++) {
    if (fields[i].role === "stock_check") return fields[i].id;
  }
  return null;
}

function incomingFieldId(categoryId) {
  const fields = listFields(categoryId);
  for (let i = 0; i < fields.length; i++) {
    if (fields[i].role === "incoming") return fields[i].id;
  }
  return null;
}

function addField(categoryId, name, type, role) {
  name = (name || "").trim();
  if (!name) throw new Error("字段名称不能为空。");
  if (type !== "number" && type !== "text") throw new Error("字段类型必须是 number 或 text。");
  if (role !== "stock_check" && role !== "incoming" && role !== "none") throw new Error("字段角色必须是 stock_check / incoming / none。");
  const fields = listFields(categoryId);
  if (role === "stock_check" && fields.some(function (f) { return f.role === "stock_check"; })) {
    throw new Error("已经有一个「库存盘点」字段了,每个分类只能设一个,请先把旧的删掉或改成普通字段。");
  }
  if (role === "incoming" && fields.some(function (f) { return f.role === "incoming"; })) {
    throw new Error("已经有一个「进货量」字段了,每个分类只能设一个,请先把旧的删掉或改成普通字段。");
  }
  const id = "field_" + new Date().getTime();
  const sheet = getOrCreateSheet(fieldsSheetName(categoryId), FIELDS_HEADERS);
  sheet.appendRow([id, name, type, role, fields.length]);
  return { id: id, name: name, type: type, role: role };
}

function deleteField(categoryId, fieldId) {
  const fields = listFields(categoryId).filter(function (f) { return f.id !== fieldId; });
  const sheet = getOrCreateSheet(fieldsSheetName(categoryId), FIELDS_HEADERS);
  sheet.clear();
  sheet.appendRow(FIELDS_HEADERS);
  sheet.setFrozenRows(1);
  fields.forEach(function (f, idx) { sheet.appendRow([f.id, f.name, f.type, f.role, idx]); });
}

function reorderFields(categoryId, fieldIds) {
  const fields = listFields(categoryId);
  const byId = {};
  fields.forEach(function (f) { byId[f.id] = f; });
  const sheet = getOrCreateSheet(fieldsSheetName(categoryId), FIELDS_HEADERS);
  sheet.clear();
  sheet.appendRow(FIELDS_HEADERS);
  sheet.setFrozenRows(1);
  fieldIds.forEach(function (fid, idx) {
    const f = byId[fid];
    if (f) sheet.appendRow([f.id, f.name, f.type, f.role, idx]);
  });
}

// ---------- 物品 ----------
function listItems(categoryId) {
  const sheet = getOrCreateSheet(itemsSheetName(categoryId), ITEMS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push({
      id: rows[i][0], name: rows[i][1], unit: rows[i][2], price: Number(rows[i][3]),
      currency: rows[i][4] || DEFAULT_CURRENCY, locked: rows[i][5] === true,
    });
  }
  return out;
}

function getItemRow(categoryId, itemId) {
  const sheet = getOrCreateSheet(itemsSheetName(categoryId), ITEMS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(itemId)) return { rowIndex: i + 1, values: rows[i] };
  }
  return null;
}

function addItem(categoryId, name, unit, price, currency) {
  const id = "item_" + new Date().getTime();
  const sheet = getOrCreateSheet(itemsSheetName(categoryId), ITEMS_HEADERS);
  sheet.appendRow([id, name, unit || "个", Number(price) || 0, currency || DEFAULT_CURRENCY, false]);
  return { id: id, name: name, unit: unit, price: price, currency: currency || DEFAULT_CURRENCY, locked: false };
}

function updateItem(categoryId, itemId, name, unit, price, currency) {
  const row = getItemRow(categoryId, itemId);
  if (!row) throw new Error("找不到这个物品。");
  if (row.values[5] === true) throw new Error("这个物品被锁定了,请先解锁。");
  const sheet = getOrCreateSheet(itemsSheetName(categoryId), ITEMS_HEADERS);
  sheet.getRange(row.rowIndex, 2, 1, 4).setValues([[name, unit || "个", Number(price) || 0, currency || DEFAULT_CURRENCY]]);
}

function deleteItemById(categoryId, itemId) {
  const row = getItemRow(categoryId, itemId);
  if (!row) return;
  if (row.values[5] === true) throw new Error("这个物品被锁定了,请先解锁。");
  const sheet = getOrCreateSheet(itemsSheetName(categoryId), ITEMS_HEADERS);
  sheet.deleteRow(row.rowIndex);
}

function setItemLocked(categoryId, itemId, locked) {
  const row = getItemRow(categoryId, itemId);
  if (!row) throw new Error("找不到这个物品。");
  const sheet = getOrCreateSheet(itemsSheetName(categoryId), ITEMS_HEADERS);
  sheet.getRange(row.rowIndex, 6).setValue(locked);
}

function verifyCurrentPassword(session, password) {
  if (!session || !password || hashPassword(password, session.salt) !== session.password_hash) {
    throw new Error("密码不对。");
  }
}

function getItemsWithLatestStockCheck(categoryId) {
  const items = listItems(categoryId);
  const stockFieldId = stockCheckFieldId(categoryId);
  const latest = {};
  if (stockFieldId) {
    const logSheet = getOrCreateSheet(logSheetName(categoryId), LOG_HEADERS);
    const rows = logSheet.getDataRange().getValues();
    const dataRows = [];
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][1]) dataRows.push(rows[i]);
    }
    dataRows.sort(function (a, b) { return new Date(a[0]).getTime() - new Date(b[0]).getTime(); });
    dataRows.forEach(function (row) {
      const values = parseValues(row[3]);
      const v = values[stockFieldId];
      if (v !== undefined && v !== null && v !== "") latest[row[1]] = Number(v);
    });
  }
  return { items: items, latestStockCheck: latest, stockCheckFieldId: stockFieldId };
}

// ---------- 每日记录(按 date+item_id upsert) ----------
function pickValue(newVal, oldVal) {
  return (newVal === undefined || newVal === null || newVal === "") ? oldVal : newVal;
}

function findPreviousStockCheck(rows, itemId, date, stockFieldId) {
  // 用 formatDate 的字符串比较,不用 new Date(...).getTime() 比较——后者对"日期字符串"
  // 和"表格里已存的 Date 单元格"解析时区不一致(见 formatDate 注释),会导致跨时区时
  // 边界日期判断错误。"YYYY-MM-DD" 字符串本身按字典序比较就等价于按时间先后比较。
  const targetDateStr = formatDate(date);
  let bestDateStr = null, bestVal = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) !== String(itemId)) continue;
    const rDateStr = formatDate(rows[i][0]);
    if (rDateStr >= targetDateStr) continue;
    const values = parseValues(rows[i][3]);
    const v = values[stockFieldId];
    if (v === undefined || v === null || v === "") continue;
    if (bestDateStr === null || rDateStr > bestDateStr) { bestDateStr = rDateStr; bestVal = Number(v); }
  }
  return bestVal;
}

function logEntries(categoryId, date, entries, editedBy) {
  const sheet = getOrCreateSheet(logSheetName(categoryId), LOG_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const stockFieldId = stockCheckFieldId(categoryId);
  const incFieldId = incomingFieldId(categoryId);
  const nowIso = new Date().toISOString();

  entries.forEach(function (entry) {
    const itemId = entry.itemId;
    let existingRowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(itemId) && isSameDate(rows[i][0], date)) { existingRowIndex = i; break; }
    }
    const existingValues = existingRowIndex >= 0 ? parseValues(rows[existingRowIndex][3]) : {};
    const incomingValues = entry.values || {};
    const merged = Object.assign({}, existingValues);
    Object.keys(incomingValues).forEach(function (fieldId) {
      merged[fieldId] = pickValue(incomingValues[fieldId], existingValues[fieldId]);
    });

    let usage = "", cost = "";
    if (stockFieldId) {
      const stockValRaw = merged[stockFieldId];
      const stockVal = (stockValRaw === undefined || stockValRaw === null || stockValRaw === "") ? null : Number(stockValRaw);
      if (stockVal !== null && !isNaN(stockVal)) {
        const prevStock = findPreviousStockCheck(rows, itemId, date, stockFieldId);
        if (prevStock !== null) {
          const incomingRaw = incFieldId ? merged[incFieldId] : null;
          const incomingVal = (incomingRaw === undefined || incomingRaw === null || incomingRaw === "") ? 0 : (Number(incomingRaw) || 0);
          usage = prevStock + incomingVal - stockVal;
          cost = usage * (Number(entry.price) || 0);
        }
      }
    }

    const rowValues = [date, itemId, entry.itemName || "", JSON.stringify(merged), usage, cost, nowIso, editedBy || ""];
    if (existingRowIndex >= 0) {
      // 覆盖之前先把旧版本存进历史表——"任何 edit 都可以 recover"靠的是这一份快照,
      // 不是靠"不覆盖、一直往下加行"(那样就是这次修的重复行 bug)。
      getOrCreateSheet(historySheetName(categoryId), HISTORY_HEADERS).appendRow(rows[existingRowIndex]);
      sheet.getRange(existingRowIndex + 1, 1, 1, rowValues.length).setValues([rowValues]);
      rows[existingRowIndex] = rowValues;
    } else {
      sheet.appendRow(rowValues);
      rows.push(rowValues);
    }
  });
}

function getRecentLog(categoryId, days) {
  const sheet = getOrCreateSheet(logSheetName(categoryId), LOG_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][1]) continue;
    if (new Date(rows[i][0]).getTime() < cutoff.getTime()) continue;
    out.push({
      date: formatDate(rows[i][0]), itemId: rows[i][1], itemName: rows[i][2],
      values: parseValues(rows[i][3]),
      usage: rows[i][4] === "" ? null : Number(rows[i][4]),
      cost: rows[i][5] === "" ? null : Number(rows[i][5]),
      editedAt: rows[i][6] || null,
      editedBy: rows[i][7] || null,
    });
  }
  out.sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
  return out;
}

function getEntryHistory(categoryId, itemId, date) {
  const sheet = getOrCreateSheet(historySheetName(categoryId), HISTORY_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][1]) continue;
    if (String(rows[i][1]) !== String(itemId)) continue;
    if (!isSameDate(rows[i][0], date)) continue;
    out.push({
      values: parseValues(rows[i][3]),
      usage: rows[i][4] === "" ? null : Number(rows[i][4]),
      cost: rows[i][5] === "" ? null : Number(rows[i][5]),
      editedAt: rows[i][6] || null,
      editedBy: rows[i][7] || null,
    });
  }
  out.sort(function (a, b) { return new Date(b.editedAt) - new Date(a.editedAt); });
  return out;
}

function deleteLogDay(categoryId, date) {
  const sheet = getOrCreateSheet(logSheetName(categoryId), LOG_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] && isSameDate(rows[i][0], date)) sheet.deleteRow(i + 1);
  }
}

// ---------- 汇率(服务器代理,浏览器 fetch 没法自定义 User-Agent) ----------
function getExchangeRates(baseCurrency) {
  const url = "https://api.frankfurter.dev/v1/latest?base=" + encodeURIComponent(baseCurrency);
  const resp = UrlFetchApp.fetch(url, { headers: { "User-Agent": "Mozilla/5.0" }, muteHttpExceptions: true });
  const data = JSON.parse(resp.getContentText());
  return { date: data.date, rates: data.rates || {} };
}

// ---------- doGet / doPost ----------
function doGet(e) {
  try {
    const p = e.parameter;
    const action = p.action;
    if (action === "hasAdmin") return jsonResponse({ hasAdmin: hasAdmin() });

    const session = resolveSession(p.token);
    if (!session) return jsonResponse({ error: "请重新登录。" });

    if (action === "listStores") return jsonResponse({ stores: listStores() });
    if (action === "listCategories") return jsonResponse({ categories: listCategories(p.storeId) });
    if (action === "getSettlementCurrency") return jsonResponse({ currency: getSettlementCurrency(p.categoryId) });
    if (action === "listFields") return jsonResponse({ fields: listFields(p.categoryId) });
    if (action === "listItems") return jsonResponse(getItemsWithLatestStockCheck(p.categoryId));
    if (action === "getLog") return jsonResponse({ log: getRecentLog(p.categoryId, parseInt(p.days || "30", 10)) });
    if (action === "getEntryHistory") return jsonResponse({ history: getEntryHistory(p.categoryId, p.itemId, p.date) });
    if (action === "getExchangeRates") return jsonResponse(getExchangeRates(p.base));
    if (action === "listUsers") { requireAdmin(session); return jsonResponse({ users: listUsers() }); }
    return jsonResponse({ error: "unknown action" });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (lockErr) {
    return jsonResponse({ error: "系统繁忙,请重试。" });
  }
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    if (action === "registerAdmin") return jsonResponse(registerAdmin(payload.username, payload.password, payload.displayName));
    if (action === "login") return jsonResponse(login(payload.username, payload.password));

    const session = resolveSession(payload.token);
    if (!session) return jsonResponse({ error: "请重新登录。" });

    if (action === "logout") { deleteSession(payload.token); return jsonResponse({ status: "ok" }); }

    if (action === "createEmployee") { requireAdmin(session); return jsonResponse({ user: createEmployee(payload.username, payload.password, payload.displayName, payload.storeId) }); }
    if (action === "deleteUser") { requireAdmin(session); deleteUserById(payload.userId); return jsonResponse({ status: "ok" }); }
    if (action === "resetPassword") { requireAdmin(session); resetPassword(payload.userId, payload.newPassword); return jsonResponse({ status: "ok" }); }

    if (action === "createStore") { requireAdmin(session); return jsonResponse({ store: createStore(payload.name) }); }
    if (action === "deleteStore") { requireAdmin(session); deleteStoreById(payload.storeId); return jsonResponse({ status: "ok" }); }

    if (action === "createCategory") { requireAdmin(session); return jsonResponse({ category: createCategory(payload.storeId, payload.name) }); }
    if (action === "deleteCategory") { requireAdmin(session); deleteCategoryById(payload.categoryId); return jsonResponse({ status: "ok" }); }
    if (action === "setSettlementCurrency") { requireAdmin(session); setSettlementCurrency(payload.categoryId, payload.currency); return jsonResponse({ status: "ok" }); }

    if (action === "addItem") { requireAdmin(session); return jsonResponse({ item: addItem(payload.categoryId, payload.name, payload.unit, payload.price, payload.currency) }); }
    if (action === "updateItem") { requireAdmin(session); updateItem(payload.categoryId, payload.itemId, payload.name, payload.unit, payload.price, payload.currency); return jsonResponse({ status: "ok" }); }
    if (action === "deleteItem") { requireAdmin(session); deleteItemById(payload.categoryId, payload.itemId); return jsonResponse({ status: "ok" }); }
    if (action === "lockItem") { requireAdmin(session); setItemLocked(payload.categoryId, payload.itemId, true); return jsonResponse({ status: "ok" }); }
    if (action === "unlockItem") {
      requireAdmin(session);
      verifyCurrentPassword(session, payload.password);
      setItemLocked(payload.categoryId, payload.itemId, false);
      return jsonResponse({ status: "ok" });
    }

    if (action === "addField") { requireAdmin(session); return jsonResponse({ field: addField(payload.categoryId, payload.name, payload.fieldType, payload.role) }); }
    if (action === "deleteField") { requireAdmin(session); deleteField(payload.categoryId, payload.fieldId); return jsonResponse({ status: "ok" }); }
    if (action === "reorderFields") { requireAdmin(session); reorderFields(payload.categoryId, payload.fieldIds); return jsonResponse({ status: "ok" }); }

    if (action === "saveLogEntries") {
      requireStoreForCategory(session, payload.categoryId);
      logEntries(payload.categoryId, payload.date, payload.entries, session.display_name || session.username);
      return jsonResponse({ status: "ok" });
    }

    if (action === "deleteLogDay") {
      requireAdmin(session);
      deleteLogDay(payload.categoryId, payload.date);
      return jsonResponse({ status: "ok" });
    }

    return jsonResponse({ error: "unknown action" });
  } catch (err) {
    return jsonResponse({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}
