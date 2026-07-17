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
// total_stock: 7th column, added 2026-07-12 — 老的物品行没有这一列,
// getDataRange() 读出来就是 undefined,跟 locked 那次一样按"没设置"处理,不用迁移。
const ITEMS_HEADERS = ["id", "name", "unit", "price", "currency", "locked", "total_stock"];
const LOG_HEADERS = ["date", "item_id", "item_name", "values_json", "usage", "cost", "edited_at", "edited_by"];
const HISTORY_HEADERS = ["date", "item_id", "item_name", "values_json", "usage", "cost", "edited_at", "edited_by"];
// 自定义条目(2026-07-15 加):跟物品无关、一天一份的自由文本,比如"当日营业额"/
// "应到货备注"。不需要 unit/type/role,只有名字——所以是三张独立的小表,不复用
// FIELDS/ITEMS 那一套。
const NOTE_FIELDS_HEADERS = ["id", "name", "order"];
const NOTE_LOG_HEADERS = ["date", "note_id", "value", "edited_at", "edited_by"];
const NOTE_HISTORY_HEADERS = ["date", "note_id", "value", "edited_at", "edited_by"];

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
function notesSheetName(categoryId) { return "notes_" + categoryId; }
function notelogSheetName(categoryId) { return "notelog_" + categoryId; }
function notehistorySheetName(categoryId) { return "notehistory_" + categoryId; }

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

// 管理员看全部账号;经理/班长(2026-07-17 新增的两级)只能看自己那家门店的
// 账号——他们本来就被锁在自己店里,没道理让他们看到别的店有谁。
function listUsers(session) {
  const sheet = getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    const u = rowToUser(rows[i]);
    if (session.role !== "admin" && String(u.store_id) !== String(session.store_id)) continue;
    out.push(publicUser(u));
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

// 管理员能建经理/班长/员工三级里的任何一个,而且能指定任意门店;经理/班长
// (level 2+)也能建账号,但只能建"员工"这一级、门店强制是自己所在的那家,
// 不管前端传了什么角色/门店——这两个值绝对不能信任客户端,不然一个班长
// 改改请求参数就能给自己或者别人建一个管理员账号,或者跨店建号。
function createEmployee(session, username, password, displayName, storeId, role) {
  username = (username || "").trim();
  if (!username || !password) throw new Error("用户名和密码不能为空。");
  if (session.role === "admin") {
    role = (role === "manager" || role === "shift_leader") ? role : "employee";
  } else {
    role = "employee";
    storeId = session.store_id;
  }
  if (!storeId) throw new Error("请给这个账号选一个门店。");
  if (findUserByUsername(username)) throw new Error("用户名「" + username + "」已经被使用。");
  const salt = genSalt();
  const id = "user_" + new Date().getTime();
  const name = (displayName || "").trim() || username;
  const sheet = getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
  sheet.appendRow([id, username, hashPassword(password, salt), salt, role, name, storeId]);
  return { id: id, username: username, displayName: name, role: role, storeId: storeId };
}

// 管理员能删任何非管理员账号;经理(level 3)只能删自己店里"员工"这一级的
// 账号——不能删同店的经理/班长(防止越权删同级或上级),也不能删别的店的。
// 班长(level 2)权限不够,连调用这个函数的资格都没有(在 dispatch 那层就被
// requireLevel 挡掉了)。
function deleteUserById(session, userId) {
  const sheet = getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(userId)) {
      const targetRole = rows[i][4];
      const targetStoreId = rows[i][6];
      if (targetRole === "admin") throw new Error("不能删除管理员账号。");
      if (session.role !== "admin" && (targetRole !== "employee" || String(targetStoreId) !== String(session.store_id))) {
        throw new Error("没有权限删除这个账号。");
      }
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
  // 经理/班长(2026-07-17 新增)跟员工一样,都是被锁在自己那家门店里的——
  // 前端 afterLogin() 靠"是不是 admin"分流,只要不是 admin 就必须带上
  // storeName 才能正常跳进"选择分类"页面,不能只判 role === "employee"。
  if (user.role !== "admin" && user.store_id) {
    const store = getStoreById(user.store_id);
    result.storeName = store ? store.name : null;
  }
  return { user: result, token: token };
}

// ---------- 权限校验 ----------
// 四级权限,数字越大权限越高,每个操作要求一个"至少要达到"的门槛——不是每级
// 单独配一份权限列表,是严格的层级(2026-07-17 跟 Kevin 核对过实际需求,
// 管理员≥经理≥班长≥员工,凡是班长能做的经理和管理员也都能做,没有例外)。
// 除了管理员,其余三级都被锁在自己所属的那一家门店里——门店范围检查见
// requireStoreForCategory/requireStoreAccess,判断标准是"是不是 admin",
// 不再是"是不是 employee"。
const ROLE_LEVELS = { admin: 4, manager: 3, shift_leader: 2, employee: 1 };
const ROLE_LABELS = { admin: "管理员", manager: "经理", shift_leader: "班长", employee: "员工" };

function requireLevel(session, minLevel) {
  if (!session || (ROLE_LEVELS[session.role] || 0) < minLevel) {
    throw new Error("权限不够,做不了这个操作。");
  }
}

function requireAdmin(session) {
  requireLevel(session, ROLE_LEVELS.admin);
}

function requireStoreForCategory(session, categoryId) {
  if (session.role !== "admin") {
    const storeId = getCategoryStoreId(categoryId);
    if (String(storeId) !== String(session.store_id)) throw new Error("没有权限操作这个门店的数据。");
  }
}

// 跟 requireStoreForCategory 是同一条规则,只是直接给 storeId(listCategories/
// listStores 用,这两个接口拿到的是门店 id,不是分类 id,没法走 getCategoryStoreId
// 反查)。2026-07-16 补的——之前只有 doPost 的三个写操作挡了员工跨店,doGet 的
// 读接口(listItems/getLog/listNoteFields 等)完全没查 session 是不是这个店的,
// 员工账号只要登录就能拿别的门店的物品/库存/进销记录/自定义条目(比如营业额)/
// 编辑历史——不需要猜 id,listStores/listCategories 本身就会把所有门店和分类原样
// 吐出来。管理员不受这条限制,继续能看全部门店。
function requireStoreAccess(session, storeId) {
  if (session.role !== "admin" && String(storeId) !== String(session.store_id)) {
    throw new Error("没有权限查看这个门店的数据。");
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

// 员工账号只应该看到自己绑定的那一家门店——之前 listStores() 原样返回全部,
// 见上面 requireStoreAccess 的说明。管理员没有限制,继续能看全部门店。
function listStoresForSession(session) {
  const all = listStores();
  if (session.role === "admin") return all;
  return all.filter(function (s) { return String(s.id) === String(session.store_id); });
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
  [fieldsSheetName(categoryId), itemsSheetName(categoryId), logSheetName(categoryId), historySheetName(categoryId),
   notesSheetName(categoryId), notelogSheetName(categoryId), notehistorySheetName(categoryId)].forEach(function (name) {
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
    const rawTotal = rows[i][6];
    out.push({
      id: rows[i][0], name: rows[i][1], unit: rows[i][2], price: Number(rows[i][3]),
      currency: rows[i][4] || DEFAULT_CURRENCY, locked: rows[i][5] === true,
      totalStock: (rawTotal === "" || rawTotal === undefined || rawTotal === null) ? null : Number(rawTotal),
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

function addItem(categoryId, name, unit, price, currency, startingStock) {
  const id = "item_" + new Date().getTime();
  const sheet = getOrCreateSheet(itemsSheetName(categoryId), ITEMS_HEADERS);
  const hasStarting = startingStock !== undefined && startingStock !== null && startingStock !== "";
  // 起始库存直接写进物品自己的 total_stock 字段当基准,不再造一条假的历史日期
  // 记录——这样这个物品第一次真正盘点就有"上一次"可以减,能算出消耗,而且
  // "全部记录"日历里不会平白多出一天,删这个物品的真实记录也不会连带把这个
  // 数字弄没(以前那版靠伪造"昨天"的日志行当基准,两个问题都踩过)。
  sheet.appendRow([id, name, unit || "个", Number(price) || 0, currency || DEFAULT_CURRENCY, false, hasStarting ? Number(startingStock) : ""]);
  return {
    id: id, name: name, unit: unit, price: price, currency: currency || DEFAULT_CURRENCY, locked: false,
    totalStock: hasStarting ? Number(startingStock) : null,
  };
}

// 批量添加物品(Excel 导入用)——一次性 setValues 写入,不是循环 appendRow,
// 避免几十上百个物品导入时一行一次单独写表格。按名称(去空格、忽略大小写)
// 跟当前分类里已有的物品去重,同一批 excel 里自己重复的名字也会去重,重复的
// 名字进 skipped 不导入,不覆盖已有物品。不接受起始库存——用户要求这批物品
// 导入后自己去点"目前总库存"手动设置,保持跟单个添加时"不填起始库存"一样的
// 空白状态(total_stock 留空,第一次真正盘点前显示"暂无记录")。
function addItemsBulk(categoryId, items) {
  const sheet = getOrCreateSheet(itemsSheetName(categoryId), ITEMS_HEADERS);
  const existingNames = new Set(listItems(categoryId).map(it => String(it.name).trim().toLowerCase()));
  const baseTime = new Date().getTime();
  const added = [];
  const skipped = [];
  const rows = [];
  (items || []).forEach((raw, idx) => {
    const name = String((raw && raw.name) || "").trim();
    if (!name) return;
    const key = name.toLowerCase();
    if (existingNames.has(key)) { skipped.push(name); return; }
    existingNames.add(key);
    const id = "item_" + (baseTime + idx);
    const unit = (raw.unit && String(raw.unit).trim()) || "个";
    const price = Number(raw.price) || 0;
    const currency = raw.currency || DEFAULT_CURRENCY;
    rows.push([id, name, unit, price, currency, false, ""]);
    added.push({ id: id, name: name, unit: unit, price: price, currency: currency, locked: false, totalStock: null });
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, ITEMS_HEADERS.length).setValues(rows);
  }
  return { added: added, skipped: skipped };
}

// AI 智能识别 Excel 导入——跟 addItemsBulk 配合用:前端先把整张表格原样解析成
// 二维数组(不假设任何固定列),这里把原始数据交给 Claude API 提取"名称/单位/
// 单价",不管表头怎么写、列顺序、有没有多余的列。返回结构化结果后前端会先
// 给用户看一遍预览,确认了才真的调 addItemsBulk 写进表格——不在这一步直接
// 写数据,AI 判断有误也不会污染库存表。
//
// API key 存在这份 Apps Script 项目的 Script Property 里(不是写死在代码里),
// 只有 Kevin 自己的 Google 账号登进 Apps Script 编辑器才能看到/改——员工账号、
// 前端网页、查看网页源代码都碰不到这个 key,它从来不会传到浏览器。设置方法:
// Apps Script 编辑器左侧齿轮图标"项目设置" -> "脚本属性" -> 新增属性,
// 名称填 ANTHROPIC_API_KEY,值填真正的 key。
function aiParseImportRows(rawRows) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) {
    throw new Error("还没有配置 AI 识别用的 API key——去 Apps Script 编辑器左侧齿轮"
      + "图标「项目设置」->「脚本属性」,新增一个名称叫 ANTHROPIC_API_KEY 的属性,"
      + "值填 Anthropic API key,保存后再试。");
  }

  const systemPrompt = "你是一个从杂乱的 Excel/表格原始数据里提取物品清单的助手。"
    + "输入是表格的原始二维数组,每个内层数组是一行的所有单元格,可能包含标题行、"
    + "空行、说明性文字、多余的列,列的顺序和语言也不一定固定。找出真正代表"
    + "\"物品\"的每一行,提取它的名称(必须有)、单位(比如 个/L/kg,没有就留空)、"
    + "单价(数字,没有就留空)。跳过标题行、空行、汇总行、任何不是具体物品的行。"
    + "不要编造数据里没有的信息。";

  const payload = {
    model: "claude-opus-4-8",
    max_tokens: 4096,
    system: systemPrompt,
    messages: [{ role: "user", content: "原始表格数据(JSON 二维数组,每个内层数组是一行):\n" + JSON.stringify(rawRows) }],
    output_config: {
      format: {
        type: "json_schema",
        schema: {
          type: "object",
          properties: {
            items: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  name: { type: "string" },
                  unit: { type: "string" },
                  price: { type: "number" },
                },
                required: ["name"],
                additionalProperties: false,
              },
            },
          },
          required: ["items"],
          additionalProperties: false,
        },
      },
    },
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: { "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const bodyText = response.getContentText();
  if (status !== 200) {
    throw new Error("AI 识别失败(状态码 " + status + "): " + bodyText);
  }
  const body = JSON.parse(bodyText);
  if (body.stop_reason === "refusal") {
    throw new Error("AI 拒绝处理这份数据,换成不联网的固定格式导入试试。");
  }
  const textBlock = (body.content || []).filter(function (b) { return b.type === "text"; })[0];
  if (!textBlock) {
    throw new Error("AI 没有返回可解析的结果,请重试。");
  }
  const parsed = JSON.parse(textBlock.text);
  return { items: parsed.items || [] };
}

// 直接改 total_stock 这个数字,不碰任何一天的日志行——不产生消耗记录,只是
// 纠正"现在库存应该是多少"。只有管理员能调这个(前端隐藏+这里 requireAdmin
// 双重把关,见 dispatch 里的 setTotalStock 分支);正常员工录入当天盘点走的
// 是 logEntries,那条路径也会自动把新数值推进 total_stock,但那不算"直接改"。
function setTotalStock(categoryId, itemId, value) {
  const row = getItemRow(categoryId, itemId);
  if (!row) throw new Error("找不到这个物品。");
  const sheet = getOrCreateSheet(itemsSheetName(categoryId), ITEMS_HEADERS);
  const v = (value === null || value === undefined || value === "") ? "" : Number(value);
  sheet.getRange(row.rowIndex, 7).setValue(v);
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

// 批量删除物品(前端多选一次性删,导入导错名字之后清理用)——跟单个删除
// (deleteItemById)同样的锁定保护:锁定的物品直接跳过不强删,收集进
// skippedLocked 让前端提示,不像单个删除那样直接报错中断。按行号从大到小
// 删,不然从上往下删的过程中,后面几行的真实行号会跟着往上移、错位删错行。
function deleteItemsBulk(categoryId, itemIds) {
  const sheet = getOrCreateSheet(itemsSheetName(categoryId), ITEMS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const idSet = new Set((itemIds || []).map(function (id) { return String(id); }));
  const rowIndicesToDelete = [];
  const deletedNames = [];
  const skippedLocked = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0] || !idSet.has(String(rows[i][0]))) continue;
    if (rows[i][5] === true) { skippedLocked.push(rows[i][1]); continue; }
    rowIndicesToDelete.push(i + 1);
    deletedNames.push(rows[i][1]);
  }
  rowIndicesToDelete.sort(function (a, b) { return b - a; }).forEach(function (rowIndex) { sheet.deleteRow(rowIndex); });
  return { deletedCount: deletedNames.length, deletedNames: deletedNames, skippedLocked: skippedLocked };
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

// "目前总库存"现在直接读物品自己的 total_stock 字段,不再扫日志表找"最新一条
// 盘点记录"——这样删掉某一天的日志不会连带把这个数字弄没(total_stock 是独立
// 存的,只会被起始库存/正常录入/管理员直接改这三种操作更新,删日志不影响它)。
function getItemsWithLatestStockCheck(categoryId) {
  const items = listItems(categoryId);
  const stockFieldId = stockCheckFieldId(categoryId);

  // 自愈式迁移:total_stock 这一列是这次改动才加的,在这之前就存在的老物品
  // (哪怕日志表里明明有真实的盘点历史)第7列本来就是空的——不退回去扫一次
  // 日志表的话,这些老物品会平白显示"暂无记录",总库存磁贴也没法点(见
  // [[project_inventory_app]] 里"总库存变成暂无记录了"那次的排查)。只在真的
  // 遇到 total_stock 为空、但日志表有值的物品时才扫,扫到就顺手写回
  // total_stock,下次就不用再扫了——不需要手动跑一次性迁移脚本。
  if (stockFieldId) {
    const needsBackfill = items.filter(function (it) { return it.totalStock === null || it.totalStock === undefined; });
    if (needsBackfill.length > 0) {
      const logSheet = getOrCreateSheet(logSheetName(categoryId), LOG_HEADERS);
      const rows = logSheet.getDataRange().getValues();
      const dataRows = [];
      for (let i = 1; i < rows.length; i++) { if (rows[i][1]) dataRows.push(rows[i]); }
      dataRows.sort(function (a, b) { return new Date(a[0]).getTime() - new Date(b[0]).getTime(); });
      const scannedLatest = {};
      dataRows.forEach(function (row) {
        const values = parseValues(row[3]);
        const v = values[stockFieldId];
        if (v !== undefined && v !== null && v !== "") scannedLatest[row[1]] = Number(v);
      });
      needsBackfill.forEach(function (it) {
        if (scannedLatest[it.id] !== undefined) {
          it.totalStock = scannedLatest[it.id];
          setTotalStock(categoryId, it.id, scannedLatest[it.id]);
        }
      });
    }
  }

  const latest = {};
  items.forEach(function (it) {
    if (it.totalStock !== null && it.totalStock !== undefined) latest[it.id] = it.totalStock;
  });
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
  // 返回 {dateStr, value},不只是 value——调用方要拿这个日期去算"这次真正盘点
  // 之后、这次盘点之前"漏了多少天单独录的进货(见 sumIncomingBetween)。
  const targetDateStr = formatDate(date);
  let best = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) !== String(itemId)) continue;
    const rDateStr = formatDate(rows[i][0]);
    if (rDateStr >= targetDateStr) continue;
    const values = parseValues(rows[i][3]);
    const v = values[stockFieldId];
    if (v === undefined || v === null || v === "") continue;
    if (best === null || rDateStr > best.dateStr) best = { dateStr: rDateStr, value: Number(v) };
  }
  return best;
}

// 上一次真正盘点(prevCheck.dateStr)严格早于、这次盘点严格早于 targetDateStr 的
// 所有单独进货行(只填了进货没跟着盘点的那种)加总——这样即使进货跟盘点没在
// 同一天录,消耗算出来也不会漏掉中间那几天的进货。上下界都不含,因为"上一次
// 盘点当天"和"这次盘点当天"各自的进货已经在调用方的 usage 公式里单独算过一次
// 了,这里再算就是重复计入。
function sumIncomingBetween(rows, itemId, afterDateStr, beforeDateStrExclusive, incFieldId) {
  if (!incFieldId) return 0;
  let sum = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) !== String(itemId)) continue;
    const rDateStr = formatDate(rows[i][0]);
    if (afterDateStr !== null && rDateStr <= afterDateStr) continue;
    if (rDateStr >= beforeDateStrExclusive) continue;
    const values = parseValues(rows[i][3]);
    const v = values[incFieldId];
    if (v === undefined || v === null || v === "" || isNaN(Number(v))) continue;
    sum += Number(v);
  }
  return sum;
}

// 如果这条记录(这个物品、这一天)之前已经存过一次、当时是"只填了进货没填
// 盘点"的状态,那次保存已经把这个进货量加进过总库存了(见 logEntries 里
// onlyIncoming 分支)。这次编辑重新计算基准/总库存前,必须先把那次的贡献减
// 掉,不然同一笔进货会被算两次——不管是"这次编辑改了进货数字"还是"这次编辑
// 补上了盘点值",都得先退回这一步。existingValues 在没有既有记录时是 {},
// 这时天然返回 0,调用方不用额外判断 existingRowIndex。
function priorIncomingOnlyBump(existingValues, incFieldId, stockFieldId) {
  const oldStockRaw = existingValues[stockFieldId];
  const oldHadStock = !(oldStockRaw === undefined || oldStockRaw === null || oldStockRaw === "");
  if (oldHadStock) return 0;
  const oldIncomingRaw = incFieldId ? existingValues[incFieldId] : null;
  if (oldIncomingRaw === undefined || oldIncomingRaw === null || oldIncomingRaw === "" || isNaN(Number(oldIncomingRaw))) return 0;
  return Number(oldIncomingRaw);
}

// findPreviousStockCheck 的镜像版本:找"严格晚于" targetDateStr 的、日期最早的那个
// 真实盘点行——editRow/deleteLogDay 改动了更早的一天之后,要用这个函数找出"下一次
// 真实盘点"存的 usage/cost 是不是需要重算(见 recomputeForwardAfterChange)。
function findNextStockCheck(rows, itemId, targetDateStr, stockFieldId) {
  let best = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) !== String(itemId)) continue;
    const rDateStr = formatDate(rows[i][0]);
    if (rDateStr <= targetDateStr) continue;
    const values = parseValues(rows[i][3]);
    const v = values[stockFieldId];
    if (v === undefined || v === null || v === "") continue;
    if (best === null || rDateStr < best.dateStr) best = { dateStr: rDateStr, rowIndex: i, value: Number(v) };
  }
  return best;
}

// 修复"编辑/删除一个更早的历史日期后,下一次真实盘点存的消耗数字变成基于一个
// 已经不存在的基准算出来的错误值"这个 bug(2026-07-15 Kevin 提出的场景)——
// total_stock 不受影响(它只在触碰"当前最新一天"时才更新),但 usage/cost 是保存
// 那一刻算好就存死的,历史日期一变就该跟着重算,不然会一直显示错的消耗数字。
// 只需要往前重算一步:下一次真实盘点的 stockVal 本身没变,只是它的计算基准
// (prevStock)可能变了;再往后一行用的基准是"下一次真实盘点的 stockVal"本身,
// 这个值没变,所以不需要继续级联下去。
function recomputeForwardAfterChange(categoryId, itemId, changedDateStr) {
  const stockFieldId = stockCheckFieldId(categoryId);
  if (!stockFieldId) return;
  const incFieldId = incomingFieldId(categoryId);
  const logSheet = getOrCreateSheet(logSheetName(categoryId), LOG_HEADERS);
  const rows = logSheet.getDataRange().getValues();
  const next = findNextStockCheck(rows, itemId, changedDateStr, stockFieldId);
  if (!next) return; // 这个物品在这之后没有真实盘点过,下次真正盘点会自然用上正确的链条

  const prevCheck = findPreviousStockCheck(rows, itemId, next.dateStr, stockFieldId);
  let usage = "", cost = "";
  if (prevCheck !== null) {
    // 特意不像 logEntries 那样在"找不到更早盘点"时退回去用物品的 total_stock
    // 字段当基准——那个字段这时候早就被"下一次真实盘点"自己(或者更晚的其它
    // 记录)写过了,已经不是"这次改动之前"的原始起始库存,拿它当 prevStock
    // 就是拿这一行自己(或者跟它无关的更晚数据)当自己的基准,会算出一个看起来
    // 正常、其实毫无意义的假消耗数字。删/改之后如果真的没有更早的真实盘点了,
    // 老老实实退回"暂无消耗数据"(usage/cost 留空)——比编一个错的数字安全。
    const prevStock = prevCheck.value + sumIncomingBetween(rows, itemId, prevCheck.dateStr, next.dateStr, incFieldId);
    const values = parseValues(rows[next.rowIndex][3]);
    const incomingRaw = incFieldId ? values[incFieldId] : null;
    const hasIncoming = !(incomingRaw === undefined || incomingRaw === null || incomingRaw === "") && !isNaN(Number(incomingRaw));
    const incomingVal = hasIncoming ? Number(incomingRaw) : 0;
    usage = prevStock + incomingVal - next.value;
    const itemRow = getItemRow(categoryId, itemId);
    const price = itemRow ? (Number(itemRow.values[3]) || 0) : 0;
    cost = usage * price;
  }
  // 不管算没算出来都要写回去(哪怕是清空)——万一这一行之前显示的是一个基于
  // 已经不存在的基准算出来的旧数字,这次也要把它正确地清成"暂无消耗数据"。
  logSheet.getRange(next.rowIndex + 1, 5, 1, 2).setValues([[usage, cost]]);
}

function logEntries(categoryId, date, entries, editedBy, confirmOverwrite) {
  const sheet = getOrCreateSheet(logSheetName(categoryId), LOG_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const stockFieldId = stockCheckFieldId(categoryId);
  const incFieldId = incomingFieldId(categoryId);
  const nowIso = new Date().toISOString();

  // 先整体检查一遍有没有"这个物品今天已经有人录过"——批量保存(今天的记录
  // 表格一次存好几个物品)要么整批一起提示、要么整批一起存,不做部分保存,
  // 不然没法跟前端解释"这几个存了那几个没存"。只有前端明确带上
  // confirmOverwrite(用户已经看过提示、确认要覆盖)才跳过这一步直接写。
  if (!confirmOverwrite) {
    const conflicts = [];
    entries.forEach(function (entry) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][1]) === String(entry.itemId) && isSameDate(rows[i][0], date)) {
          conflicts.push({
            itemId: entry.itemId, itemName: entry.itemName || rows[i][2],
            editedBy: rows[i][7] || "未知", editedAt: rows[i][6] || null,
          });
          break;
        }
      }
    });
    if (conflicts.length > 0) return { needsConfirm: true, conflicts: conflicts };
  }

  // 物品表只读一次、这一批 entries 共用——total_stock 的读取/写入都在这份
  // 内存数据上做完再统一落盘。之前每个 entry 各自调 getItemRow/setTotalStock,
  // 都会重新把整张 items 表读一遍,一次保存好几个物品时越存越慢,是这次顺手
  // 修的性能问题(见 [[project_inventory_app]] 加载慢的排查)。
  const itemsSheet = stockFieldId ? getOrCreateSheet(itemsSheetName(categoryId), ITEMS_HEADERS) : null;
  const itemRows = itemsSheet ? itemsSheet.getDataRange().getValues() : null;
  const itemRowIndexById = {};
  if (itemRows) {
    for (let i = 1; i < itemRows.length; i++) {
      if (itemRows[i][0]) itemRowIndexById[String(itemRows[i][0])] = i;
    }
  }

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
    let needsForwardRecompute = false; // 这次改动的是不是这个物品的历史日期,见下面 recomputeForwardAfterChange
    if (stockFieldId) {
      const stockValRaw = merged[stockFieldId];
      const stockVal = (stockValRaw === undefined || stockValRaw === null || stockValRaw === "") ? null : Number(stockValRaw);
      const incomingRaw = incFieldId ? merged[incFieldId] : null;
      const hasIncoming = !(incomingRaw === undefined || incomingRaw === null || incomingRaw === "") && !isNaN(Number(incomingRaw));
      const incomingVal = hasIncoming ? Number(incomingRaw) : 0;
      const itemRowIdx = itemRowIndexById[String(itemId)];
      // 只有这条记录是这个物品目前"最新"的一条(日期不早于其它所有记录),才
      // 允许它推进 total_stock——补录/编辑更早的历史日期不该覆盖"现在"的
      // 总库存,那个数字永远该反映"最近一次真实变化",不是"最近一次被编辑的"。
      const targetDateStr = formatDate(date);
      const isLatestForItem = !rows.some(function (r, idx) {
        return idx > 0 && idx !== existingRowIndex && String(r[1]) === String(itemId) && formatDate(r[0]) > targetDateStr;
      });
      needsForwardRecompute = !isLatestForItem;
      if (stockVal !== null && !isNaN(stockVal)) {
        // 基准 = 上一次真正盘点的值 + 那次盘点之后、这次盘点之前(都不含)
        // 单独录入的所有进货累加——这样即使某天只录了进货没顺手盘点,后面
        // 真正盘点时也不会漏算那笔进货。日期链扫描不受"这条记录是不是编辑
        // 已有的一天"影响,天然对重复编辑安全,不会重复计入。
        const prevCheck = findPreviousStockCheck(rows, itemId, date, stockFieldId);
        let prevStock = null;
        if (prevCheck !== null) {
          prevStock = prevCheck.value + sumIncomingBetween(rows, itemId, prevCheck.dateStr, targetDateStr, incFieldId);
        } else if (itemRowIdx !== undefined && isLatestForItem && existingRowIndex === -1) {
          // 真的一条历史盘点记录都没有(全新物品的第一次盘点):退回去用物品
          // 自己的 total_stock 字段当基准——这就是起始库存/纯进货累加生效的
          // 地方。这个字段可能已经被之前"只填进货"的保存累加过,如果这次编辑
          // 的正是那一天的记录,要先退回那次的贡献,不然会被重复计入一次。
          // 加了 isLatestForItem && existingRowIndex === -1 这两个条件
          // (2026-07-15):必须同时满足"这是一条全新插入的记录"(不是在编辑
          // 已经存在的行)和"后面没有更晚的真实盘点",才能保证 total_stock
          // 从来没被任何一行动过、还是原始起始状态。少了任一个条件都可能读到
          // 一个已经被别的行(不管是这一行自己更早的旧值,还是后面更晚的真实
          // 盘点)推进过的 total_stock,拿来当基准就是用别的时间点的数字反推
          // 现在,算出一个看起来正常、其实毫无意义的假消耗数字——这种情况
          // 老实退回"暂无消耗数据"更安全,跟 recomputeForwardAfterChange 里
          // 同一个道理。
          const rawTotal = itemRows[itemRowIdx][6];
          if (rawTotal !== "" && rawTotal !== undefined && rawTotal !== null) {
            prevStock = Number(rawTotal) - priorIncomingOnlyBump(existingValues, incFieldId, stockFieldId);
          }
        }
        if (prevStock !== null) {
          usage = prevStock + incomingVal - stockVal;
          cost = usage * (Number(entry.price) || 0);
        }
        if (isLatestForItem && itemRowIdx !== undefined) {
          itemsSheet.getRange(itemRowIdx + 1, 7).setValue(stockVal);
          itemRows[itemRowIdx][6] = stockVal; // 保持内存副本同步,防止同一批里同个物品出现两次时读到旧值
        }
      } else if (hasIncoming && incomingVal !== 0 && isLatestForItem && itemRowIdx !== undefined) {
        // 只填了进货、没填盘点:这天没法算消耗(缺实际盘点数),但进货是确定
        // 发生的事实,不用等哪天顺手一起盘点才生效——直接累加进现有总库存,
        // 免得"我明明填了进货,总库存却没变"(2026-07-13 Kevin 报的 bug)。
        // 先退回"这一天"自己之前可能已经算过一次的进货贡献(编辑同一天的进货
        // 数字时),再加这次的新值,避免同一笔进货被重复累加。
        const rawTotal = itemRows[itemRowIdx][6];
        const currentTotal = (rawTotal === "" || rawTotal === undefined || rawTotal === null) ? 0 : Number(rawTotal);
        const newTotal = currentTotal - priorIncomingOnlyBump(existingValues, incFieldId, stockFieldId) + incomingVal;
        itemsSheet.getRange(itemRowIdx + 1, 7).setValue(newTotal);
        itemRows[itemRowIdx][6] = newTotal;
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

    // 这次写的是这个物品的历史日期(不是当前最新一天)——它下一次真实盘点存的
    // usage/cost 可能是基于这次改之前的旧数据算的,补一次向前重算。必须放在
    // 上面这行的 setValues/appendRow 之后:recomputeForwardAfterChange 会重新读一遍
    // 这张表,得先让这次的改动真正落盘,不然读到的还是没改之前的旧值。
    if (needsForwardRecompute) recomputeForwardAfterChange(categoryId, itemId, formatDate(date));
  });
  return { needsConfirm: false };
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

// "这个人今天实际存过的每一条"——跟 getEntryHistory(按物品查"这条记录被谁
// 改过")不一样:这个是按人查"这个人今天动过哪些物品",包含他存的当前值
// (还留在 log 表里)和被后来别人/自己覆盖掉的旧版本(history 表)。用途是
// 每人一个"查看编辑历史"按钮点开后要看到的是"这个人自己今天做了什么",
// 不是"这个物品现在这个值之前长什么样"——一个物品如果后来被别人重新存过,
// 这个人当初存的那条应该照样出现在他自己的记录里,不能因为"现在不是最新的
// 了"就从他的历史里消失。
function getDayHistoryByEditor(categoryId, date, editedBy) {
  const out = [];
  const logRows = getOrCreateSheet(logSheetName(categoryId), LOG_HEADERS).getDataRange().getValues();
  for (let i = 1; i < logRows.length; i++) {
    if (!logRows[i][1]) continue;
    if (!isSameDate(logRows[i][0], date)) continue;
    if (String(logRows[i][7] || "") !== String(editedBy || "")) continue;
    out.push({
      itemId: logRows[i][1], itemName: logRows[i][2],
      values: parseValues(logRows[i][3]),
      usage: logRows[i][4] === "" ? null : Number(logRows[i][4]),
      cost: logRows[i][5] === "" ? null : Number(logRows[i][5]),
      editedAt: logRows[i][6] || null,
      editedBy: logRows[i][7] || null,
      current: true,
    });
  }
  const historyRows = getOrCreateSheet(historySheetName(categoryId), HISTORY_HEADERS).getDataRange().getValues();
  for (let i = 1; i < historyRows.length; i++) {
    if (!historyRows[i][1]) continue;
    if (!isSameDate(historyRows[i][0], date)) continue;
    if (String(historyRows[i][7] || "") !== String(editedBy || "")) continue;
    out.push({
      itemId: historyRows[i][1], itemName: historyRows[i][2],
      values: parseValues(historyRows[i][3]),
      usage: historyRows[i][4] === "" ? null : Number(historyRows[i][4]),
      cost: historyRows[i][5] === "" ? null : Number(historyRows[i][5]),
      editedAt: historyRows[i][6] || null,
      editedBy: historyRows[i][7] || null,
      current: false,
    });
  }
  out.sort(function (a, b) { return new Date(b.editedAt) - new Date(a.editedAt); });
  return out;
}

// 删掉一批日志行之后,要把它们各自当初对 total_stock 的影响"退回去"——而
// 不是从剩下的记录里重新扫一遍最新的"真实盘点"当新总库存(那样会把"进货
// 单独累加""起始库存"这些根本不是来自某一次盘点的贡献一起弄丢:一个物品
// 如果这次被删的是它唯一的一条记录,而那条记录当初是"只填了进货"或者靠
// 起始库存打的底,重新扫描会一条真实盘点都找不到,把总库存直接清空——这就
// 是 2026-07-13 Kevin 报的"填完进货、删掉当天,总库存整个没了")。
// 按物品分组、一次性批量读写 items/log 表,跟 logEntries 里的批量读写模式
// 保持一致,避免删一整天(可能牵涉多个物品)时逐个物品重新读整张表拖慢。
function undoDeletedRowsStockContribution(categoryId, deletedRows) {
  if (deletedRows.length === 0) return;
  const stockFieldId = stockCheckFieldId(categoryId);
  if (!stockFieldId) return;
  const incFieldId = incomingFieldId(categoryId);
  const logRows = getOrCreateSheet(logSheetName(categoryId), LOG_HEADERS).getDataRange().getValues();
  const itemsSheet = getOrCreateSheet(itemsSheetName(categoryId), ITEMS_HEADERS);
  const itemRows = itemsSheet.getDataRange().getValues();
  const itemRowIndexById = {};
  for (let i = 1; i < itemRows.length; i++) {
    if (itemRows[i][0]) itemRowIndexById[String(itemRows[i][0])] = i;
  }

  deletedRows.forEach(function (deletedRow) {
    const itemId = deletedRow[1];
    const itemRowIdx = itemRowIndexById[String(itemId)];
    if (itemRowIdx === undefined) return; // 物品本身也被删了,不用管

    // 被删的这行如果不是这个物品"当时最新"的一条(后面还有更晚日期的记录),
    // 它早就不是"现在总库存"的来源了,删它对现在的总库存没有任何影响。
    const deletedDateStr = formatDate(deletedRow[0]);
    const hasLaterRow = logRows.some(function (r) {
      return String(r[1]) === String(itemId) && formatDate(r[0]) > deletedDateStr;
    });
    if (hasLaterRow) {
      // 对总库存没影响,但如果被删的这行本身填过盘点值或者进货值,它可能是
      // "下一次真实盘点"消耗计算用的基准之一——删掉之后那一行存的 usage/cost
      // 会变成基于一个已经不存在的基准算出来的错误数字,补一次向前重算(跟
      // logEntries 里编辑历史记录时触发的是同一个函数,见 recomputeForwardAfterChange)。
      const delValues = parseValues(deletedRow[3]);
      const delHasStock = !(delValues[stockFieldId] === undefined || delValues[stockFieldId] === null || delValues[stockFieldId] === "");
      const delIncomingRaw = incFieldId ? delValues[incFieldId] : null;
      const delHasIncoming = !(delIncomingRaw === undefined || delIncomingRaw === null || delIncomingRaw === "") && !isNaN(Number(delIncomingRaw));
      if (delHasStock || delHasIncoming) recomputeForwardAfterChange(categoryId, itemId, deletedDateStr);
      return;
    }

    const values = parseValues(deletedRow[3]);
    const stockValRaw = values[stockFieldId];
    const hasStockVal = !(stockValRaw === undefined || stockValRaw === null || stockValRaw === "");
    const incomingRaw = incFieldId ? values[incFieldId] : null;
    const hasIncoming = !(incomingRaw === undefined || incomingRaw === null || incomingRaw === "") && !isNaN(Number(incomingRaw));
    const incomingVal = hasIncoming ? Number(incomingRaw) : 0;
    const usageRaw = deletedRow[4];

    let newTotal;
    if (hasStockVal) {
      // 这行当初是真盘点,把总库存直接推成了 stockVal——用它存的 usage 反推
      // 回当初用的基准(usage = prevStock + incomingVal - stockVal)。当初这
      // 行根本没算出消耗(全新物品的第一次盘点,没有基准可退),删掉后就是
      // "还没盘点过"的状态,退回空,而不是留着一个来源不明的数字。
      newTotal = (usageRaw === "" || usageRaw === undefined || usageRaw === null)
        ? ""
        : Number(usageRaw) + Number(stockValRaw) - incomingVal;
    } else if (hasIncoming && incomingVal !== 0) {
      // 这行当初是只填了进货,把总库存加了 incomingVal——原样减回去。
      const currentTotalRaw = itemRows[itemRowIdx][6];
      const currentTotal = (currentTotalRaw === "" || currentTotalRaw === undefined || currentTotalRaw === null) ? null : Number(currentTotalRaw);
      if (currentTotal === null) return; // 没有可退的基准,保持原样
      newTotal = currentTotal - incomingVal;
    } else {
      return; // 这行既没盘点也没进货,本来就没影响过总库存
    }
    itemsSheet.getRange(itemRowIdx + 1, 7).setValue(newTotal);
    itemRows[itemRowIdx][6] = newTotal; // 保持内存副本同步,防止同一批里同个物品被处理两次时读到旧值
  });
}

function deleteLogDay(categoryId, date) {
  const sheet = getOrCreateSheet(logSheetName(categoryId), LOG_HEADERS);
  const historySheet = getOrCreateSheet(historySheetName(categoryId), HISTORY_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const deletedRows = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] && isSameDate(rows[i][0], date)) {
      historySheet.appendRow(rows[i]); // 删除前存一份快照,配合"历史版本"面板,删了也能找回
      deletedRows.push(rows[i]);
      sheet.deleteRow(i + 1);
    }
  }
  undoDeletedRowsStockContribution(categoryId, deletedRows);
  deleteNoteLogDay(categoryId, date); // 物品记录删了,这天的自定义条目也要一起清掉,不然界面上还残留
}

// 只删某一天里"某个人"录入/最后编辑的那部分记录,别人当天录的不动——
// 用 edited_by 字段区分,不是按"谁创建",是按"最后一次是谁存的"。
function deleteLogDayByEditor(categoryId, date, editedBy) {
  const sheet = getOrCreateSheet(logSheetName(categoryId), LOG_HEADERS);
  const historySheet = getOrCreateSheet(historySheetName(categoryId), HISTORY_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const deletedRows = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] && isSameDate(rows[i][0], date) && String(rows[i][7] || "") === String(editedBy || "")) {
      historySheet.appendRow(rows[i]);
      deletedRows.push(rows[i]);
      sheet.deleteRow(i + 1);
    }
  }
  undoDeletedRowsStockContribution(categoryId, deletedRows);
  deleteNoteLogDayByEditor(categoryId, date, editedBy);
}

// ---------- 自定义条目(2026-07-15 加):跟物品无关、一天一份的自由文本,比如
// "当日营业额"(每天记录的分类)或者"应到货/实际缺"备注(两周/一个月记一次的
// 分类)。故意不复用 fields/items 那一套——那套是"每个物品一行、每个字段一列"
// 的表格结构,这里是"一天一个值",数据形状完全不一样,硬凑会让 logEntries 的
// usage/cost 计算逻辑平白多出一堆"这个字段其实不是物品字段"的特殊判断。 ----------
function listNoteFields(categoryId) {
  const sheet = getOrCreateSheet(notesSheetName(categoryId), NOTE_FIELDS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push({ id: rows[i][0], name: rows[i][1], order: Number(rows[i][2]) });
  }
  out.sort(function (a, b) { return a.order - b.order; });
  return out;
}

function addNoteField(categoryId, name) {
  name = (name || "").trim();
  if (!name) throw new Error("条目名称不能为空。");
  const noteFields = listNoteFields(categoryId);
  const id = "note_" + new Date().getTime();
  const sheet = getOrCreateSheet(notesSheetName(categoryId), NOTE_FIELDS_HEADERS);
  sheet.appendRow([id, name, noteFields.length]);
  return { id: id, name: name };
}

function deleteNoteField(categoryId, noteId) {
  const noteFields = listNoteFields(categoryId).filter(function (f) { return f.id !== noteId; });
  const sheet = getOrCreateSheet(notesSheetName(categoryId), NOTE_FIELDS_HEADERS);
  sheet.clear();
  sheet.appendRow(NOTE_FIELDS_HEADERS);
  sheet.setFrozenRows(1);
  noteFields.forEach(function (f, idx) { sheet.appendRow([f.id, f.name, idx]); });
}

// 按 date+note_id upsert,覆盖前存历史快照——照抄 logEntries 的模式,但不算
// usage/cost(自定义条目纯文本,没有单位/单价,谈不上"消耗")。值是空字符串的
// 条目直接跳过不写,不然每天没填的条目也会堆一堆空行进日志表。
function saveNoteEntries(categoryId, date, values, editedBy) {
  const sheet = getOrCreateSheet(notelogSheetName(categoryId), NOTE_LOG_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const nowIso = new Date().toISOString();
  Object.keys(values || {}).forEach(function (noteId) {
    const value = values[noteId];
    if (value === undefined || value === null || value === "") return;
    let existingRowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(noteId) && isSameDate(rows[i][0], date)) { existingRowIndex = i; break; }
    }
    const rowValues = [date, noteId, value, nowIso, editedBy || ""];
    if (existingRowIndex >= 0) {
      getOrCreateSheet(notehistorySheetName(categoryId), NOTE_HISTORY_HEADERS).appendRow(rows[existingRowIndex]);
      sheet.getRange(existingRowIndex + 1, 1, 1, rowValues.length).setValues([rowValues]);
      rows[existingRowIndex] = rowValues;
    } else {
      sheet.appendRow(rowValues);
      rows.push(rowValues);
    }
  });
}

function getRecentNoteLog(categoryId, days) {
  const sheet = getOrCreateSheet(notelogSheetName(categoryId), NOTE_LOG_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][1]) continue;
    if (new Date(rows[i][0]).getTime() < cutoff.getTime()) continue;
    out.push({
      date: formatDate(rows[i][0]), noteId: rows[i][1], value: rows[i][2],
      editedAt: rows[i][3] || null, editedBy: rows[i][4] || null,
    });
  }
  return out;
}

function deleteNoteLogDay(categoryId, date) {
  const sheet = getOrCreateSheet(notelogSheetName(categoryId), NOTE_LOG_HEADERS);
  const historySheet = getOrCreateSheet(notehistorySheetName(categoryId), NOTE_HISTORY_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] && isSameDate(rows[i][0], date)) {
      historySheet.appendRow(rows[i]);
      sheet.deleteRow(i + 1);
    }
  }
}

function deleteNoteLogDayByEditor(categoryId, date, editedBy) {
  const sheet = getOrCreateSheet(notelogSheetName(categoryId), NOTE_LOG_HEADERS);
  const historySheet = getOrCreateSheet(notehistorySheetName(categoryId), NOTE_HISTORY_HEADERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] && isSameDate(rows[i][0], date) && String(rows[i][4] || "") === String(editedBy || "")) {
      historySheet.appendRow(rows[i]);
      sheet.deleteRow(i + 1);
    }
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

    if (action === "listStores") return jsonResponse({ stores: listStoresForSession(session) });
    if (action === "listCategories") { requireStoreAccess(session, p.storeId); return jsonResponse({ categories: listCategories(p.storeId) }); }
    if (action === "getSettlementCurrency") { requireStoreForCategory(session, p.categoryId); return jsonResponse({ currency: getSettlementCurrency(p.categoryId) }); }
    if (action === "listFields") { requireStoreForCategory(session, p.categoryId); return jsonResponse({ fields: listFields(p.categoryId) }); }
    if (action === "listItems") { requireStoreForCategory(session, p.categoryId); return jsonResponse(getItemsWithLatestStockCheck(p.categoryId)); }
    if (action === "getLog") { requireStoreForCategory(session, p.categoryId); return jsonResponse({ log: getRecentLog(p.categoryId, parseInt(p.days || "30", 10)) }); }
    if (action === "listNoteFields") { requireStoreForCategory(session, p.categoryId); return jsonResponse({ noteFields: listNoteFields(p.categoryId) }); }
    if (action === "getNoteLog") { requireStoreForCategory(session, p.categoryId); return jsonResponse({ noteLog: getRecentNoteLog(p.categoryId, parseInt(p.days || "30", 10)) }); }
    if (action === "getEntryHistory") { requireStoreForCategory(session, p.categoryId); return jsonResponse({ history: getEntryHistory(p.categoryId, p.itemId, p.date) }); }
    if (action === "getDayHistoryByEditor") { requireStoreForCategory(session, p.categoryId); return jsonResponse({ entries: getDayHistoryByEditor(p.categoryId, p.date, p.editedBy) }); }
    if (action === "getExchangeRates") return jsonResponse(getExchangeRates(p.base));
    if (action === "listUsers") { requireLevel(session, ROLE_LEVELS.shift_leader); return jsonResponse({ users: listUsers(session) }); }
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

    // 2026-07-17 加了经理(manager)/班长(shift_leader)两级,权限严格分层
    // (管理员≥经理≥班长≥员工),见 requireLevel 上面的说明。下面每个动作要求
    // 的最低等级都来自 Kevin 核对过的那份权限对照表。
    if (action === "createEmployee") { requireLevel(session, ROLE_LEVELS.shift_leader); return jsonResponse({ user: createEmployee(session, payload.username, payload.password, payload.displayName, payload.storeId, payload.role) }); }
    if (action === "deleteUser") { requireLevel(session, ROLE_LEVELS.manager); deleteUserById(session, payload.userId); return jsonResponse({ status: "ok" }); }
    if (action === "resetPassword") { requireAdmin(session); resetPassword(payload.userId, payload.newPassword); return jsonResponse({ status: "ok" }); }

    if (action === "createStore") { requireAdmin(session); return jsonResponse({ store: createStore(payload.name) }); }
    if (action === "deleteStore") { requireAdmin(session); deleteStoreById(payload.storeId); return jsonResponse({ status: "ok" }); }

    if (action === "createCategory") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreAccess(session, payload.storeId); return jsonResponse({ category: createCategory(payload.storeId, payload.name) }); }
    if (action === "deleteCategory") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); deleteCategoryById(payload.categoryId); return jsonResponse({ status: "ok" }); }
    if (action === "setSettlementCurrency") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); setSettlementCurrency(payload.categoryId, payload.currency); return jsonResponse({ status: "ok" }); }

    if (action === "addItem") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); return jsonResponse({ item: addItem(payload.categoryId, payload.name, payload.unit, payload.price, payload.currency, payload.startingStock) }); }
    if (action === "addItemsBulk") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); return jsonResponse(addItemsBulk(payload.categoryId, payload.items)); }
    if (action === "aiParseImportRows") { requireLevel(session, ROLE_LEVELS.shift_leader); return jsonResponse(aiParseImportRows(payload.rawRows)); }
    if (action === "updateItem") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); updateItem(payload.categoryId, payload.itemId, payload.name, payload.unit, payload.price, payload.currency); return jsonResponse({ status: "ok" }); }
    if (action === "setTotalStock") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); setTotalStock(payload.categoryId, payload.itemId, payload.value); return jsonResponse({ status: "ok" }); }
    if (action === "deleteItem") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); deleteItemById(payload.categoryId, payload.itemId); return jsonResponse({ status: "ok" }); }
    if (action === "deleteItemsBulk") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); return jsonResponse(deleteItemsBulk(payload.categoryId, payload.itemIds)); }
    if (action === "lockItem") { requireLevel(session, ROLE_LEVELS.manager); requireStoreForCategory(session, payload.categoryId); setItemLocked(payload.categoryId, payload.itemId, true); return jsonResponse({ status: "ok" }); }
    if (action === "unlockItem") {
      requireLevel(session, ROLE_LEVELS.manager);
      requireStoreForCategory(session, payload.categoryId);
      verifyCurrentPassword(session, payload.password);
      setItemLocked(payload.categoryId, payload.itemId, false);
      return jsonResponse({ status: "ok" });
    }

    if (action === "addField") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); return jsonResponse({ field: addField(payload.categoryId, payload.name, payload.fieldType, payload.role) }); }
    if (action === "deleteField") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); deleteField(payload.categoryId, payload.fieldId); return jsonResponse({ status: "ok" }); }
    if (action === "reorderFields") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); reorderFields(payload.categoryId, payload.fieldIds); return jsonResponse({ status: "ok" }); }

    if (action === "addNoteField") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); return jsonResponse({ noteField: addNoteField(payload.categoryId, payload.name) }); }
    if (action === "deleteNoteField") { requireLevel(session, ROLE_LEVELS.shift_leader); requireStoreForCategory(session, payload.categoryId); deleteNoteField(payload.categoryId, payload.noteId); return jsonResponse({ status: "ok" }); }
    if (action === "saveNoteEntries") {
      requireStoreForCategory(session, payload.categoryId);
      saveNoteEntries(payload.categoryId, payload.date, payload.values, session.display_name || session.username);
      return jsonResponse({ status: "ok" });
    }

    if (action === "saveLogEntries") {
      requireStoreForCategory(session, payload.categoryId);
      const result = logEntries(payload.categoryId, payload.date, payload.entries, session.display_name || session.username, !!payload.confirmOverwrite);
      if (result && result.needsConfirm) return jsonResponse({ needsConfirm: true, conflicts: result.conflicts });
      return jsonResponse({ status: "ok" });
    }

    if (action === "deleteLogDay") {
      requireLevel(session, ROLE_LEVELS.manager);
      requireStoreForCategory(session, payload.categoryId);
      deleteLogDay(payload.categoryId, payload.date);
      return jsonResponse({ status: "ok" });
    }

    if (action === "deleteLogDayByEditor") {
      // 员工只能删自己录入/最后编辑的那部分;班长(level 2)开始就能删这一天
      // 任何人的记录了,不用再是自己录的——服务器端强制核对角色等级,不能只靠
      // 前端把按钮藏起来防止越权删除。
      requireStoreForCategory(session, payload.categoryId);
      if ((ROLE_LEVELS[session.role] || 0) < ROLE_LEVELS.shift_leader && String(payload.editedBy) !== String(session.display_name || session.username)) {
        throw new Error("只能删除自己录入的记录。");
      }
      deleteLogDayByEditor(payload.categoryId, payload.date, payload.editedBy);
      return jsonResponse({ status: "ok" });
    }

    return jsonResponse({ error: "unknown action" });
  } catch (err) {
    return jsonResponse({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}
