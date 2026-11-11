/* 寶寶性別大猜猜：Google Apps Script 後端
 * 1. 建立一份 Google 試算表，選單「擴充功能 → Apps Script」，把本檔內容貼上並儲存。
 * 2. 專案設定 → 指令碼屬性 → 新增 ADMIN_PASSWORD = 管理者密碼（密碼只存在這裡，不要寫進程式碼）。
 * 3. 部署 → 新增部署作業 → 類型「網頁應用程式」，執行身分「我」，誰可以存取「所有人」，複製網址。
 * 4. 把網址填入 index.html 的 API_URL。
 * 資料存放在試算表第一個工作表的 A1 儲存格，請勿手動修改。
 */

const DATA_CELL = "A1";
const MAX_CELL_CHARS = 49000;
const MAX_BETS = 300;
const MAX_NAME_LEN = 12;
const MAX_AMOUNT = 1000000;
const MAX_FAILS = 5;
const LOCK_SECONDS = 60;
const TOKEN_SECONDS = 6 * 60 * 60;
const DEFAULT_SETTINGS = { maxBets: 3, open: true, result: null };

function doGet() {
  try {
    return reply_({ ok: true, data: readData_() });
  } catch (err) {
    return reply_(errorReply_(err));
  }
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    switch (req.action) {
      case "login": return reply_(login_(req.password));
      case "check": requireAdmin_(req.token); return reply_({ ok: true });
      case "logout": logout_(req.token); return reply_({ ok: true });
      case "save": return reply_(save_(req.token, req.data));
      default: throw fail_("BAD_REQUEST", "未知操作");
    }
  } catch (err) {
    return reply_(errorReply_(err));
  }
}

function login_(password) {
  const cache = CacheService.getScriptCache();
  if (cache.get("lock")) throw fail_("LOCKED", `嘗試次數過多，請 ${LOCK_SECONDS} 秒後再試`);
  const expected = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD");
  if (!expected) throw fail_("NOT_CONFIGURED", "伺服器尚未設定 ADMIN_PASSWORD");
  if (typeof password !== "string" || !safeEqual_(password, expected)) {
    const fails = Number(cache.get("fails") || 0) + 1;
    if (fails >= MAX_FAILS) {
      cache.remove("fails");
      cache.put("lock", "1", LOCK_SECONDS);
      throw fail_("LOCKED", `失敗次數過多，請 ${LOCK_SECONDS} 秒後再試`);
    }
    cache.put("fails", String(fails), LOCK_SECONDS);
    throw fail_("BAD_PASSWORD", "密碼錯誤");
  }
  cache.remove("fails");
  const token = Utilities.getUuid();
  cache.put("tok_" + token, "1", TOKEN_SECONDS);
  return { ok: true, token };
}

function logout_(token) {
  if (typeof token === "string") CacheService.getScriptCache().remove("tok_" + token);
}

function save_(token, data) {
  requireAdmin_(token);
  const clean = sanitize_(data);
  const json = JSON.stringify(clean);
  if (json.length > MAX_CELL_CHARS) throw fail_("TOO_LARGE", "資料過大，請減少下注筆數");
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw fail_("BUSY", "伺服器忙碌，請稍後再試");
  try {
    sheet_().getRange(DATA_CELL).setValue(json);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, data: clean };
}

function requireAdmin_(token) {
  if (typeof token !== "string" || !CacheService.getScriptCache().get("tok_" + token)) {
    throw fail_("AUTH", "登入已過期，請重新登入");
  }
}

function readData_() {
  const raw = sheet_().getRange(DATA_CELL).getValue();
  if (!raw) return { settings: Object.assign({}, DEFAULT_SETTINGS), bets: [] };
  return sanitize_(JSON.parse(raw));
}

// 伺服器端的最終驗證，前端的檢查只是為了使用體驗
function sanitize_(raw) {
  if (!raw || typeof raw !== "object" || !Array.isArray(raw.bets)) throw fail_("BAD_DATA", "格式錯誤：缺少 bets 陣列");
  if (raw.bets.length > MAX_BETS) throw fail_("BAD_DATA", `下注筆數超過上限 ${MAX_BETS} 筆`);
  const s = raw.settings && typeof raw.settings === "object" ? raw.settings : {};
  const settings = Object.assign({}, DEFAULT_SETTINGS);
  if (Number.isInteger(s.maxBets) && s.maxBets >= 1 && s.maxBets <= 50) settings.maxBets = s.maxBets;
  if (typeof s.open === "boolean") settings.open = s.open;
  if (s.result === "boy" || s.result === "girl") settings.result = s.result;
  const bets = raw.bets.map((b, i) => {
    const name = b && typeof b.name === "string" ? b.name.trim() : "";
    const ok = name.length >= 1 && name.length <= MAX_NAME_LEN
      && (b.side === "boy" || b.side === "girl")
      && Number.isInteger(b.amount) && b.amount > 0 && b.amount <= MAX_AMOUNT
      && typeof b.createdAt === "number" && isFinite(b.createdAt);
    if (!ok) throw fail_("BAD_DATA", `第 ${i + 1} 筆下注格式錯誤`);
    const id = typeof b.id === "string" && /^[\w-]{1,40}$/.test(b.id) ? b.id : Utilities.getUuid();
    return { id, name, side: b.side, amount: b.amount, createdAt: b.createdAt };
  });
  return { settings, bets };
}

function safeEqual_(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function fail_(code, message) {
  const err = new Error(message);
  err.code = code;
  return err;
}

function errorReply_(err) {
  return err.code
    ? { ok: false, code: err.code, error: err.message }
    : { ok: false, code: "SERVER", error: "伺服器錯誤" };
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
