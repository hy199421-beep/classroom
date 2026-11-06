/**
 * 學習島共用積分資料 API
 * 試算表：1SWdGYSpAKRoA8q8YpLPYqcpbDaRDBPFiR7piEjeOKo0
 * 工作表：Records
 * JSON 資料存放：A1
 *
 * 支援：
 * 1. GET：讀取完整資料
 * 2. POST action=saveGroup：只更新指定日期、指定班級，不覆蓋其他日期/班級/商店紀錄
 * 3. POST action=redeem：以 id 去重後新增商店兌換紀錄
 */
const SPREADSHEET_ID = '1SWdGYSpAKRoA8q8YpLPYqcpbDaRDBPFiR7piEjeOKo0';
const SHEET_NAME = 'Records';

/**
 * v3 增補：教師專用課堂加分 API
 * 每次固定 +1，eventId 去重；保留舊 bonusScore、簽到及商店資料。
 */
const ALLOWED_GROUPS_V3 = {
  "數學9A": ["慶豐","詠甯","祺叡","冠閔","鴻昱","聖民","柏曜"],
  "國文8A": ["淮喆","鼎杰","佳瑋","子田","司浩"],
  "國文8B": ["昱佑","奕訢","日良","羽琪","世賢","宣澔"]
};

function addClassPoint_(db, payload) {
  const date = String(payload.date || "");
  const group = String(payload.group || "");
  const name = String(payload.name || "");
  const sessionId = String(payload.sessionId || "");
  const eventId = String(payload.eventId || "");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error("日期格式錯誤");
  if (!ALLOWED_GROUPS_V3[group] || !ALLOWED_GROUPS_V3[group].includes(name)) throw new Error("班級或學生姓名不在授權名單中");
  if (!/^.{1,40}$/.test(sessionId)) throw new Error("課程節次資料錯誤");
  if (!/^[A-Za-z0-9_-]{8,100}$/.test(eventId)) throw new Error("加分識別碼格式錯誤");
  if (!db[date] || typeof db[date] !== "object" || Array.isArray(db[date])) db[date] = {};
  if (!Array.isArray(db[date][group])) db[date][group] = [];
  let record = db[date][group].find(item => item && item.name === name);
  if (!record) {
    record = { id: name + "_" + Date.now(), name, bonusScore: 0, attendance: "未標記", isDuty: false, warnCount: 0 };
    db[date][group].push(record);
  }
  if (!Array.isArray(record.classPointEvents)) record.classPointEvents = [];
  const duplicate = record.classPointEvents.some(item => item && item.id === eventId);
  if (duplicate) return { status: "success", action: "addClassPoint", duplicate: true, eventId, date, group, name };
  record.bonusScore = (Number(record.bonusScore) || 0) + 1;
  record.classPointEvents.push({
    id: eventId, type: "課堂加分", points: 1, date, group, name, sessionId,
    createdAt: new Date().toISOString()
  });
  getRecordsSheet_().getRange("A1").setValue(JSON.stringify(db));
  return { status: "success", action: "addClassPoint", duplicate: false, eventId, date, group, name, points: 1 };
}


function getRecordsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('找不到工作表：' + SHEET_NAME);
  return sheet;
}

function readDatabase_() {
  const raw = getRecordsSheet_().getRange('A1').getValue();
  if (!raw) return {};
  const data = JSON.parse(String(raw));
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw new Error('Records!A1 的 JSON 格式不正確');
  }
  return data;
}

function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  try {
    const db = readDatabase_();
    // 回傳版本標記供前端確認 API 已升級；此欄位不寫入試算表。
    db.__apiVersion = 2;
    return jsonOutput_(db);
  } catch (err) {
    return jsonOutput_({ status: 'error', message: String(err && err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('沒有收到 POST 資料');
    }

    const payload = JSON.parse(e.postData.contents);
    const db = readDatabase_();

    if (payload.action === "addClassPoint") {
      const result = addClassPoint_(db, payload);
      return jsonOutput_(result);
    }

    if (payload.action === 'saveGroup') {
      const date = String(payload.date || '');
      const group = String(payload.group || '');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('日期格式錯誤');
      if (!group) throw new Error('缺少班級名稱');
      if (!Array.isArray(payload.records)) throw new Error('records 必須是陣列');

      if (!db[date] || typeof db[date] !== 'object' || Array.isArray(db[date])) {
        db[date] = {};
      }
      // 只覆寫指定日期、指定班級；其餘歷史日期、其他班級及 __shopRedemptions 原樣保留。
      db[date][group] = payload.records;
      getRecordsSheet_().getRange('A1').setValue(JSON.stringify(db));
      return jsonOutput_({ status: 'success', action: 'saveGroup', date: date, group: group });
    }

    if (payload.action === 'redeem') {
      const redemption = payload.redemption;
      if (!redemption || !redemption.id || !redemption.group || !redemption.name ||
          !redemption.reward || !Number.isFinite(Number(redemption.cost))) {
        throw new Error('兌換資料不完整');
      }
      if (!Array.isArray(db.__shopRedemptions)) db.__shopRedemptions = [];
      const exists = db.__shopRedemptions.some(item => item && item.id === redemption.id);
      if (!exists) db.__shopRedemptions.push(redemption);
      getRecordsSheet_().getRange('A1').setValue(JSON.stringify(db));
      return jsonOutput_({
        status: 'success',
        action: 'redeem',
        redemptionId: redemption.id,
        duplicate: exists
      });
    }

    throw new Error('不支援的操作：' + String(payload.action || '(未提供 action)'));
  } catch (err) {
    return jsonOutput_({ status: 'error', message: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}
