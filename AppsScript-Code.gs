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
    return jsonOutput_(readDatabase_());
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
