/**
 * 車検・整備リクエストシステム - 【真・最終完全版：工場入庫日ロジック＆全通知統合】
 */

// 1. 設定エリア
const ADMIN_EMAILS = ["honbu-admin@example.com", "honbu-sub@example.com", "office@example.com"];
const STORE_EMAILS = {
  "車検課": "syaken@example.com", "板橋ss": "it@ex.com", "志村ss": "shi@ex.com",
  "小茂根ss": "ko@ex.com", "赤羽西ss": "aka@ex.com", "東坂下ss": "hi@ex.com",
  "高島平ss": "taka@ex.com", "武蔵関ss": "mu@ex.com"
};

const FOLDER_ID = "120Pjmdr36BTFFTQxtgrJN6GGY9ovobrQ"; 

function doGet(e) {
  const page = e.parameter.p || 'index';
  const template = HtmlService.createTemplateFromFile(page);
  template.scriptUrl = ScriptApp.getService().getUrl(); 
  return template.evaluate()
      .setTitle('車検・整備リクエストシステム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getHolidays() {
  try {
    const cal = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com");
    return cal.getEvents(new Date(2026,0,1), new Date(2027,11,31)).map(e => Utilities.formatDate(e.getStartTime(), "JST", "yyyy-MM-dd"));
  } catch(e) { return []; }
}

function parseFixedDate(str) {
  if (!str) return null;
  const p = String(str).split('-');
  return p.length === 3 ? new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0) : null;
}

function calculateEntryDate(baseDate, courseName) {
  let eDate = new Date(baseDate.getTime());
  if (courseName && courseName.indexOf('セーフティー') !== -1) {
    eDate.setDate(eDate.getDate() - 1); 
    if (eDate.getDay() === 0) eDate.setDate(eDate.getDate() - 2); 
    if (eDate.getDay() === 6) eDate.setDate(eDate.getDate() - 1); 
  }
  return eDate;
}

function sendSafeEmail(to, subject, body) {
  if (!to || to.trim() === "") return;
  try { GmailApp.sendEmail(to, subject, body); } catch(e) { console.error("Mail fail: " + to); }
}

function submitRequest(formData) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) return { success: false, message: "混雑中" };
    const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約データ');
    const date1 = parseFixedDate(formData.date1);
    const entryDate = calculateEntryDate(date1, formData.course);

    let fileUrls = ["", "", "", ""]; 
    try {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      ['file1', 'file2', 'file3', 'file4'].forEach((key, index) => {
        if (formData[key] && formData[key].data) {
          const blob = Utilities.newBlob(Utilities.base64Decode(formData[key].data), formData[key].type, formData[key].name);
          fileUrls[index] = folder.createFile(blob).getUrl();
        }
      });
    } catch(e) {}

    const requestId = "REQ-" + Utilities.formatDate(new Date(), "JST", "MMdd-HHmm") + "-" + Math.floor(Math.random()*100);
    ss.appendRow([requestId, formData.shop, formData.name, formData.car, formData.num, formData.course, formData.date1, entryDate, "リクエスト中", formData.note, new Date(), "", fileUrls.join("\n"), formData.date2 || ""]);
    
    const commonBody = `店舗：${formData.shop}\n顧客：${formData.name}様\n車種：${formData.car}\n通検希望日：${formData.date1}\nコース：${formData.course}`;
    ADMIN_EMAILS.forEach(email => sendSafeEmail(email, `【新規】${formData.shop}よりリクエスト`, commonBody));
    sendSafeEmail(STORE_EMAILS[formData.shop], `【リクエスト控え】送信完了しました`, commonBody);

    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function finalizeRequest(id, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約データ');
  const v = ss.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(id)) {
      const r = i + 1;
      const fDate = parseFixedDate(dateStr);
      const eDate = calculateEntryDate(fDate, String(v[i][5]));
      ss.getRange(r, 9).setValue('確定'); 
      ss.getRange(r, 7).setValue(fDate);
      ss.getRange(r, 8).setValue(eDate);
      const body = `${v[i][2]}様の内容が確定しました。\n\n通検予定日：${dateStr}\n工場入庫日：${Utilities.formatDate(eDate, "JST", "yyyy-MM-dd")}`;
      sendSafeEmail(STORE_EMAILS[v[i][1]], `【確定通知】予約が確定しました`, body);
      return "完了";
    }
  }
}

function updateRequestData(obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約データ');
  const v = ss.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(obj.id)) {
      const r = i + 1;
      const nDate = parseFixedDate(obj.date);
      const eDate = calculateEntryDate(nDate, obj.course);
      ss.getRange(r, 2, 1, 5).setValues([[obj.shop, obj.name, obj.car, obj.num, obj.course]]);
      ss.getRange(r, 7).setValue(nDate); ss.getRange(r, 8).setValue(eDate); ss.getRange(r, 10).setValue(obj.note);
      const body = `予約内容が修正されました。\n\n顧客：${obj.name}様\n通検予定日：${obj.date}\n工場入庫日：${Utilities.formatDate(eDate, "JST", "yyyy-MM-dd")}`;
      sendSafeEmail(STORE_EMAILS[obj.shop], `【重要】予約内容が変更されました`, body);
      return "完了";
    }
  }
}

function deleteRequest(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約データ');
  const v = ss.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(id)) {
      const shop = v[i][1]; const name = v[i][2];
      ss.deleteRow(i + 1);
      sendSafeEmail(STORE_EMAILS[shop], `【削除通知】リクエストが取り消されました`, `${name}様のリクエストは削除されました。`);
      return "完了";
    }
  }
}

function getAdminData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約データ');
  const v = ss.getDataRange().getValues();
  const d = ss.getDataRange().getDisplayValues();
  let inspectList = [], factoryList = [], pendingList = [];
  const days = ['日','月','火','水','木','金','土'];
  const fmt = (date) => (date instanceof Date && !isNaN(date.getTime())) ? (date.getMonth()+1)+"/"+date.getDate()+"("+days[date.getDay()]+")" : "-";
  for (let i = 1; i < v.length; i++) {
    const row = v[i]; if (!row[8]) continue;
    const obj = { id: row[0], shop: row[1], name: row[2], car: row[3], num: row[4], course: String(row[5]), note: row[9], dateStr: fmt(row[6]), entryStr: fmt(row[7]), rawDate1: d[i][6], rawDate2: d[i][13], date1Str: d[i][6], date2Str: d[i][13], fileUrls: row[12] };
    if (String(row[8]) !== "確定") { pendingList.push(obj); continue; }
    if (obj.course.indexOf("車検") !== -1) inspectList.push(obj);
    else factoryList.push(obj);
  }
  return { inspect: inspectList, factory: factoryList, pending: pendingList };
}

function getShopStatusData(shopName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約データ');
  const v = ss.getDataRange().getValues();
  const d = ss.getDataRange().getDisplayValues();
  let pending = [], confirmed = [];
  const fmt = (date) => (date instanceof Date && !isNaN(date.getTime())) ? (date.getMonth()+1)+"/"+date.getDate() : "-";
  for (let i = 1; i < v.length; i++) {
    if (v[i][1] !== shopName) continue;
    const obj = { name: v[i][2], car: v[i][3], num: v[i][4], course: v[i][5], note: v[i][9], date1: d[i][6], date2: d[i][13], confirmedDate: fmt(v[i][6]), entryDate: fmt(v[i][7]), status: v[i][8] };
    if (v[i][8] === "確定") confirmed.push(obj); else pending.push(obj);
  }
  return { pending, confirmed };
}

function checkAdminPassword(pw) { return pw === "1234"; }