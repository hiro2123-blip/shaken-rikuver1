/**
 * 車検・整備リクエストシステム - 【真・最終完全版：3拠点メール通知＆全機能統合】
 */

// 1. 設定エリア
// ★ここにメールアドレスを3つ入力してください（例: "admin1@ex.com", "admin2@ex.com"...）
const ADMIN_EMAILS = ["", "", ""]; 

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

// セーフティー車検入庫日計算（前日。土日なら金曜）
function calculateEntryDate(baseDate, courseName) {
  if (!baseDate) return null;
  let eDate = new Date(baseDate.getTime());
  if (courseName && courseName.indexOf('セーフティー') !== -1) {
    eDate.setDate(eDate.getDate() - 1); 
    if (eDate.getDay() === 0) eDate.setDate(eDate.getDate() - 2); 
    if (eDate.getDay() === 6) eDate.setDate(eDate.getDate() - 1); 
  }
  return eDate;
}

function submitRequest(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約データ');
    const date1 = parseFixedDate(formData.date1);
    const entryDate = calculateEntryDate(date1, formData.course);
    
    // ファイルアップロード処理
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
    
    // 管理者3名へのメール通知
    const validEmails = ADMIN_EMAILS.filter(email => email !== "");
    if (validEmails.length > 0) {
      MailApp.sendEmail({
        to: validEmails.join(","),
        subject: "【新規リクエスト】" + formData.shop + " / " + formData.name + "様",
        body: "新規リクエストが届きました。\n\n店舗: " + formData.shop + "\n顧客: " + formData.name + "様\n車種: " + formData.car + "\n希望日: " + formData.date1 + "\nコース: " + formData.course + "\n備考: " + formData.note
      });
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function finalizeRequest(id, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約データ');
  const v = ss.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(id)) {
      const fDate = parseFixedDate(dateStr);
      const eDate = calculateEntryDate(fDate, String(v[i][5]));
      ss.getRange(i+1, 9).setValue('確定'); 
      ss.getRange(i+1, 7).setValue(fDate); ss.getRange(i+1, 8).setValue(eDate);
      return "完了";
    }
  }
}

function updateRequestData(obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約データ');
  const v = ss.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(obj.id)) {
      const nDate = parseFixedDate(obj.date);
      const eDate = calculateEntryDate(nDate, obj.course);
      ss.getRange(i+1, 2, 1, 5).setValues([[obj.shop, obj.name, obj.car, obj.num, obj.course]]);
      ss.getRange(i+1, 7).setValue(nDate); ss.getRange(i+1, 8).setValue(eDate); ss.getRange(i+1, 10).setValue(obj.note);
      return "完了";
    }
  }
}

function deleteRequest(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約データ');
  const v = ss.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(id)) { ss.deleteRow(i + 1); return "完了"; }
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
    const row = v[i]; if (!row[0]) continue;
    const obj = { 
      id: row[0], shop: row[1], name: row[2], car: row[3], num: row[4], course: String(row[5]), note: row[9], status: row[8], 
      dateStr: fmt(row[6]), entryStr: fmt(row[7]), rawDate1: d[i][6], rawDate2: d[i][13], 
      inspectTime: (row[6] instanceof Date) ? row[6].getTime() : 0,
      entryTime: (row[7] instanceof Date) ? row[7].getTime() : 0,
      date1Str: d[i][6], date2Str: d[i][13], fileUrls: row[12] 
    };
    if (obj.status !== "確定") { pendingList.push(obj); continue; }
    if (obj.course.indexOf("車検") !== -1) inspectList.push(obj);
    if (obj.course.indexOf("セーフティー") !== -1 || obj.course.indexOf("一般") !== -1) factoryList.push(obj);
  }
  inspectList.sort((a, b) => a.inspectTime - b.inspectTime);
  factoryList.sort((a, b) => a.entryTime - b.entryTime);
  return { pending: pendingList, inspect: inspectList, factory: factoryList };
}

function getShopStatusData(shopName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約データ');
  const v = ss.getDataRange().getValues();
  const d = ss.getDataRange().getDisplayValues();
  let pending = [], confirmed = [];
  const fmt = (date) => (date instanceof Date && !isNaN(date.getTime())) ? (date.getMonth()+1)+"/"+date.getDate() : "-";
  for (let i = 1; i < v.length; i++) {
    if (v[i][1] !== shopName) continue;
    const obj = { name: v[i][2], car: v[i][3], num: v[i][4], course: v[i][5], note: v[i][9], confirmedDate: fmt(v[i][6]), entryDate: fmt(v[i][7]), status: v[i][8], date1: d[i][6], date2: d[i][13] };
    if (v[i][8] === "確定") confirmed.push(obj); else pending.push(obj);
  }
  return { pending, confirmed };
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

function checkAdminPassword(pw) { return pw === "1234"; }