// ==================== 設定 ====================
var VERSION = "v3.0 (2026/06/28)";
var MASTER_SS_ID   = '1L5aFDXAmfUDkBg8d7X3WqJgMhdMq5tM5sfUZ2G-M58E';
var DATA_SS_ID     = '1pPID9hCnSHAzs155KQYRW7ZymgeA1fSyHmnkRqVwFGE';
var SHINRYO_SS_ID  = '1vlCqsHYq837RyKfzTDk6gIwhChJvr1wxS0jDpN7xrJk';
var OTEMACHI_SS_ID = '1LMbLBRwRXDt7hi8_mouk2GNs7425RIXsNv_diTfCAbg';
var KASE_EMAIL     = 'mintcocoajasmine@gmail.com';
var ONO_EMAIL      = 'chloeandnina1@gmail.com';
var ADMIN_URL    = 'https://script.google.com/macros/s/AKfycbys7A1hwDTvpJms24uUJQLiD-oN8trVFMBfUEaU-99RePefcemkkA754rxBW0cnnY6o/exec?page=admin';
var TIME_SLOTS     = ['13:00-14:15','14:20-15:35','15:45-17:00','17:10-18:25','18:35-19:50','20:00-21:15'];
var SLOT_COLS      = [5, 9, 13, 17, 21, 25];
var OBON_DATES     = ['2026/08/10','2026/08/11','2026/08/12','2026/08/13','2026/08/14','2026/08/15','2026/08/16'];

// ==================== ルーティング ====================
function doGet(e) {
  ensureSheets();
  var params = e.parameter || {};
  var page = params.page;
  var action = params.action;

  // GitHub Pages用API（JSONP/JSON形式）
  if (action) {
    var result = {};
    try {
      if (action === 'login') {
        result = login(params.code);
      } else if (action === 'getSlots') {
        result = getSlots(params.school);
      } else if (action === 'getMyAvailability') {
        result = getMyAvailability(params.code, params.school);
      } else if (action === 'saveAvailability') {
        var entries = JSON.parse(params.entries || '[]');
        result = saveAvailability(params.code, params.school, entries, params.doSendEmail !== 'false');
      } else if (action === 'checkDeadline') {
        result = checkDeadline(params.school);
      } else {
        result = { ok: false, error: 'unknown action' };
      }
    } catch(ex) {
      result = { ok: false, error: ex.message };
    }
    var json = JSON.stringify(result);
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (page === 'admin') {
    return HtmlService.createHtmlOutput(adminHtmlStr())
      .setTitle('管理者画面 - ステップ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1');
  }
  return HtmlService.createHtmlOutput(teacherHtml())
    .setTitle('出勤登録 - 個別指導ステップ')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}

// ==================== 認証 ====================
function login(code) {
  var sheet = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('講師マスター');
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(code).trim()) {
      return { ok: true, name: String(data[i][1]), email: String(data[i][15] || '') };
    }
  }
  return { ok: false, error: 'コード番号が見つかりません' };
}

// ==================== 締め切り取得 ====================
function getDeadline(school) {
  var ss    = SpreadsheetApp.openById(DATA_SS_ID);
  var sheet = ss.getSheetByName('締め切り設定');
  if (!sheet) return { ok: true, deadline: null };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === school) {
      var d = data[i][1];
      if (!d) return { ok: true, deadline: null };
      var dt = d instanceof Date ? d : new Date(d);
      var raw = Utilities.formatDate(dt, 'Asia/Tokyo', 'yyyy/MM/dd');
      return { ok: true, deadline: raw };
    }
  }
  return { ok: true, deadline: null };
}

// ==================== 締め切り確認（先生用） ====================
function checkDeadline(school) {
  var r = getDeadline(school);
  if (!r.deadline) return { ok: true, locked: false, deadline: null, daysLeft: null };
  var today = new Date();
  var todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
  var dl = new Date(r.deadline);
  var todayMidnight = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var dlMidnight    = new Date(dl.getFullYear(), dl.getMonth(), dl.getDate());
  var diff = Math.floor((dlMidnight - todayMidnight) / 86400000);
  var locked = diff < 0;
  return { ok: true, locked: locked, deadline: r.deadline, daysLeft: locked ? 0 : diff };
}

// ==================== 授業枠取得 ====================
function getSlots(school) {
  var sheetName = (school === '大手町') ? '授業枠設定_大手町' : '授業枠設定';
  var ss = SpreadsheetApp.openById(DATA_SS_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.getSheetByName('授業枠設定');
  var data  = sheet.getDataRange().getValues();

  // まず全行を日付順にソート
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var d = data[i][0] instanceof Date ? data[i][0] : new Date(data[i][0]);
    var raw = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
    rows.push({ raw: raw, d: d, cols: data[i] });
  }
  rows.sort(function(a, b) { return a.raw < b.raw ? -1 : 1; });

  var slots = [];
  var obonAdded = false;
  for (var i = 0; i < rows.length; i++) {
    var raw = rows[i].raw;
    var d   = rows[i].d;
    if (OBON_DATES.indexOf(raw) !== -1) {
      if (!obonAdded) {
        slots.push({ date: 'お盆休み（8/10〜8/16）', rawDate: 'HOLIDAY', enabled: [false,false,false,false,false,false], holiday: true });
        obonAdded = true;
      }
      continue;
    }
    var label = Utilities.formatDate(d,'Asia/Tokyo','M/d') + '（' + ['日','月','火','水','木','金','土'][d.getDay()] + '）';
    var enabled = [0,1,2,3,4,5].map(function(j) {
      var v = rows[i].cols[j+1];
      return v === true || v === 'TRUE' || Number(v) === 1;
    });
    slots.push({ date: label, rawDate: raw, enabled: enabled });
  }
  return { ok: true, slots: slots };
}

// ==================== 入力済みデータ取得 ====================
function getMyAvailability(code, school) {
  var sheetName = (school === '大手町') ? '出勤可否_大手町' : '出勤可否';
  var sheet = SpreadsheetApp.openById(DATA_SS_ID).getSheetByName(sheetName);
  if (!sheet) return { ok: true, entries: {} };
  var data = sheet.getDataRange().getValues();
  var map  = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== String(code).trim()) continue;
    var v = String(data[i][4]);
    if (v === '') continue;
    var dv2 = data[i][2];
    var dk2 = (dv2 instanceof Date) ? Utilities.formatDate(dv2, 'Asia/Tokyo', 'yyyy/MM/dd') : String(dv2);
    map[dk2 + '|' + String(data[i][3])] = v;
  }
  return { ok: true, entries: map };
}

// ==================== 出勤可否保存 ====================
function saveAvailability(code, school, entries, doSendEmail) {
  // 締め切りチェック
  var dl = checkDeadline(school);
  if (dl.locked) return { ok: false, error: '締め切りを過ぎているため変更できません。' };

  var teacher = login(code);
  if (!teacher.ok) return { ok: false, error: '先生情報が取得できません' };

  var sheetName = (school === '大手町') ? '出勤可否_大手町' : '出勤可否';
  var ss    = SpreadsheetApp.openById(DATA_SS_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['コード','氏名','日付','枠インデックス','可否','更新日時','変更フラグ']);
  }
  var now = new Date();

  // この講師の既存行を全削除
  var data = sheet.getDataRange().getValues();
  var delRows = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(code).trim()) delRows.push(i + 1);
  }
  for (var j = delRows.length - 1; j >= 0; j--) sheet.deleteRow(delRows[j]);

  // 新データを全追記
  var newRows = [];
  entries.forEach(function(entry) {
    if (entry.rawDate === 'HOLIDAY' || !entry.value) return;
    newRows.push([String(code), teacher.name, entry.rawDate, Number(entry.slotIndex), entry.value, now, '']);
  });
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
  }

  var adminEmail = (school === '大手町') ? ONO_EMAIL : KASE_EMAIL;

  if (doSendEmail !== false && newRows.length > 0) {
    var lines = newRows.map(function(r) {
      return (r[4] === '△' ? '△' : '○') + ' ' + r[2] + ' ' + TIME_SLOTS[r[3]];
    }).join('\n');
    var subject = '【出勤登録】' + school + '校 ' + teacher.name + 'さん';
    var bodyT = teacher.name + 'さん、以下の内容で登録しました。\n\n' + lines + '\n\n個別指導ステップ';
    var bodyA = school + '校 ' + teacher.name + 'さんが登録しました。\n\n' + lines + '\n\n▼管理者画面\n' + ADMIN_URL;
    if (teacher.email) { try { MailApp.sendEmail(teacher.email, subject, bodyT); } catch(ex){} }
    try { MailApp.sendEmail(adminEmail, subject, bodyA); } catch(ex){}
  }

  return { ok: true };
}

// ==================== 全体時間割 一括書き込み（管理者用） ====================
function bulkWriteToSchedule(school) {
  var schedSsId  = (school === '大手町') ? OTEMACHI_SS_ID : SHINRYO_SS_ID;
  var availSheet = (school === '大手町') ? '出勤可否_大手町' : '出勤可否';
  var sheet      = SpreadsheetApp.openById(schedSsId).getSheetByName('全体時間割');
  var avail      = SpreadsheetApp.openById(DATA_SS_ID).getSheetByName(availSheet);
  if (!sheet || !avail) return { ok: false, error: 'シートが見つかりません' };

  var rows = avail.getDataRange().getValues();

  // 先生ごとに登録データをまとめる
  // key: date+'|'+slotIndex → [{name, value}]
  var slotMap = {};
  for (var i = 1; i < rows.length; i++) {
    var val = String(rows[i][4]);
    if (val === '') continue;
    var key = rows[i][2] + '|' + rows[i][3];
    if (!slotMap[key]) slotMap[key] = [];
    slotMap[key].push({ name: String(rows[i][1]), value: val });
  }

  var schedData = sheet.getDataRange().getValues();

  // 書き込み処理
  Object.keys(slotMap).forEach(function(key) {
    var parts    = key.split('|');
    var rawDate  = parts[0];
    var slotIdx  = parseInt(parts[1]);
    var col      = SLOT_COLS[slotIdx];
    var teachers = slotMap[key]; // [{name, value}]
    var em = parseInt(rawDate.split('/')[1]);
    var ed = parseInt(rawDate.split('/')[2]);

    // 対象日の行を探す
    var matchRows = [];
    for (var r = 1; r < schedData.length; r++) {
      var cell = schedData[r][0];
      var matched = false;
      if (cell instanceof Date && !isNaN(cell)) {
        matched = Utilities.formatDate(cell, 'Asia/Tokyo', 'yyyy/MM/dd') === rawDate;
      } else {
        var s = String(cell);
        matched = s.indexOf(em + '/' + ed + '（') !== -1 || s.indexOf(em + '/' + ed) === 0;
      }
      if (matched) matchRows.push(r + 1); // 1-indexed
    }
    if (matchRows.length === 0) return;

    // 既存のその列の空きを調べ、足りなければ行を追加
    var writeRow = 0;
    teachers.forEach(function(t) {
      var displayName = (t.value === '△') ? '△' + t.name : t.name;
      // 空き行を探す（対象日の行の中で）
      var placed = false;
      for (var ri = 0; ri < matchRows.length; ri++) {
        var cur = String(sheet.getRange(matchRows[ri], col).getValue());
        if (cur === '' || cur === displayName) {
          sheet.getRange(matchRows[ri], col).setValue(displayName);
          placed = true;
          break;
        }
      }
      if (!placed) {
        // 行を追加
        var last = matchRows[matchRows.length - 1];
        sheet.insertRowAfter(last);
        // 追加した行にA列・B列をコピー
        sheet.getRange(last + 1, 1).setValue(sheet.getRange(matchRows[0], 1).getValue());
        var dow = sheet.getRange(matchRows[0], 2).getValue();
        if (dow) sheet.getRange(last + 1, 2).setValue(dow);
        sheet.getRange(last + 1, col).setValue(displayName);
        // schedDataを更新（後続の行番号がずれないよう）
        matchRows.push(last + 1);
        // schedDataも差し込む（簡易）
        var newRow = schedData[matchRows[0] - 1].slice();
        schedData.splice(last, 0, newRow);
      }
    });
  });

  return { ok: true };
}

// ==================== メール ====================
function sendInitialMail(teacher, entries, school, adminEmail) {
  var lines = entries.map(function(e){
    return (e.value === '△' ? '△ ' : '○ ') + e.date + ' ' + e.slot;
  }).join('\n');
  var subject = '【出勤登録】' + school + ' ' + teacher.name + 'さんの登録内容';
  var bodyTeacher = teacher.name + 'さん、以下の内容で登録しました。\n\n' + lines + '\n\n個別指導ステップ';
  var bodyAdmin   = school + '校 ' + teacher.name + 'さんが出勤登録しました。\n\n' + lines + '\n\n▼管理者画面\n' + ADMIN_URL;
  if (teacher.email) { try { MailApp.sendEmail(teacher.email, subject, bodyTeacher); } catch(ex){} }
  try { MailApp.sendEmail(adminEmail, subject, bodyAdmin); } catch(ex){}
}

function sendChangeMail(teacher, changeLog, school, adminEmail) {
  var lines = changeLog.map(function(c) {
    var mark = c.type === '削除' ? '✕' : (c.newVal === '△' ? '△' : '○');
    return mark + ' ' + c.type + '：' + c.date + ' ' + c.slot;
  }).join('\n');
  var subject = '【変更あり】' + school + ' ' + teacher.name + 'さんの出勤変更';
  var bodyTeacher = teacher.name + 'さん、以下の内容に変更しました。\n\n' + lines + '\n\n個別指導ステップ';
  var bodyAdmin   = school + '校 ' + teacher.name + 'さんが変更しました。\n\n' + lines + '\n\n▼管理者画面\n' + ADMIN_URL;
  if (teacher.email) { try { MailApp.sendEmail(teacher.email, subject, bodyTeacher); } catch(ex){} }
  try { MailApp.sendEmail(adminEmail, subject, bodyAdmin); } catch(ex){}
}

// ==================== 管理者：変更確認済み ====================
function clearTeacherChanges(teacherName, school) {
  var sheetName = (school === '大手町') ? '出勤可否_大手町' : '出勤可否';
  var sheet = SpreadsheetApp.openById(DATA_SS_ID).getSheetByName(sheetName);
  if (!sheet) return { ok: true };
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === teacherName && data[i][6]) {
      sheet.getRange(i + 1, 7).setValue('');
    }
  }
  return { ok: true };
}

// ==================== 管理者：先生別カレンダー取得 ====================
function getTeacherCalendar(school) {
  var sheetName = (school === '大手町') ? '出勤可否_大手町' : '出勤可否';
  var slotName  = (school === '大手町') ? '授業枠設定_大手町' : '授業枠設定';
  var ss = SpreadsheetApp.openById(DATA_SS_ID);

  // 授業枠の全日程を取得
  var slotSheet = ss.getSheetByName(slotName);
  if (!slotSheet) slotSheet = ss.getSheetByName('授業枠設定');
  var allDates = [];
  if (slotSheet) {
    var sd = slotSheet.getDataRange().getValues();
    var slotRows = [];
    for (var i = 1; i < sd.length; i++) {
      if (!sd[i][0]) continue;
      var d = sd[i][0] instanceof Date ? sd[i][0] : new Date(sd[i][0]);
      var raw = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
      var en = [0,1,2,3,4,5].map(function(j){ var v=sd[i][j+1]; return v===true||v==='TRUE'||Number(v)===1; });
      slotRows.push({ raw: raw, enabled: en });
    }
    slotRows.sort(function(a,b){ return a.raw<b.raw?-1:1; });
    allDates = slotRows;
  }

  // 出勤可否データを取得
  var avail = ss.getSheetByName(sheetName);
  var teachers = {};
  if (avail) {
    var rows = avail.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var name = String(rows[i][1]);
      var code = String(rows[i][0]);
      if (!name) continue;
      if (!teachers[name]) teachers[name] = { code: code, entries: {} };
      var dv = rows[i][2];
      var dk = (dv instanceof Date) ? Utilities.formatDate(dv, 'Asia/Tokyo', 'yyyy/MM/dd') : String(dv);
      var key = dk + '|' + String(rows[i][3]);
      var val  = String(rows[i][4]);
      if (val !== '') {
        teachers[name].entries[key] = { value: val, flag: String(rows[i][6]||'') };
      }
    }
  }
  return { ok: true, teachers: teachers, allDates: allDates };
}

// ==================== 管理者：授業枠設定取得 ====================
function getSlotConfig(school) {
  var sheetName = (school === '大手町') ? '授業枠設定_大手町' : '授業枠設定';
  var ss = SpreadsheetApp.openById(DATA_SS_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.getSheetByName('授業枠設定');
  var data  = sheet.getDataRange().getValues();
  var config = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var d = data[i][0] instanceof Date ? data[i][0] : new Date(data[i][0]);
    var raw = Utilities.formatDate(d,'Asia/Tokyo','yyyy/MM/dd');
    var enabled = [0,1,2,3,4,5].map(function(j){
      var v = data[i][j+1]; return v === true || v === 'TRUE' || Number(v) === 1;
    });
    config.push({ rawDate: raw, enabled: enabled });
  }
  return { ok: true, config: config };
}

// ==================== 管理者：授業枠設定保存 ====================
function saveSlotConfig(config, school) {
  var sheetName = (school === '大手町') ? '授業枠設定_大手町' : '授業枠設定';
  var ss    = SpreadsheetApp.openById(DATA_SS_ID);
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.clearContents();
  sheet.appendRow(['日付'].concat(TIME_SLOTS));
  config.forEach(function(row) {
    sheet.appendRow([new Date(row.rawDate)].concat(row.enabled));
  });
  return { ok: true };
}

// ==================== 管理者：締め切り設定保存 ====================
function saveDeadline(school, dateStr) {
  var ss    = SpreadsheetApp.openById(DATA_SS_ID);
  var sheet = ss.getSheetByName('締め切り設定');
  if (!sheet) {
    sheet = ss.insertSheet('締め切り設定');
    sheet.appendRow(['校舎', '締め切り日']);
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === school) {
      sheet.getRange(i + 1, 2).setValue(dateStr ? new Date(dateStr) : '');
      return { ok: true };
    }
  }
  sheet.appendRow([school, dateStr ? new Date(dateStr) : '']);
  return { ok: true };
}

// ==================== 管理者：締め切り取得（両校） ====================
function getDeadlines() {
  var shinryo  = getDeadline('神領');
  var otemachi = getDeadline('大手町');
  return { ok: true, 神領: shinryo.deadline, 大手町: otemachi.deadline };
}

// ==================== シート初期化 ====================
function ensureSheets() {
  var ss = SpreadsheetApp.openById(DATA_SS_ID);
  if (!ss.getSheetByName('授業枠設定')) {
    var s = ss.insertSheet('授業枠設定');
    s.appendRow(['日付'].concat(TIME_SLOTS));
    initDefaultSlots(s);
  }
  if (!ss.getSheetByName('出勤可否')) {
    var s2 = ss.insertSheet('出勤可否');
    s2.appendRow(['コード','氏名','日付','枠インデックス','可否','更新日時','変更フラグ']);
  }
  if (!ss.getSheetByName('出勤可否_大手町')) {
    var s3 = ss.insertSheet('出勤可否_大手町');
    s3.appendRow(['コード','氏名','日付','枠インデックス','可否','更新日時','変更フラグ']);
  }
  if (!ss.getSheetByName('締め切り設定')) {
    var s4 = ss.insertSheet('締め切り設定');
    s4.appendRow(['校舎', '締め切り日']);
  }
}

function initDefaultSlots(sheet) {
  var groups = [
    { dates:['2026/07/21','2026/07/22','2026/07/23','2026/07/24'],
      en:[false,false,false,true,true,true] },
    { dates:['2026/07/28','2026/07/29','2026/07/30'],
      en:[false,false,true,true,true,true] },
    { dates:['2026/07/31','2026/08/03','2026/08/04','2026/08/05','2026/08/06','2026/08/07',
             '2026/08/17','2026/08/18','2026/08/19','2026/08/20','2026/08/21',
             '2026/08/24','2026/08/25','2026/08/26','2026/08/27','2026/08/28'],
      en:[true,true,true,true,true,true] },
    { dates:OBON_DATES, en:[false,false,false,false,false,false] },
    { dates:['2026/09/01','2026/09/02','2026/09/03','2026/09/04'],
      en:[false,false,false,true,true,true] }
  ];
  groups.forEach(function(g) {
    g.dates.forEach(function(d) { sheet.appendRow([new Date(d)].concat(g.en)); });
  });
}

// ==================== 先生用HTML ====================
function teacherHtml() {
  var h = '';
  h += "<!DOCTYPE html><html lang='ja'><head>\n";
  h += "<meta charset='UTF-8'><meta name='viewport' content='width=device-width,initial-scale=1'>\n";
  h += "<title>夏休み出勤可能日登録</title>\n";
  h += "<style>\n";
  h += "*{box-sizing:border-box;margin:0;padding:0}\n";
  h += "body{font-family:sans-serif;font-size:14px;background:#EEF2FF;min-height:100vh}\n";
  h += ".hero{background:linear-gradient(135deg,#1565C0,#42A5F5);color:#fff;padding:28px 20px 20px;text-align:center}\n";
  h += ".hero h1{font-size:22px;font-weight:bold;margin-bottom:8px;letter-spacing:.05em}\n";
  h += ".hero p{font-size:12px;opacity:.85}\n";
  h += ".illust{margin:12px auto 0;line-height:1}\n";
  h += ".brand{text-align:center;padding:24px 0 32px;color:#90A4AE;font-size:13px;letter-spacing:.1em}\n";
  h += ".card{background:#fff;border-radius:12px;padding:18px;margin:16px;box-shadow:0 2px 10px rgba(0,0,0,.1)}\n";
  h += "label{display:block;font-weight:bold;margin-bottom:8px;color:#555;font-size:13px}\n";
  h += "input{width:100%;padding:13px;border:1.5px solid #ccc;border-radius:8px;font-size:18px;background:#fafafa;text-align:center}\n";
  h += "input:focus{outline:none;border-color:#1976D2}\n";
  h += ".err{color:#e53935;margin-top:8px;font-size:13px;min-height:18px;text-align:center}\n";
  // 校舎選択ボタン
  h += ".school-sel{display:flex;gap:12px;margin:16px;}\n";
  h += ".school-btn{flex:1;padding:16px;border:2.5px solid #ccc;border-radius:12px;font-size:16px;font-weight:bold;cursor:pointer;background:#fff;color:#555;transition:.15s;text-align:center}\n";
  h += ".school-btn.selected{border-color:#1565C0;background:#1565C0;color:#fff;box-shadow:0 3px 10px rgba(21,101,192,.4)}\n";
  h += ".deadline-banner{margin:0 16px 4px;padding:10px 14px;border-radius:8px;font-size:13px;font-weight:bold;text-align:center}\n";
  h += ".deadline-ok{background:#E8F5E9;color:#2E7D32;border:1px solid #A5D6A7}\n";
  h += ".deadline-warn{background:#FFF8E1;color:#F57F17;border:1px solid #FFE082}\n";
  h += ".deadline-urgent{background:#FFEBEE;color:#C62828;border:1px solid #EF9A9A}\n";
  h += ".deadline-locked{background:#ECEFF1;color:#546E7A;border:1px solid #B0BEC5}\n";
  h += ".btn-login{display:block;width:calc(100% - 32px);margin:8px 16px 0;padding:15px;border:none;border-radius:12px;font-size:17px;font-weight:bold;cursor:pointer;background:#1565C0;color:#fff;box-shadow:0 3px 10px rgba(21,101,192,.4);transition:.15s}\n";
  h += ".btn-login:active{transform:scale(.97)}\n";
  h += ".btn-login:disabled{background:#aaa;box-shadow:none}\n";
  h += ".gbar{position:sticky;top:0;z-index:10;background:linear-gradient(90deg,#1565C0,#42A5F5);color:#fff;padding:11px 16px;font-size:17px;font-weight:bold;box-shadow:0 2px 6px rgba(0,0,0,.2)}\n";
  h += ".locked-msg{background:#ECEFF1;color:#546E7A;padding:14px 16px;font-size:13px;font-weight:bold;border-left:4px solid #90A4AE;margin:8px 0 0;text-align:center}\n";
  h += ".legend{display:flex;gap:10px;padding:8px 12px;background:#fff;font-size:12px;flex-wrap:wrap;align-items:center;border-bottom:1px solid #e0e0e0}\n";
  h += ".lb{width:28px;height:28px;border-radius:6px;display:inline-flex;align-items:center;justify-content:center;font-size:15px;font-weight:bold;margin-right:2px}\n";
  h += ".table-wrap{overflow-x:auto;margin-top:0}\n";
  h += "table{border-collapse:collapse;font-size:12px;min-width:420px;width:100%}\n";
  h += "thead th{background:#37474f;color:#fff;padding:8px 3px;text-align:center;white-space:nowrap;position:sticky;top:0;z-index:3}\n";
  h += "thead th.th-d{left:0;z-index:4;min-width:68px;text-align:left;padding-left:6px}\n";
  h += "td{border:1px solid #ddd;padding:4px 2px;text-align:center;vertical-align:middle}\n";
  h += "tr:nth-child(even) td{background:#f9f9f9}\n";
  h += "td.dc{text-align:left;font-weight:bold;white-space:nowrap;padding:5px 6px;min-width:68px;position:sticky;left:0;background:#fff;z-index:1;font-size:11px}\n";
  h += "tr:nth-child(even) td.dc{background:#f9f9f9}\n";
  h += ".dash{color:#bbb;font-size:12px}\n";
  h += ".sb{width:40px;height:40px;border-radius:8px;border:2px solid #ddd;background:#fff;font-size:16px;font-weight:bold;cursor:pointer;display:flex;align-items:center;justify-content:center;margin:auto;transition:.1s}\n";
  h += ".sb.ci{background:#1976D2;border-color:#1976D2;color:#fff}\n";
  h += ".sb.tr{background:#F57C00;border-color:#F57C00;color:#fff}\n";
  h += ".sb:disabled{opacity:.3;cursor:default}\n";
  h += ".obon td{background:#E3F2FD!important;color:#0D47A1;font-weight:bold;text-align:center;padding:9px;font-size:13px}\n";
  h += ".save-area{padding:12px 16px 8px}\n";
  h += ".btn-send{display:block;width:100%;padding:14px;border:none;border-radius:12px;font-size:16px;font-weight:bold;cursor:pointer;background:#2E7D32;color:#fff;box-shadow:0 3px 10px rgba(46,125,50,.4);transition:.15s}\n";
  h += ".btn-send:active{transform:scale(.97);background:#1B5E20}\n";
  h += ".btn-send:disabled{background:#aaa;box-shadow:none;cursor:default}\n";
  h += ".msg{margin:4px 16px 16px;padding:12px;border-radius:8px;display:none;white-space:pre-wrap;font-size:13px}\n";
  h += ".msg-ok{background:#e8f5e9;color:#2e7d32}\n";
  h += ".msg-err{background:#ffebee;color:#c62828}\n";
  h += ".loading{text-align:center;color:#777;padding:20px}\n";
  h += ".popup{position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:999;display:flex;align-items:center;justify-content:center;display:none}\n";
  h += ".popup-box{background:#fff;border-radius:20px;padding:36px 40px;text-align:center;box-shadow:0 12px 40px rgba(0,0,0,.25);animation:pop .3s ease}\n";
  h += ".popup-ico{font-size:64px;margin-bottom:10px}\n";
  h += ".popup-txt{font-size:22px;font-weight:bold;color:#2E7D32}\n";
  h += "@keyframes pop{from{transform:scale(.5);opacity:0}to{transform:scale(1);opacity:1}}\n";
  h += "</style></head><body>\n";

  // === ログイン画面 ===
  h += "<div id='ls'>\n";
  h += "<div class='hero'>\n";
  h += "<h1>☀️ 夏休み出勤可能日登録</h1>\n";
  h += "<p>個別指導ステップ</p>\n";
  h += "<div class='illust'>";
  h += "<svg viewBox='0 0 300 120' xmlns='http://www.w3.org/2000/svg' width='100%' style='max-width:300px'>";
  h += "<rect width='300' height='80' fill='#87CEEB'/>";
  h += "<ellipse cx='150' cy='44' rx='28' ry='28' fill='#FFD700'/>";
  h += "<line x1='150' y1='5' x2='150' y2='0' stroke='#FFD700' stroke-width='4' stroke-linecap='round'/>";
  h += "<line x1='185' y1='18' x2='189' y2='14' stroke='#FFD700' stroke-width='4' stroke-linecap='round'/>";
  h += "<line x1='198' y1='50' x2='204' y2='50' stroke='#FFD700' stroke-width='4' stroke-linecap='round'/>";
  h += "<line x1='185' y1='76' x2='189' y2='80' stroke='#FFD700' stroke-width='4' stroke-linecap='round'/>";
  h += "<line x1='115' y1='18' x2='111' y2='14' stroke='#FFD700' stroke-width='4' stroke-linecap='round'/>";
  h += "<line x1='102' y1='50' x2='96' y2='50' stroke='#FFD700' stroke-width='4' stroke-linecap='round'/>";
  h += "<line x1='115' y1='76' x2='111' y2='80' stroke='#FFD700' stroke-width='4' stroke-linecap='round'/>";
  h += "<rect x='0' y='80' width='300' height='40' fill='#1565C0'/>";
  h += "<path d='M0 80 Q37 70 75 80 Q112 90 150 80 Q187 70 225 80 Q262 90 300 80' fill='#42A5F5'/>";
  h += "<path d='M0 90 Q37 80 75 90 Q112 100 150 90 Q187 80 225 90 Q262 100 300 90' fill='#64B5F6' opacity='.6'/>";
  h += "<text x='50' y='108' fill='#fff' font-size='12' font-family='sans-serif'>🐠</text>";
  h += "<text x='220' y='112' fill='#fff' font-size='16' font-family='sans-serif'>🐚</text>";
  h += "</svg></div>\n";
  h += "</div>\n";

  // 校舎選択
  h += "<p style='text-align:center;font-size:13px;font-weight:bold;color:#555;margin:16px 0 8px'>校舎を選んでください</p>\n";
  h += "<div class='school-sel'>\n";
  h += "  <button class='school-btn' id='btn-shinryo' onclick='selectSchool(\"神領\")'>🏫 神領校</button>\n";
  h += "  <button class='school-btn' id='btn-otemachi' onclick='selectSchool(\"大手町\")'>🏫 大手町校</button>\n";
  h += "</div>\n";

  // 締め切りバナー（校舎選択後に表示）
  h += "<div id='dl-banner' style='display:none'></div>\n";

  h += "<div class='card'>\n";
  h += "<label>コード番号（半角数字）</label>\n";
  h += "<input type='number' id='code' placeholder='例：7002' inputmode='numeric'>\n";
  h += "<div class='err' id='lerr'></div>\n";
  h += "</div>\n";
  h += "<button class='btn-login' id='loginbtn' onclick='doLogin()' disabled>ログイン</button>\n";
  h += "<div class='brand'>個別指導ステップ</div>\n";
  h += "</div>\n";

  // === メイン画面 ===
  h += "<div id='ms' style='display:none'>\n";
  h += "<div class='gbar' id='gbar'></div>\n";
  h += "<div id='lock-msg' style='display:none'></div>\n";
  h += "<div class='legend' id='legend-area'>\n";
  h += "  <span>凡例：</span>\n";
  h += "  <span><span class='lb' style='background:#1976D2;color:#fff'>○</span>出勤できる</span>\n";
  h += "  <span><span class='lb' style='background:#F57C00;color:#fff'>△</span>無理すれば可</span>\n";
  h += "  <span><span class='lb' style='border:2px solid #ddd'></span>未入力（もう一度押すと戻る）</span>\n";
  h += "</div>\n";
  h += "<div class='table-wrap'><div id='ta'><div class='loading'>読み込み中...</div></div></div>\n";
  h += "<div class='save-area' id='save-area'>\n";
  h += "<button class='btn-send' id='sendbtn' onclick='saveData()'>📨 保存＆送信する</button>\n";
  h += "</div>\n";
  h += "<div class='msg' id='msg'></div>\n";
  h += "</div>\n";

  // === 送信ポップアップ ===
  h += "<div class='popup' id='popup'><div class='popup-box'><div class='popup-ico'>📧</div><div class='popup-txt'>送信しました！</div></div></div>\n";

  // === JavaScript ===
  h += "<script>\n";
  h += "var code='',name_='',school='',slots=[],avail={},isLocked_=false;\n";
  h += "function lsGet(k){try{return localStorage.getItem(k)||'';}catch(e){return '';}}\n";
  h += "function lsSet(k,v){try{localStorage.setItem(k,v);}catch(e){}}\n";
  h += "(function(){var s=lsGet('step_code');if(s)document.getElementById('code').value=s;})();\n";

  // 校舎選択
  h += "function selectSchool(s){\n";
  h += "  school=s;\n";
  h += "  document.getElementById('btn-shinryo').className='school-btn'+(s==='神領'?' selected':'');\n";
  h += "  document.getElementById('btn-otemachi').className='school-btn'+(s==='大手町'?' selected':'');\n";
  h += "  document.getElementById('loginbtn').disabled=true;\n";
  h += "  document.getElementById('dl-banner').style.display='none';\n";
  h += "  google.script.run.withSuccessHandler(function(r){\n";
  h += "    showDeadlineBanner(r);\n";
  h += "    document.getElementById('loginbtn').disabled=false;\n";
  h += "  }).checkDeadline(s);\n";
  h += "}\n";

  // 締め切りバナー表示
  h += "function showDeadlineBanner(r){\n";
  h += "  var b=document.getElementById('dl-banner');\n";
  h += "  if(!r.deadline){b.style.display='none';return;}\n";
  h += "  var cls,txt;\n";
  h += "  if(r.locked){\n";
  h += "    cls='deadline-locked';txt='🔒 登録受付は終了しました（締め切り：'+r.deadline+'）';\n";
  h += "  }else if(r.daysLeft===0){\n";
  h += "    cls='deadline-urgent';txt='⚠️ 本日が締め切りです！（'+r.deadline+'）';\n";
  h += "  }else if(r.daysLeft<=3){\n";
  h += "    cls='deadline-warn';txt='⏰ 登録・変更締め切りまであと'+r.daysLeft+'日（'+r.deadline+'）';\n";
  h += "  }else{\n";
  h += "    cls='deadline-ok';txt='📅 登録・変更締め切り：'+r.deadline+'（あと'+r.daysLeft+'日）';\n";
  h += "  }\n";
  h += "  b.className='deadline-banner '+cls;\n";
  h += "  b.textContent=txt;\n";
  h += "  b.style.display='block';\n";
  h += "}\n";

  h += "function doLogin(){\n";
  h += "  if(!school){setErr('校舎を選択してください');return;}\n";
  h += "  var c=document.getElementById('code').value.trim();\n";
  h += "  if(!c){setErr('コード番号を入力してください');return;}\n";
  h += "  setErr('照合中...');\n";
  h += "  google.script.run\n";
  h += "    .withSuccessHandler(function(r){\n";
  h += "      if(r.ok){code=c;name_=r.name;lsSet('step_code',c);\n";
  h += "        document.getElementById('gbar').textContent=school+'校 '+r.name+'さん、こんにちは！';\n";
  h += "        google.script.run.withSuccessHandler(function(dl){\n";
  h += "          var pop=document.getElementById('dl-popup');\n";
  h += "          var msg=document.getElementById('dl-popup-msg');\n";
  h += "          if(dl.locked){\n";
  h += "            msg.textContent='登録締切日を過ぎています。このアプリからの登録・変更はできません。';\n";
  h += "          } else {\n";
  h += "            var dLeft=dl.daysLeft;\n";
  h += "            if(dLeft===0)msg.textContent='本日が登録締切日です。忘れずに登録してください。';\n";
  h += "            else if(dLeft>0)msg.textContent='登録締切まで残り'+dLeft+'日です。期日までに登録してください。';\n";
  h += "            else msg.textContent='登録期間中です。出勤可能な日程を登録してください。';\n";
  h += "          }\n";
  h += "          document.getElementById('dl-popup-icon').textContent=dl.locked?'🔒':'📅';\n";
  h += "          pop.style.display='flex';\n";
  h += "        }).checkDeadline(school);\n";
  h += "        loadData();\n";
  h += "      }else{setErr(r.error);}\n";
  h += "    })\n";
  h += "    .withFailureHandler(function(e){setErr('エラー:'+e.message);})\n";
  h += "    .login(c);\n";
  h += "}\n";
  h += "function setErr(t){document.getElementById('lerr').textContent=t;}\n";

  h += "function loadData(){\n";
  h += "  google.script.run\n";
  h += "    .withSuccessHandler(function(dlr){\n";
  h += "      isLocked_=dlr.locked;\n";
  h += "      google.script.run\n";
  h += "        .withSuccessHandler(function(r){\n";
  h += "          slots=r.slots;\n";
  h += "          google.script.run\n";
  h += "            .withSuccessHandler(function(r2){\n";
  h += "              avail=r2.entries||{};\n";
  h += "              var local=JSON.parse(lsGet('step_avail_'+code+'_'+school)||'{}');\n";
  h += "              Object.keys(local).forEach(function(k){if(!(k in avail))avail[k]=local[k];});\n";
  h += "              renderTable();\n";
  h += "              document.getElementById('ls').style.display='none';\n";
  h += "              document.getElementById('ms').style.display='block';\n";
  h += "              if(isLocked_){\n";
  h += "                document.getElementById('lock-msg').innerHTML='<div class=\"locked-msg\">🔒 締め切りを過ぎているため、変更はできません。<br>変更がある場合は直接担当者にご連絡ください。</div>';\n";
  h += "                document.getElementById('lock-msg').style.display='block';\n";
  h += "                document.getElementById('save-area').style.display='none';\n";
  h += "              }\n";
  h += "            }).getMyAvailability(code,school);\n";
  h += "        }).getSlots(school);\n";
  h += "    }).checkDeadline(school);\n";
  h += "}\n";

  h += "function toggle(raw,si){\n";
  h += "  if(isLocked_)return;\n";
  h += "  var key=raw+'|'+si;\n";
  h += "  var cur=avail[key]||'';\n";
  h += "  avail[key]=cur===''?'○':cur==='○'?'△':'';\n";
  h += "  lsSet('step_avail_'+code+'_'+school,JSON.stringify(avail));\n";
  h += "  renderTable();\n";
  h += "}\n";

  h += "function renderTable(){\n";
  h += "  var T1=['13:00','14:20','15:45','17:10','18:35','20:00'];\n";
  h += "  var T2=['14:15','15:35','17:00','18:25','19:50','21:15'];\n";
  h += "  var h='<table><thead><tr><th class=\"th-d\">日付</th>';\n";
  h += "  for(var x=0;x<6;x++)h+='<th>'+T1[x]+'<br>'+T2[x]+'</th>';\n";
  h += "  h+='</tr></thead><tbody>';\n";
  h += "  slots.forEach(function(sl){\n";
  h += "    if(sl.holiday){h+='<tr class=\"obon\"><td colspan=\"7\">🌊 '+sl.date+'</td></tr>';return;}\n";
  h += "    h+='<tr><td class=\"dc\">'+sl.date+'</td>';\n";
  h += "    for(var i=0;i<6;i++){\n";
  h += "      h+='<td>';\n";
  h += "      if(!sl.enabled[i]){h+='<span class=\"dash\">ー</span>';}\n";
  h += "      else{\n";
  h += "        var key=sl.rawDate+'|'+i;\n";
  h += "        var val=avail[key]||'';\n";
  h += "        var cls=val==='○'?'ci':val==='△'?'tr':'';\n";
  h += "        h+='<button class=\"sb '+cls+'\"'+(isLocked_?' disabled':'')+' data-raw=\"'+sl.rawDate+'\" data-slot=\"'+i+'\">'+val+'</button>';\n";
  h += "      }\n";
  h += "      h+='</td>';\n";
  h += "    }\n";
  h += "    h+='</tr>';\n";
  h += "  });\n";
  h += "  h+='</tbody></table>';\n";
  h += "  document.getElementById('ta').innerHTML=h;\n";
  h += "}\n";

  h += "document.addEventListener('click',function(e){\n";
  h += "  var btn=e.target.closest('.sb');\n";
  h += "  if(!btn||btn.disabled)return;\n";
  h += "  toggle(btn.dataset.raw,parseInt(btn.dataset.slot));\n";
  h += "});\n";

  h += "function saveData(){\n";
  h += "  if(isLocked_){showMsg('締め切りを過ぎているため変更できません。',false);return;}\n";
  h += "  var sb=document.getElementById('sendbtn');\n";
  h += "  sb.disabled=true;\n";
  h += "  sb.textContent='保存中...';\n";
  h += "  var entries=[];\n";
  h += "  slots.forEach(function(sl){\n";
  h += "    if(sl.holiday)return;\n";
  h += "    for(var i=0;i<6;i++){\n";
  h += "      if(!sl.enabled[i])continue;\n";
  h += "      entries.push({rawDate:sl.rawDate,date:sl.date,slotIndex:i,value:avail[sl.rawDate+'|'+i]||''});\n";
  h += "    }\n";
  h += "  });\n";
  h += "  google.script.run\n";
  h += "    .withSuccessHandler(function(r){\n";
  h += "      sb.disabled=false;\n";
  h += "      sb.textContent='📨 保存＆送信する';\n";
  h += "      if(r.ok){\n";
  h += "        showPopup();\n";
  h += "        showMsg('保存・送信しました！',true);\n";
  h += "        lsSet('step_avail_'+code+'_'+school,'{}');\n";
  h += "      }else{showMsg('エラー: '+r.error,false);}\n";
  h += "    })\n";
  h += "    .withFailureHandler(function(e){\n";
  h += "      sb.disabled=false;\n";
  h += "      sb.textContent='📨 保存＆送信する';\n";
  h += "      showMsg('エラー: '+e.message,false);\n";
  h += "    })\n";
  h += "    .saveAvailability(code,school,entries,true);\n";
  h += "}\n";

  h += "function showMsg(t,ok){var el=document.getElementById('msg');el.className='msg '+(ok?'msg-ok':'msg-err');el.textContent=t;el.style.display='block';}\n";
  h += "function showPopup(){var p=document.getElementById('popup');p.style.display='flex';setTimeout(function(){p.style.display='none';},2500);}\n";
  h += "</script></body></html>";
  return h;
}

// ==================== 管理者用HTML ====================
function adminHtml() { return adminHtmlStr(); }

function adminHtmlStr() {
  var h = '';
  h += "<!DOCTYPE html><html lang='ja'><head>\n";
  h += "<meta charset='UTF-8'><meta name='viewport' content='width=device-width,initial-scale=1'>\n";
  h += "<title>管理者画面 - ステップ</title>\n";
  h += "<style>\n";
  h += "*{box-sizing:border-box;margin:0;padding:0}\n";
  h += "body{font-family:sans-serif;font-size:14px;background:#f5f5f5}\n";
  h += ".hdr{background:linear-gradient(135deg,#1565C0,#42A5F5);color:#fff;padding:16px}\n";
  h += ".hdr h1{font-size:18px;font-weight:bold}\n";
  h += ".hdr p{font-size:12px;opacity:.85;margin-top:4px}\n";
  h += ".tabs{display:flex;gap:4px;padding:0 12px;margin-bottom:12px;margin-top:12px;flex-wrap:wrap}\n";
  h += ".tab{padding:10px 18px;border:none;border-radius:8px 8px 0 0;cursor:pointer;font-size:13px;font-weight:bold;background:#ddd;color:#555}\n";
  h += ".tab.active{background:#fff;color:#1565C0;box-shadow:0 -2px 6px rgba(0,0,0,.1)}\n";
  h += ".school-tabs{display:flex;gap:6px;margin:0 12px 12px}\n";
  h += ".stab{padding:8px 20px;border:2px solid #ccc;border-radius:20px;cursor:pointer;font-size:13px;font-weight:bold;background:#fff;color:#666}\n";
  h += ".stab.active{border-color:#1565C0;background:#1565C0;color:#fff}\n";
  h += ".body{padding:0 12px 20px}\n";
  h += ".btn-g{display:block;width:100%;padding:12px;border:none;border-radius:8px;font-size:15px;font-weight:bold;cursor:pointer;margin-top:8px;background:#4CAF50;color:#fff}\n";
  h += ".btn-bulk{display:block;width:100%;padding:14px;border:none;border-radius:8px;font-size:15px;font-weight:bold;cursor:pointer;margin-top:12px;background:#E65100;color:#fff;box-shadow:0 3px 10px rgba(230,81,0,.4)}\n";
  h += ".btn-sm{display:inline-block;padding:4px 10px;font-size:12px;border-radius:4px;background:#e53935;color:#fff;border:none;cursor:pointer}\n";
  h += ".btn-print{display:inline-block;padding:8px 16px;font-size:13px;border-radius:6px;background:#1565C0;color:#fff;border:none;cursor:pointer;margin-right:8px;margin-bottom:12px}\n";
  h += ".btn-sel{display:inline-block;padding:8px 16px;font-size:13px;border-radius:6px;background:#455A64;color:#fff;border:none;cursor:pointer;margin-bottom:12px}\n";
  h += ".btn-conf{padding:6px 12px;font-size:12px;border-radius:6px;background:#607D8B;color:#fff;border:none;cursor:pointer;margin-top:8px}\n";
  h += ".msg{margin-top:12px;padding:10px;border-radius:6px;display:none}\n";
  h += ".msg-ok{background:#e8f5e9;color:#2e7d32}\n";
  h += ".msg-err{background:#ffebee;color:#c62828}\n";
  h += ".tw{overflow-x:auto;background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.1)}\n";
  h += "table{border-collapse:collapse;font-size:12px;width:100%}\n";
  h += "th{background:#37474f;color:#fff;padding:8px 4px;text-align:center;white-space:nowrap}\n";
  h += "td{border:1px solid #ddd;padding:5px 3px;text-align:center;vertical-align:middle}\n";
  h += ".dc{text-align:left;padding:5px 8px;min-width:80px;white-space:nowrap;font-weight:bold}\n";
  h += "input[type=checkbox]{width:16px;height:16px;cursor:pointer}\n";
  h += ".add-row{display:flex;gap:8px;margin-top:12px;align-items:center}\n";
  h += ".add-row input[type=date]{flex:1;padding:8px;border:1px solid #ccc;border-radius:6px;font-size:14px}\n";
  h += ".add-btn{padding:9px 16px;font-size:14px;border:none;border-radius:6px;background:#2196F3;color:#fff;cursor:pointer}\n";
  h += ".tc{background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.1);margin-bottom:16px;overflow:hidden}\n";
  h += ".tc h2{background:#37474f;color:#fff;padding:10px 12px;font-size:15px}\n";
  h += ".dot{display:inline-block;width:22px;height:22px;border-radius:4px;line-height:22px;font-weight:bold;font-size:12px}\n";
  h += ".dot.ci{background:#1976D2;color:#fff}\n";
  h += ".dot.tr{background:#F57C00;color:#fff}\n";
  h += ".dot.add{background:#C8E6C9;color:#1B5E20;outline:2px solid #4CAF50}\n";
  h += ".dot.chg{background:#FFF9C4;color:#F57F17;outline:2px solid #FFC107}\n";
  h += ".dot.del{background:#FFCDD2;color:#B71C1C;outline:2px solid #E53935;text-decoration:line-through}\n";
  h += ".cal-legend{font-size:10px;padding:4px 10px;background:#f9f9f9;border-bottom:1px solid #eee;display:flex;gap:10px;flex-wrap:wrap}\n";
  h += ".leg{display:inline-flex;align-items:center;gap:3px}\n";
  h += ".swatch{width:12px;height:12px;border-radius:2px;display:inline-block}\n";
  h += ".tc table th{padding:4px 2px;font-size:11px}\n";
  h += ".tc table td{padding:2px 1px;font-size:11px}\n";
  h += ".tc table td.dc{padding:3px 5px;font-size:11px;min-width:60px}\n";
  // 締め切り設定セクション
  h += ".dl-box{background:#fff;border-radius:8px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.1);margin-bottom:12px}\n";
  h += ".dl-box h3{font-size:14px;font-weight:bold;color:#333;margin-bottom:10px}\n";
  h += ".dl-row{display:flex;align-items:center;gap:8px;margin-bottom:8px;flex-wrap:wrap}\n";
  h += ".dl-row label{min-width:70px;font-size:13px;color:#555}\n";
  h += ".dl-row input[type=date]{flex:1;padding:8px;border:1px solid #ccc;border-radius:6px;font-size:14px;min-width:0}\n";
  h += ".btn-dl-save{padding:8px 20px;border:none;border-radius:6px;background:#1565C0;color:#fff;font-size:13px;font-weight:bold;cursor:pointer}\n";
  h += "@media print{\n";
  h += "  .no-print,.tabs,.school-tabs,.add-row,.btn-g,.btn-bulk,.btn-print,.btn-sel,.dl-box{display:none!important}\n";
  h += "  .tc{break-inside:avoid;break-after:page;margin-bottom:0;box-shadow:none;border:1px solid #000}\n";
  h += "  .tc:last-child{break-after:auto}\n";
  h += "  body,html{background:#fff}\n";
  h += "  *{color:#000!important}\n";
  h += "  .tc h2{background:#fff!important;color:#000!important;border-bottom:2px solid #000}\n";
  h += "  th{background:#fff!important;color:#000!important;border:1px solid #000!important}\n";
  h += "  td{border:1px solid #000!important}\n";
  h += "  .dot.ci,.dot.tr{background:#fff!important;color:#000!important;border:1px solid #000!important}\n";
  h += "  .hide-print{display:none!important}\n";
  h += "  @page{size:A4;margin:10mm}\n";
  h += "}\n";
  h += "</style></head><body>\n";

  h += "<div class='hdr'><h1>出勤日管理</h1><p>個別指導ステップ 2026年 夏期講習</p></div>\n";

  h += "<div class='tabs no-print'>\n";
  h += "  <button class='tab active' id='btn-cfg' onclick='showConfig()'>授業枠設定</button>\n";
  h += "  <button class='tab' id='btn-cal' onclick='showCalendar()'>先生別カレンダー</button>\n";
  h += "  <button class='tab' id='btn-dl' onclick='showDeadline()'>締め切り設定</button>\n";
  h += "</div>\n";

  // 校舎切り替え
  h += "<div class='school-tabs no-print'>\n";
  h += "  <button class='stab active' id='stab-shinryo' onclick='switchSchool(\"神領\")'>神領校</button>\n";
  h += "  <button class='stab' id='stab-otemachi' onclick='switchSchool(\"大手町\")'>大手町校</button>\n";
  h += "</div>\n";

  h += "<div class='body'>\n";

  // ---- 授業枠設定タブ ----
  h += "<div id='tab-cfg'>\n";
  h += "<p style='font-size:12px;color:#666;margin-bottom:10px'>各日付の実施する時間枠にチェックを入れて「保存する」を押してください。</p>\n";
  h += "<div class='tw'><table>\n";
  h += "<thead><tr><th class='dc'>日付</th>";
  h += "<th>13:00<br>14:15</th><th>14:20<br>15:35</th><th>15:45<br>17:00</th>";
  h += "<th>17:10<br>18:25</th><th>18:35<br>19:50</th><th>20:00<br>21:15</th>";
  h += "<th class='no-print'>削除</th></tr></thead>\n";
  h += "<tbody id='cb'></tbody></table></div>\n";
  h += "<div class='add-row no-print'><input type='date' id='nd'><button class='add-btn' onclick='addDate()'>日付を追加</button></div>\n";
  h += "<button class='btn-g no-print' onclick='save()' style='margin-top:10px'>保存する</button>\n";
  h += "<div class='msg' id='cmsg'></div>\n";
  h += "</div>\n";

  // ---- カレンダータブ ----
  h += "<div id='tab-cal' style='display:none'>\n";
  h += "<div class='no-print' style='margin-bottom:8px'>\n";
  h += "  <button class='btn-print' onclick='window.print()'>🖨️ 印刷する</button>\n";
  h += "  <button class='btn-sel' onclick='showPrintSel()'>👤 印刷する先生を選ぶ</button>\n";
  h += "  <button class='btn-sel' onclick='toggleFilter()' style='background:#00796B'>🔍 先生を絞り込む</button>\n";
  h += "</div>\n";
  // 一括登録ボタン
  h += "<button class='btn-bulk no-print' onclick='doBulkWrite()'>📋 時間割に一括登録する</button>\n";
  h += "<div id='bulk-msg' class='msg' style='margin-top:8px'></div>\n";
  h += "<div id='bpop' style='display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:999;align-items:center;justify-content:center'><div style='background:#fff;border-radius:20px;padding:50px 60px;text-align:center'><div style='font-size:80px'>📋</div><div id='bpop-msg' style='font-size:32px;font-weight:bold;color:#E65100;margin-top:16px'></div></div></div>\n";
  h += "<div id='teacher-filter' style='display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:200;align-items:center;justify-content:center'>\n";
  h += "  <div style='background:#fff;border-radius:12px;padding:24px;max-width:380px;width:90%;max-height:80vh;overflow-y:auto'>\n";
  h += "    <h3 style='margin-bottom:14px;font-size:16px'>表示する先生を選択</h3>\n";
  h += "    <div id='filter-checks' style='margin-bottom:16px;max-height:300px;overflow-y:auto'></div>\n";
  h += "    <div style='display:flex;gap:8px;flex-wrap:wrap'>\n";
  h += "      <button onclick='filterSelAll()' style='flex:1;padding:10px;border:1px solid #ccc;border-radius:8px;cursor:pointer'>全選択</button>\n";
  h += "      <button onclick='filterSelNone()' style='flex:1;padding:10px;border:1px solid #ccc;border-radius:8px;cursor:pointer'>全解除</button>\n";
  h += "    </div>\n";
  h += "    <div style='display:flex;gap:8px;margin-top:8px'>\n";
  h += "      <button onclick='applyFilter()' style='flex:1;padding:10px;border:none;border-radius:8px;background:#1565C0;color:#fff;font-weight:bold;cursor:pointer'>表示を更新</button>\n";
  h += "      <button onclick='document.getElementById(\"teacher-filter\").style.display=\"none\"' style='flex:1;padding:10px;border:1px solid #ccc;border-radius:8px;cursor:pointer'>閉じる</button>\n";
  h += "    </div>\n";
  h += "  </div>\n";
  h += "</div>\n";
  h += "</div>\n";
  h += "<div id='cal-area'><p style='color:#999;padding:16px'>読み込み中...</p></div>\n";
  h += "</div>\n";

  // ---- 締め切り設定タブ ----
  h += "<div id='tab-dl' style='display:none'>\n";
  h += "<p style='font-size:12px;color:#666;margin-bottom:12px'>各校舎の登録締め切り日を設定してください。締め切りを過ぎると講師アプリから変更できなくなります。</p>\n";
  h += "<div class='dl-box'>\n";
  h += "  <h3>🏫 神領校</h3>\n";
  h += "  <div class='dl-row'><label>締め切り日</label><input type='date' id='dl-shinryo' onchange='showDow(this,\"dow-s\")'></div>\n";
  h += "  <div id='dow-s' style='font-size:18px;font-weight:bold;color:#1565C0;margin:4px 0 8px'></div>\n";
  h += "  <button class='btn-dl-save' onclick='saveDeadlineUI(\"神領\",\"dl-shinryo\")'>保存</button>\n";
  h += "</div>\n";
  h += "<div class='dl-box'>\n";
  h += "  <h3>🏫 大手町校</h3>\n";
  h += "  <div class='dl-row'><label>締め切り日</label><input type='date' id='dl-otemachi' onchange='showDow(this,\"dow-o\")'></div>\n";
  h += "  <div id='dow-o' style='font-size:18px;font-weight:bold;color:#1565C0;margin:4px 0 8px'></div>\n";
  h += "  <button class='btn-dl-save' onclick='saveDeadlineUI(\"大手町\",\"dl-otemachi\")'>保存</button>\n";
  h += "</div>\n";
  h += "<div id='dl-msg' class='msg'></div>\n";
  h += "</div>\n";

  // ---- 授業枠保存ポップアップ ----
  h += "<div id='spop' style='display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:200;align-items:center;justify-content:center'>";
  h += "<div style='background:#fff;border-radius:16px;padding:32px 40px;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,.2);animation:pop .3s ease'>";
  h += "<div style='font-size:48px;margin-bottom:10px'>✅</div>";
  h += "<div style='font-size:20px;font-weight:bold;color:#2E7D32'>保存しました！</div>";
  h += "</div></div>\n";
  // 印刷対象選択ポップアップ
  h += "<div id='psel' style='display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:100;align-items:center;justify-content:center'>\n";
  h += "<div style='background:#fff;border-radius:12px;padding:24px;max-width:360px;width:90%;max-height:80vh;overflow-y:auto'>\n";
  h += "<h3 style='margin-bottom:14px;font-size:16px'>印刷する先生を選択</h3>\n";
  h += "<div id='pchks' style='margin-bottom:16px'></div>\n";
  h += "<div style='display:flex;gap:8px'>\n";
  h += "<button onclick='selAll()' style='flex:1;padding:10px;border:1px solid #ccc;border-radius:8px;cursor:pointer;background:#f5f5f5'>全選択</button>\n";
  h += "<button onclick='doPrint()' style='flex:1;padding:10px;border:none;border-radius:8px;cursor:pointer;background:#1565C0;color:#fff;font-weight:bold'>印刷</button>\n";
  h += "<button onclick='closePsel()' style='flex:1;padding:10px;border:1px solid #ccc;border-radius:8px;cursor:pointer;background:#f5f5f5'>キャンセル</button>\n";
  h += "</div></div></div>\n";

  // ---- JavaScript ----
  h += "<script>\n";
  h += "var config=[],calData={},currentSchool='神領';\n";
  h += "var DAYS=['日','月','火','水','木','金','土'];\n";
  h += "(function(){loadConfig();loadDeadlines();})();\n";

  // 校舎切り替え
  h += "function switchSchool(s){\n";
  h += "  currentSchool=s;\n";
  h += "  document.getElementById('stab-shinryo').className='stab'+(s==='神領'?' active':'');\n";
  h += "  document.getElementById('stab-otemachi').className='stab'+(s==='大手町'?' active':'');\n";
  h += "  // 現在開いているタブを再読み込み\n";
  h += "  var active=document.querySelector('.tab.active');\n";
  h += "  if(active)active.click();\n";
  h += "}\n";

  h += "function loadConfig(){\n";
  h += "  google.script.run.withSuccessHandler(function(r){config=r.config;render();}).getSlotConfig(currentSchool);\n";
  h += "}\n";

  h += "function render(){\n";
  h += "  var tb=document.getElementById('cb');tb.innerHTML='';\n";
  h += "  config.forEach(function(row,idx){\n";
  h += "    var d=new Date(row.rawDate);\n";
  h += "    var lb=(d.getMonth()+1)+'/'+d.getDate()+'（'+DAYS[d.getDay()]+'）';\n";
  h += "    var tr='<tr><td class=\"dc\">'+lb+'</td>';\n";
  h += "    for(var i=0;i<6;i++){\n";
  h += "      tr+='<td><input type=\"checkbox\" '+(row.enabled[i]?'checked':'')+' onchange=\"upd('+idx+','+i+',this.checked)\"></td>';\n";
  h += "    }\n";
  h += "    tr+='<td><button class=\"btn-sm\" onclick=\"del('+idx+')\">削除</button></td></tr>';\n";
  h += "    tb.innerHTML+=tr;\n";
  h += "  });\n";
  h += "}\n";

  h += "function upd(idx,slot,val){config[idx].enabled[slot]=val;}\n";
  h += "function del(idx){config.splice(idx,1);render();}\n";
  h += "function addDate(){\n";
  h += "  var v=document.getElementById('nd').value;\n";
  h += "  if(!v)return;\n";
  h += "  var raw=v.replace(/-/g,'/');\n";
  h += "  if(config.some(function(c){return c.rawDate===raw;})){alert('すでに追加されています');return;}\n";
  h += "  config.push({rawDate:raw,enabled:[false,false,false,false,false,false]});\n";
  h += "  config.sort(function(a,b){return a.rawDate<b.rawDate?-1:1;});\n";
  h += "  render();\n";
  h += "}\n";
  h += "function save(){\n";
  h += "  google.script.run.withSuccessHandler(function(r){\n";
  h += "    if(r.ok){var p=document.getElementById('spop');p.style.display='flex';setTimeout(function(){p.style.display='none';},2000);}\n";
  h += "    else{var msg=document.getElementById('cmsg');msg.className='msg msg-err';msg.textContent='エラー:'+r.error;msg.style.display='block';}\n";
  h += "  }).saveSlotConfig(config,currentSchool);\n";
  h += "}\n";

  // タブ切り替え
  h += "function showCalendar(){\n";
  h += "  document.getElementById('tab-cfg').style.display='none';\n";
  h += "  document.getElementById('tab-dl').style.display='none';\n";
  h += "  document.getElementById('tab-cal').style.display='block';\n";
  h += "  document.getElementById('btn-cfg').className='tab';\n";
  h += "  document.getElementById('btn-cal').className='tab active';\n";
  h += "  document.getElementById('btn-dl').className='tab';\n";
  h += "  google.script.run.withSuccessHandler(function(r){calData=r;renderCalendar();}).getTeacherCalendar(currentSchool);\n";
  h += "}\n";
  h += "function showConfig(){\n";
  h += "  document.getElementById('tab-cal').style.display='none';\n";
  h += "  document.getElementById('tab-dl').style.display='none';\n";
  h += "  document.getElementById('tab-cfg').style.display='block';\n";
  h += "  document.getElementById('btn-cal').className='tab';\n";
  h += "  document.getElementById('btn-dl').className='tab';\n";
  h += "  document.getElementById('btn-cfg').className='tab active';\n";
  h += "  google.script.run.withSuccessHandler(function(r){config=r.config;render();}).getSlotConfig(currentSchool);\n";
  h += "}\n";
  h += "function showDeadline(){\n";
  h += "  document.getElementById('tab-cal').style.display='none';\n";
  h += "  document.getElementById('tab-cfg').style.display='none';\n";
  h += "  document.getElementById('tab-dl').style.display='block';\n";
  h += "  document.getElementById('btn-cal').className='tab';\n";
  h += "  document.getElementById('btn-cfg').className='tab';\n";
  h += "  document.getElementById('btn-dl').className='tab active';\n";
  h += "}\n";

  // 締め切り読み込み
  h += "function loadDeadlines(){\n";
  h += "  google.script.run.withSuccessHandler(function(r){\n";
  h += "    if(r['神領']){var v=r['神領'];document.getElementById('dl-shinryo').value=v.replace(/\\//g,'-');var p=v.split('/');var d=new Date(parseInt(p[0]),parseInt(p[1])-1,parseInt(p[2]));document.getElementById('dow-s').textContent=v+'（'+['日','月','火','水','木','金','土'][d.getDay()]+'）';}\n";
  h += "    if(r['大手町']){var v=r['大手町'];document.getElementById('dl-otemachi').value=v.replace(/\\//g,'-');var p=v.split('/');var d=new Date(parseInt(p[0]),parseInt(p[1])-1,parseInt(p[2]));document.getElementById('dow-o').textContent=v+'（'+['日','月','火','水','木','金','土'][d.getDay()]+'）';}\n";
  h += "  }).getDeadlines();\n";
  h += "}\n";

  // 締め切り保存
  h += "function showDow(inp,spanId){\n";
  h += "  var v=inp.value;if(!v)return;\n";
  h += "  var p=v.split('-');var d=new Date(parseInt(p[0]),parseInt(p[1])-1,parseInt(p[2]));\n";
  h += "  document.getElementById(spanId).textContent=p[0]+'/'+p[1]+'/'+p[2]+'（'+['日','月','火','水','木','金','土'][d.getDay()]+'）';\n";
  h += "}\n";
  h += "function saveDeadlineUI(school,inputId){\n";
  h += "  var v=document.getElementById(inputId).value;\n";
  h += "  var dateStr=v?v.replace(/-/g,'/'):''; \n";
  h += "  google.script.run.withSuccessHandler(function(r){\n";
  h += "    var m=document.getElementById('dl-msg');\n";
  h += "    if(r.ok){m.className='msg msg-ok';m.textContent=school+'校の締め切りを保存しました。';}\n";
  h += "    else{m.className='msg msg-err';m.textContent='エラー:'+r.error;}\n";
  h += "    m.style.display='block';\n";
  h += "    setTimeout(function(){m.style.display='none';},3000);\n";
  h += "  }).saveDeadline(school,dateStr);\n";
  h += "}\n";

  // カレンダー描画
  h += "function renderCalendar(){\n";
  h += "  var teachers=calData.teachers||{};\n";
  h += "  var allDates=calData.allDates||[];\n";
  h += "  var OBON=['2026/08/10','2026/08/11','2026/08/12','2026/08/13','2026/08/14','2026/08/15','2026/08/16'];\n";
  h += "  var T1=['13:00','14:20','15:45','17:10','18:35','20:00'];\n";
  h += "  var T2=['14:15','15:35','17:00','18:25','19:50','21:15'];\n";
  h += "  var DAYS=['日','月','火','水','木','金','土'];\n";
  h += "  var names=Object.keys(teachers).sort();\n";
  h += "  if(!names.length){document.getElementById('cal-area').innerHTML='<p style=\"color:#999;padding:16px\">まだ登録データがありません</p>';return;}\n";
  h += "  var html='';\n";
  h += "  names.forEach(function(name){\n";
  h += "    var info=teachers[name];\n";
  h += "    var entries=info.entries;\n";
  h += "    html+='<div class=\"tc\" data-name=\"'+name+'\">';\n";
  h += "    html+='<h2>'+currentSchool+'校<br><span style=\"font-size:16pt;font-weight:bold\">'+name+'</span></h2>';\n";
  h += "    html+='<table><thead><tr><th>日付</th>';\n";
  h += "    for(var x=0;x<6;x++)html+='<th>'+T1[x]+'<br>'+T2[x]+'</th>';\n";
  h += "    html+='</tr></thead><tbody>';\n";
  h += "    var obonDone=false;\n";
  h += "    allDates.forEach(function(sl){\n";
  h += "      if(OBON.indexOf(sl.raw)!==-1){\n";
  h += "        if(!obonDone){html+='<tr><td colspan=\"7\" style=\"background:#E3F2FD;color:#0D47A1;font-weight:bold;padding:8px;text-align:center\">🌊 お盆休み（8/10〜8/16）</td></tr>';obonDone=true;}\n";
  h += "        return;\n";
  h += "      }\n";
  h += "      var d=new Date(sl.raw);\n";
  h += "      var lb=(d.getMonth()+1)+'/'+d.getDate()+'（'+DAYS[d.getDay()]+'）';\n";
  h += "      html+='<tr><td class=\"dc\">'+lb+'</td>';\n";
  h += "      for(var i=0;i<6;i++){\n";
  h += "        if(!sl.enabled[i]){html+='<td style=\"color:#ccc\">ー</td>';continue;}\n";
  h += "        var key=sl.raw+'|'+i;\n";
  h += "        var ev=entries[key]||{value:''};\n";
  h += "        var val=ev.value;\n";
  h += "        var cls=val==='○'?'ci':val==='△'?'tr':'';\n";
  h += "        html+='<td><span class=\"dot '+cls+'\">'+val+'</span></td>';\n";
  h += "      }\n";
  h += "      html+='</tr>';\n";
  h += "    });\n";
  h += "    html+='</tbody></table></div>';\n";
  h += "  });\n";
  h += "  document.getElementById('cal-area').innerHTML=html;\n";
  h += "}\n";
  h += "\n";

  h += "var filterSel={};\n";
  h += "function toggleFilter(){\n";
  h += "  var f=document.getElementById('teacher-filter');\n";
  h += "  if(f.style.display==='none'){\n";
  h += "    f.style.display='';\n";
  h += "    var names=Object.keys(calData.teachers||{}).sort();\n";
  h += "    var html='';\n";
  h += "    names.forEach(function(n){\n";
  h += "      var chk=filterSel[n]!==false;\n";
  h += "      html+='<label style=\"display:flex;align-items:center;gap:8px;padding:4px 0;cursor:pointer\"><input type=\"checkbox\"'+( chk?' checked':'')+' data-n=\"'+n+'\">'+ n +'</label>';\n";
  h += "    });\n";
  h += "    document.getElementById('filter-checks').innerHTML=html;\n";
  h += "  }else{f.style.display='none';}\n";
  h += "}\n";
  h += "function filterSelAll(){document.querySelectorAll('#filter-checks input').forEach(function(cb){cb.checked=true;});}\n";
  h += "function filterSelNone(){document.querySelectorAll('#filter-checks input').forEach(function(cb){cb.checked=false;});}\n";
  h += "function applyFilter(){\n";
  h += "  document.querySelectorAll('#filter-checks input').forEach(function(cb){filterSel[cb.dataset.n]=cb.checked;});\n";
  h += "  document.querySelectorAll('#cal-area .tc').forEach(function(el){\n";
  h += "    var n=el.dataset.name;\n";
  h += "    el.style.display=(filterSel[n]===false)?'none':'';\n";
  h += "  });\n";
  h += "  document.getElementById('teacher-filter').style.display='none';\n";
  h += "}\n";
  h += "function confirmTeacher(name){\n";
  h += "  if(!confirm(name+'さんの変更フラグをすべて確認済みにします。よろしいですか？'))return;\n";
  h += "  google.script.run.withSuccessHandler(function(){\n";
  h += "    google.script.run.withSuccessHandler(function(r){calData=r;renderCalendar();}).getTeacherCalendar(currentSchool);\n";
  h += "  }).clearTeacherChanges(name,currentSchool);\n";
  h += "}\n";

  // 一括登録
  h += "function doBulkWrite(){\n";
  h += "  if(!confirm(currentSchool+'校の登録データを全体時間割に一括登録します。\\n\\n既存の担当列を上書きすることはありません。\\n（空き行に追記、なければ行を追加します）\\n\\nよろしいですか？'))return;\n";
  h += "  var btn=document.querySelector('.btn-bulk');\n";
  h += "  btn.disabled=true;btn.textContent='登録中...';\n";
  h += "  google.script.run\n";
  h += "    .withSuccessHandler(function(r){\n";
  h += "      btn.disabled=false;btn.textContent='📋 時間割に一括登録する';\n";
  h += "      var m=document.getElementById('bulk-msg');\n";
  h += "      if(r.ok){\n";
  h += "        var bp=document.getElementById('bpop');\n";
  h += "        document.getElementById('bpop-msg').textContent=currentSchool+'校 一括登録完了！';\n";
  h += "        bp.style.display='flex';\n";
  h += "        setTimeout(function(){bp.style.display='none';},3000);\n";
  h += "        m.className='msg msg-ok';m.textContent='✅ '+currentSchool+'校 一括登録しました！';\n";
  h += "      }\n";
  h += "      else{m.className='msg msg-err';m.textContent='エラー: '+r.error;}\n";
  h += "      m.style.display='block';\n";
  h += "    })\n";
  h += "    .withFailureHandler(function(e){\n";
  h += "      btn.disabled=false;btn.textContent='📋 時間割に一括登録する';\n";
  h += "      var m=document.getElementById('bulk-msg');\n";
  h += "      m.className='msg msg-err';m.textContent='エラー: '+e.message;\n";
  h += "      m.style.display='block';\n";
  h += "    })\n";
  h += "    .bulkWriteToSchedule(currentSchool);\n";
  h += "}\n";

  // 印刷対象選択
  h += "function showPrintSel(){\n";
  h += "  var names=Object.keys(calData.teachers||{}).sort();\n";
  h += "  var html='';\n";
  h += "  names.forEach(function(n){\n";
  h += "    html+='<label style=\"display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid #eee;cursor:pointer\">';\n";
  h += "    html+='<input type=\"checkbox\" checked data-pname=\"'+n+'\">'+n+'</label>';\n";
  h += "  });\n";
  h += "  document.getElementById('pchks').innerHTML=html;\n";
  h += "  document.getElementById('psel').style.display='flex';\n";
  h += "}\n";
  h += "function selAll(){document.querySelectorAll('#pchks input').forEach(function(cb){cb.checked=true;});}\n";
  h += "function closePsel(){document.getElementById('psel').style.display='none';}\n";
  h += "function doPrint(){\n";
  h += "  var selected={};\n";
  h += "  document.querySelectorAll('#pchks input').forEach(function(cb){selected[cb.dataset.pname]=cb.checked;});\n";
  h += "  document.querySelectorAll('#cal-area .tc').forEach(function(el){\n";
  h += "    var n=el.dataset.name;\n";
  h += "    el.classList.toggle('hide-print',!selected[n]);\n";
  h += "  });\n";
  h += "  closePsel();\n";
  h += "  window.print();\n";
  h += "  document.querySelectorAll('#cal-area .tc').forEach(function(el){el.classList.remove('hide-print');});\n";
  h += "}\n";
  h += "</script><style>@keyframes pop{from{transform:scale(.5);opacity:0}to{transform:scale(1);opacity:1}}</style></body></html>";
  return h;
}