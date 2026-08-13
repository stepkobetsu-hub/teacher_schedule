/**
 * 講師予定・夏休み出勤登録 専用認証API
 * 正式データは「給与明細2026-6-」の「講師マスター」。
 * GitHub Pagesへ生年月日・パスワード・秘密情報を置かない。
 */
const TEACHER_MASTER_SPREADSHEET_ID = '1L5aFDXAmfUDkBg8d7X3WqJgMhdMq5tM5sfUZ2G-M58E';
const TEACHER_MASTER_SHEET_NAME = '講師マスター';
const SHARED_SECRET_PROPERTY = 'EDGE_SHARED_SECRET';

function doPost(e) {
  const output = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);
  try {
    const input = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (input.action !== 'teacherScheduleLogin') {
      console.warn('teacher schedule auth rejected: action');
      return output.setContent(JSON.stringify({ success: false }));
    }
    const expectedSecret = PropertiesService.getScriptProperties().getProperty(SHARED_SECRET_PROPERTY) || '';
    if (!expectedSecret || !constantTimeEqual_(String(input.sharedSecret || ''), expectedSecret)) {
      console.warn('teacher schedule auth rejected: shared secret');
      return output.setContent(JSON.stringify({ success: false }));
    }
    const result = authenticateTeacherSchedule_(input.code, input.password);
    console.info('teacher schedule auth result', { success: result.success === true });
    return output.setContent(JSON.stringify(result));
  } catch (error) {
    console.error('teacher schedule auth error', error);
    return output.setContent(JSON.stringify({ success: false }));
  }
}

function authenticateTeacherSchedule_(rawCode, rawPassword) {
  const code = String(rawCode || '').trim();
  const password = String(rawPassword || '');
  if (!/^\d{4,8}$/.test(code) || !/^\d{4}$/.test(password)) return { success: false };

  const sheet = SpreadsheetApp.openById(TEACHER_MASTER_SPREADSHEET_ID).getSheetByName(TEACHER_MASTER_SHEET_NAME);
  if (!sheet) throw new Error('講師マスターが見つかりません');
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return { success: false };
  const rows = sheet.getRange(5, 1, lastRow - 4, 37).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() !== code) continue;
    const active = String(rows[i][3] || '').trim() === '1';
    const savedPassword = String(rows[i][35] || '').trim();
    if (!active || !savedPassword || !constantTimeEqual_(password, savedPassword)) return { success: false };
    return { success: true, code: code };
  }
  return { success: false };
}

function constantTimeEqual_(left, right) {
  left = String(left || '');
  right = String(right || '');
  let mismatch = left.length ^ right.length;
  const length = Math.max(left.length, right.length);
  for (let i = 0; i < length; i++) mismatch |= (left.charCodeAt(i) || 0) ^ (right.charCodeAt(i) || 0);
  return mismatch === 0;
}
