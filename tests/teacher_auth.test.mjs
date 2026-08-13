import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const teacher = readFileSync(new URL('../teacher_app.html', import.meta.url), 'utf8');
const admin = readFileSync(new URL('../admin.html', import.meta.url), 'utf8');
const migration = readFileSync(new URL('../supabase/migrations/202608130001_teacher_code_auth.sql', import.meta.url), 'utf8');
const edge = readFileSync(new URL('../supabase/functions/teacher-login/index.ts', import.meta.url), 'utf8');
const gas = readFileSync(new URL('../teacher_auth_gas.gs', import.meta.url), 'utf8');

test('講師画面は講師番号と誕生日4桁でログインする', () => {
  assert.match(teacher, /id="code"/);
  assert.match(teacher, /id="password"/);
  assert.match(teacher, /誕生日4桁（例：4月15日 → 0415）/);
  assert.doesNotMatch(teacher, /ログインリンクを送信/);
  assert.doesNotMatch(teacher, /id="email"/);
});

test('講師番号と誕生日は端末へ保存しない', () => {
  assert.doesNotMatch(teacher, /localStorage\.setItem\([^\n]*(?:code|password)/i);
  assert.match(teacher, /step_teacher_schedule_session/);
  assert.match(teacher, /refresh_token/);
});

test('RLS所有者判定は改変不能なapp_metadataの講師コードを使う', () => {
  assert.match(migration, /app_metadata[^\n]+schedule_teacher_code/);
  assert.doesNotMatch(migration, /user_metadata/);
  assert.match(migration, /code = private\.schedule_teacher_code\(\)/);
  assert.match(migration, /code = p_teacher_code[\s\S]+code = private\.schedule_teacher_code\(\)/);
});

test('Edge Functionは正式認証API・試行回数制限・Supabase Authセッションを使う', () => {
  assert.match(edge, /schedule_auth_settings/);
  assert.match(edge, /schedule_login_rate_allowed/);
  assert.match(edge, /app_metadata: \{ schedule_teacher: true, schedule_teacher_code: code \}/);
  assert.match(edge, /generateLink/);
  assert.match(edge, /verifyOtp/);
});

test('認証APIは正式講師マスターをサーバー側参照し共有秘密を必須にする', () => {
  assert.match(gas, /TEACHER_MASTER_SPREADSHEET_ID/);
  assert.match(gas, /EDGE_SHARED_SECRET/);
  assert.match(gas, /rows\[i\]\[3\]/);
  assert.match(gas, /rows\[i\]\[35\]/);
  assert.doesNotMatch(gas, /7002.*0415/);
});

test('管理者Magic Linkと許可メールは変更しない', () => {
  assert.match(admin, /ADMIN_EMAILS = \['stepkobetsu@gmail\.com','stepkobetsustaff@gmail\.com'\]/);
  assert.match(admin, /sendAdminLoginLink/);
  assert.match(admin, /redirectTo=location\.origin\+location\.pathname/);
});
