import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const admin = fs.readFileSync(new URL('../admin.html', import.meta.url), 'utf8');
const edge = fs.readFileSync(new URL('../supabase/functions/admin-login/index.ts', import.meta.url), 'utf8');
const migration = fs.readFileSync(new URL('../supabase/migrations/202608130002_admin_step_session_auth.sql', import.meta.url), 'utf8');
const teacher = fs.readFileSync(new URL('../teacher_app.html', import.meta.url), 'utf8');

test('管理者画面はMagic Link UIとメール送信処理を持たない', () => {
  assert.doesNotMatch(admin, /adminEmail|sendLoginLink|sendAdminLoginLink|\/auth\/v1\/otp|consumeAdminAuthHash/);
  assert.match(admin, /STEPスタッフログインが必要です。/);
  assert.match(admin, /STEP資産管理へ戻る/);
});

test('起動時は認証画面も管理画面も隠し、STEP共通トークンをEdgeへ送る', () => {
  assert.match(admin, /id="authGate" class="authGate hidden"/);
  assert.match(admin, /id="adminApp" class="wrap hidden"/);
  assert.match(admin, /STEP_STAFF_AUTH_KEY = 'stepStaffAppAuth'/);
  assert.match(admin, /systemPortalSessionToken/);
  assert.match(admin, /\/functions\/v1\/admin-login/);
  assert.doesNotMatch(admin, /stepStaffAppPassword/);
});

test('Edge Functionは正式verifySystemPortalだけでスタッフセッションを検証する', () => {
  assert.match(edge, /action: "verifySystemPortal"/);
  assert.match(edge, /systemPortalSessionToken: sessionToken/);
  assert.match(edge, /ADMIN_PERMISSION_LEVELS = new Set\(\["2", "3", "4"\]\)/);
  assert.doesNotMatch(edge, /staffLogin|password/);
});

test('管理者権限は改変不能なapp_metadataへ設定される', () => {
  assert.match(edge, /schedule_admin: true/);
  assert.match(edge, /schedule_staff_code: code/);
  assert.match(migration, /app_metadata' ->> 'schedule_admin'/);
  assert.match(migration, /stepkobetsu@gmail\.com/);
  assert.match(migration, /stepkobetsustaff@gmail\.com/);
});

test('講師番号＋誕生日4桁ログインは維持される', () => {
  assert.match(teacher, /id="code"/);
  assert.match(teacher, /id="password"/);
  assert.match(teacher, /\/functions\/v1\/teacher-login/);
  assert.doesNotMatch(teacher, /ログインリンクを送信/);
});
