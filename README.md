# teacher_schedule

講師予定・夏休み出勤登録の GitHub Pages 版です。

- 講師画面: `teacher_app.html`
- 管理者画面: `admin.html`
- 公開URL: <https://stepkobetsu-hub.github.io/teacher_schedule/teacher_app.html>

## 講師認証

講師は校舎・講師番号・誕生日4桁でログインします。誕生日はブラウザや
Supabase の講師テーブルへ保存せず、専用 Apps Script API が正式な
「給与明細2026-6- / 講師マスター」をサーバー側で照合します。

認証成功後は `teacher-login` Supabase Edge Function が講師コードを
`app_metadata` に固定した Supabase Auth セッションを発行します。
`teachers` と `availability` の RLS はこの改変不能な講師コードを使い、
本人のレコードだけを読み書き可能にします。

認証APIのApps ScriptプロジェクトID:
`1woGgT3Xm5MmPAjFyL9Oz2BvcBUN1Yxb5dtcbsibjROlHeU-Z67xY9oLO`

共有秘密は Apps Script の Script Properties と Supabase の private schema
だけに置き、GitHub Pages・リポジトリには保存しません。

## デプロイ対象

- `supabase/migrations/202608130001_teacher_code_auth.sql`
- `supabase/functions/teacher-login/index.ts`
- `teacher_auth_gas.gs`

## 管理者認証

管理者画面は同一 GitHub Pages origin に保存された `stepStaffAppAuth` の
`systemPortalSessionToken` を使用します。`admin-login` Edge Function が
STEPスタッフ共通認証APIの `verifySystemPortal` でトークンと最新権限を
サーバー側検証し、`app_metadata.schedule_admin` を持つSupabase Auth
セッションをメール送信なしで発行します。

管理者Magic Linkは使用しません。従来の許可メール判定はRLSの互換経路として
残し、STEP共通認証から発行された管理者セッションも同じRLSで許可します。
