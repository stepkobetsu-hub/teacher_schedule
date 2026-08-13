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

管理者画面の Magic Link 認証は講師認証とは分離して維持しています。
