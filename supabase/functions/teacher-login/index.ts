import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "npm:@supabase/supabase-js@2.111.0";

const ALLOWED_ORIGIN = "https://stepkobetsu-hub.github.io";
const ALLOWED_SCHOOLS = new Set(["神領", "大手町"]);
const GENERIC_AUTH_ERROR = "講師番号またはパスワードが違います。";

function json(body: Record<string, unknown>, status = 200, origin = "") {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": origin === ALLOWED_ORIGIN ? origin : ALLOWED_ORIGIN,
      "Access-Control-Allow-Headers": "apikey, content-type",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Vary": "Origin",
    },
  });
}

async function sha256(value: string) {
  const bytes = new TextEncoder().encode(value);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

Deno.serve(async (req) => {
  const origin = req.headers.get("origin") || "";
  if (req.method === "OPTIONS") return json({ ok: true }, 200, origin);
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405, origin);
  if (origin && origin !== ALLOWED_ORIGIN) return json({ error: "Origin not allowed" }, 403, origin);

  const supabaseUrl = Deno.env.get("SUPABASE_URL") || "";
  const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") || "";
  const anonKey = Deno.env.get("SUPABASE_ANON_KEY") || "";
  if (!supabaseUrl || !serviceRoleKey || !anonKey) return json({ error: "認証サーバーの設定エラーです。" }, 500, origin);

  let input: { school?: unknown; code?: unknown; password?: unknown };
  try { input = await req.json(); } catch { return json({ error: GENERIC_AUTH_ERROR }, 400, origin); }
  const school = String(input.school || "").trim();
  const code = String(input.code || "").trim();
  const password = String(input.password || "");
  if (!ALLOWED_SCHOOLS.has(school) || !/^\d{4,8}$/.test(code) || !/^\d{4}$/.test(password)) {
    return json({ error: GENERIC_AUTH_ERROR }, 401, origin);
  }

  const admin = createClient(supabaseUrl, serviceRoleKey, {
    auth: { autoRefreshToken: false, persistSession: false, detectSessionInUrl: false },
  });
  const forwarded = req.headers.get("x-forwarded-for")?.split(",")[0]?.trim() || "unknown";
  const pairKey = "pair:" + await sha256(`${forwarded}|${code}`);
  const ipKey = "ip:" + await sha256(forwarded);

  const [{ data: pairAllowed }, { data: ipAllowed }] = await Promise.all([
    admin.rpc("schedule_login_rate_allowed", { p_key: pairKey }),
    admin.rpc("schedule_login_rate_allowed", { p_key: ipKey }),
  ]);
  if (pairAllowed !== true || ipAllowed !== true) return json({ error: "試行回数が多すぎます。15分後にお試しください。" }, 429, origin);

  const { data: teacherRows, error: teacherError } = await admin
    .from("teachers").select("code,name,active").eq("code", code).eq("active", true).limit(1);
  if (teacherError) return json({ error: "認証サーバーへ接続できません。" }, 503, origin);

  const { data: authSettings, error: secretError } = await admin.rpc("schedule_auth_settings");
  const sharedSecret = String(authSettings?.shared_secret || "");
  const authApiUrl = String(authSettings?.api_url || "");
  if (secretError || !sharedSecret || !authApiUrl) return json({ error: "認証サーバーの設定エラーです。" }, 500, origin);

  let verified = false;
  if (teacherRows?.length) {
    try {
      let authResponse = await fetch(authApiUrl, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "teacherScheduleLogin", code, password, sharedSecret }),
        redirect: "manual",
      });
      if ([301, 302, 303, 307, 308].includes(authResponse.status)) {
        const location = authResponse.headers.get("location");
        if (location) authResponse = await fetch(new URL(location, authApiUrl), { headers: { Accept: "application/json" } });
      }
      const authText = await authResponse.text();
      let authResult: { success?: boolean; code?: unknown } = {};
      try { authResult = JSON.parse(authText); } catch { authResult = {}; }
      if (!authResponse.ok || authResult?.success !== true) {
        console.error("teacher auth upstream rejected", {
          status: authResponse.status,
          responseLength: authText.length,
          success: authResult?.success === true,
          codeMatches: String(authResult?.code) === code,
        });
      }
      verified = authResponse.ok && authResult?.success === true && String(authResult.code) === code;
    } catch (error) {
      console.error("teacher auth upstream request failed", error instanceof Error ? error.message : "unknown error");
      verified = false;
    }
  }

  if (!verified) {
    await Promise.all([
      admin.rpc("schedule_login_rate_record", { p_key: pairKey, p_limit: 5, p_success: false }),
      admin.rpc("schedule_login_rate_record", { p_key: ipKey, p_limit: 30, p_success: false }),
    ]);
    return json({ error: GENERIC_AUTH_ERROR }, 401, origin);
  }

  await Promise.all([
    admin.rpc("schedule_login_rate_record", { p_key: pairKey, p_limit: 5, p_success: true }),
    admin.rpc("schedule_login_rate_record", { p_key: ipKey, p_limit: 30, p_success: true }),
  ]);

  const syntheticEmail = `teacher-schedule-${code}@auth.invalid`;
  const { data: usersData, error: usersError } = await admin.auth.admin.listUsers({ page: 1, perPage: 1000 });
  if (usersError) return json({ error: "安全なセッションを発行できません。" }, 500, origin);
  let authUser = usersData.users.find((user) => user.email === syntheticEmail);
  if (!authUser) {
    const created = await admin.auth.admin.createUser({
      email: syntheticEmail,
      email_confirm: true,
      app_metadata: { schedule_teacher: true, schedule_teacher_code: code },
    });
    if (created.error || !created.data.user) return json({ error: "安全なセッションを発行できません。" }, 500, origin);
    authUser = created.data.user;
  } else if (authUser.app_metadata?.schedule_teacher_code !== code || authUser.app_metadata?.schedule_teacher !== true) {
    const updated = await admin.auth.admin.updateUserById(authUser.id, {
      app_metadata: { ...authUser.app_metadata, schedule_teacher: true, schedule_teacher_code: code },
    });
    if (updated.error) return json({ error: "安全なセッションを発行できません。" }, 500, origin);
  }

  const link = await admin.auth.admin.generateLink({ type: "magiclink", email: syntheticEmail });
  const tokenHash = link.data?.properties?.hashed_token;
  if (link.error || !tokenHash) return json({ error: "安全なセッションを発行できません。" }, 500, origin);
  const publicClient = createClient(supabaseUrl, anonKey, {
    auth: { autoRefreshToken: false, persistSession: false, detectSessionInUrl: false },
  });
  const verifiedOtp = await publicClient.auth.verifyOtp({ token_hash: tokenHash, type: "email" });
  if (verifiedOtp.error || !verifiedOtp.data.session) return json({ error: "安全なセッションを発行できません。" }, 500, origin);

  const session = verifiedOtp.data.session;
  return json({
    session: {
      access_token: session.access_token,
      refresh_token: session.refresh_token,
      expires_at: session.expires_at,
    },
  }, 200, origin);
});
