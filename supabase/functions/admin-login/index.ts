import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "npm:@supabase/supabase-js@2.111.0";

const ALLOWED_ORIGIN = "https://stepkobetsu-hub.github.io";
const STEP_AUTH_API = "https://script.google.com/macros/s/AKfycbypkUc0MqZ07E7pZRglNPeRM56WbCcuWaLpRzi9bVFcPklHDxaaLC7GfzG6ozTGCbEX/exec";
const ADMIN_PERMISSION_LEVELS = new Set(["2", "3", "4"]);
const GENERIC_AUTH_ERROR = "STEPスタッフログインを確認できません。STEP資産管理へ戻って再ログインしてください。";

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
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(value));
  return Array.from(new Uint8Array(digest)).map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

async function fetchStepVerification(sessionToken: string) {
  let response = await fetch(STEP_AUTH_API, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "verifySystemPortal", systemPortalSessionToken: sessionToken }),
    redirect: "manual",
  });
  if ([301, 302, 303, 307, 308].includes(response.status)) {
    const location = response.headers.get("location");
    if (location) response = await fetch(new URL(location, STEP_AUTH_API), { headers: { Accept: "application/json" } });
  }
  if (!response.ok) return null;
  const text = await response.text();
  try { return JSON.parse(text); } catch { return null; }
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

  let input: { systemPortalSessionToken?: unknown };
  try { input = await req.json(); } catch { return json({ error: GENERIC_AUTH_ERROR }, 400, origin); }
  const sessionToken = String(input.systemPortalSessionToken || "").trim();
  if (sessionToken.length < 32 || sessionToken.length > 256) return json({ error: GENERIC_AUTH_ERROR }, 401, origin);

  const admin = createClient(supabaseUrl, serviceRoleKey, {
    auth: { autoRefreshToken: false, persistSession: false, detectSessionInUrl: false },
  });
  const forwarded = req.headers.get("x-forwarded-for")?.split(",")[0]?.trim() || "unknown";
  const pairKey = "admin-pair:" + await sha256(`${forwarded}|${sessionToken}`);
  const ipKey = "admin-ip:" + await sha256(forwarded);
  const [{ data: pairAllowed }, { data: ipAllowed }] = await Promise.all([
    admin.rpc("schedule_login_rate_allowed", { p_key: pairKey }),
    admin.rpc("schedule_login_rate_allowed", { p_key: ipKey }),
  ]);
  if (pairAllowed !== true || ipAllowed !== true) return json({ error: "試行回数が多すぎます。15分後に再度お試しください。" }, 429, origin);

  let verified: { success?: boolean; code?: unknown; permissionLevel?: unknown } | null = null;
  try { verified = await fetchStepVerification(sessionToken); } catch { verified = null; }
  const code = String(verified?.code || "").trim();
  const permissionLevel = String(verified?.permissionLevel || "").trim();
  if (verified?.success !== true || !/^\d{4,8}$/.test(code) || !ADMIN_PERMISSION_LEVELS.has(permissionLevel)) {
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

  const syntheticEmail = `schedule-admin-${code}@auth.invalid`;
  const { data: usersData, error: usersError } = await admin.auth.admin.listUsers({ page: 1, perPage: 1000 });
  if (usersError) return json({ error: "安全な管理者セッションを発行できません。" }, 500, origin);
  let authUser = usersData.users.find((user) => user.email === syntheticEmail);
  const appMetadata = { schedule_admin: true, schedule_staff_code: code, schedule_permission_level: permissionLevel };
  if (!authUser) {
    const created = await admin.auth.admin.createUser({ email: syntheticEmail, email_confirm: true, app_metadata: appMetadata });
    if (created.error || !created.data.user) return json({ error: "安全な管理者セッションを発行できません。" }, 500, origin);
    authUser = created.data.user;
  } else {
    const current = authUser.app_metadata || {};
    if (current.schedule_admin !== true || current.schedule_staff_code !== code || String(current.schedule_permission_level) !== permissionLevel) {
      const updated = await admin.auth.admin.updateUserById(authUser.id, { app_metadata: { ...current, ...appMetadata } });
      if (updated.error) return json({ error: "安全な管理者セッションを発行できません。" }, 500, origin);
    }
  }

  const link = await admin.auth.admin.generateLink({ type: "magiclink", email: syntheticEmail });
  const tokenHash = link.data?.properties?.hashed_token;
  if (link.error || !tokenHash) return json({ error: "安全な管理者セッションを発行できません。" }, 500, origin);
  const publicClient = createClient(supabaseUrl, anonKey, {
    auth: { autoRefreshToken: false, persistSession: false, detectSessionInUrl: false },
  });
  const verifiedOtp = await publicClient.auth.verifyOtp({ token_hash: tokenHash, type: "email" });
  if (verifiedOtp.error || !verifiedOtp.data.session) return json({ error: "安全な管理者セッションを発行できません。" }, 500, origin);

  const session = verifiedOtp.data.session;
  return json({
    session: {
      access_token: session.access_token,
      refresh_token: session.refresh_token,
      expires_at: session.expires_at,
    },
  }, 200, origin);
});
