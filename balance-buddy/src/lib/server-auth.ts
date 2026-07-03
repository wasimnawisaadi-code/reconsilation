// Server-side auth guard for the AI endpoints.
//
// Every AI server function costs real Gemini money, so in production each call
// must carry the signed-in user's Supabase access token. The token is verified
// against Supabase's /auth/v1/user endpoint (signature, expiry, revocation all
// checked by Supabase itself — no JWT library needed here).
//
// When Supabase env vars are absent (fresh clone, local experiments) the guard
// is a no-op so development stays frictionless — same policy as the UI gate.

type AuthedUser = { id: string; email: string | null };

function supabaseEnv(): { url: string; anon: string } {
  const url =
    (typeof process !== "undefined" && process.env?.VITE_SUPABASE_URL) ||
    import.meta.env?.VITE_SUPABASE_URL ||
    "";
  const anon =
    (typeof process !== "undefined" && process.env?.VITE_SUPABASE_ANON_KEY) ||
    import.meta.env?.VITE_SUPABASE_ANON_KEY ||
    "";
  return { url, anon };
}

/**
 * Verify the caller's access token. Returns the user when valid, null when
 * auth is not configured (open dev mode), and THROWS for missing/invalid
 * tokens — the server function surfaces that as a failed call.
 */
export async function requireUser(accessToken: string | undefined): Promise<AuthedUser | null> {
  const { url, anon } = supabaseEnv();
  if (!url.startsWith("http") || anon.length <= 20) return null;

  if (!accessToken) {
    throw new Error("Sign in required to use the AI engine.");
  }
  const res = await fetch(`${url}/auth/v1/user`, {
    headers: { apikey: anon, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    throw new Error("Your session has expired — please sign in again.");
  }
  const u = (await res.json()) as { id?: string; email?: string };
  if (!u?.id) {
    throw new Error("Your session has expired — please sign in again.");
  }
  return { id: u.id, email: u.email ?? null };
}

/* ── Per-user AI rate limit ──────────────────────────────────────────────
   Best-effort in-memory sliding window (per serverless instance). Stops a
   single account from hammering the Gemini budget; a persistent store can
   replace it when billing lands. */
const hits = new Map<string, number[]>();

export function assertAiRateLimit(key: string, max = 40, windowMs = 60 * 60 * 1000): void {
  const now = Date.now();
  const recent = (hits.get(key) ?? []).filter((t) => now - t < windowMs);
  if (recent.length >= max) {
    throw new Error("AI usage limit reached for this hour — please try again later.");
  }
  recent.push(now);
  hits.set(key, recent);
}
