// Supabase client for authentication (email + password).
//
// Configuration comes from Vite env vars — put them in `.env.local` (never
// committed; see `.env.example` in the project root):
//   VITE_SUPABASE_URL=https://YOUR-PROJECT-ref.supabase.co
//   VITE_SUPABASE_ANON_KEY=eyJ....
//
// The anon (public) key is safe to ship to the browser — Supabase is designed
// for that; row-level security and auth rules protect the data server-side.

import { createClient } from "@supabase/supabase-js";

const url = (import.meta.env.VITE_SUPABASE_URL as string | undefined) ?? "";
const anonKey = (import.meta.env.VITE_SUPABASE_ANON_KEY as string | undefined) ?? "";

/** False until the two env vars above are set — the login page then shows
 *  setup instructions instead of a broken form, and the app stays usable. */
export const isSupabaseConfigured = url.startsWith("http") && anonKey.length > 20;

// Placeholder credentials keep module load from throwing when unconfigured;
// every real call is gated behind `isSupabaseConfigured`.
export const supabase = createClient(
  isSupabaseConfigured ? url : "https://placeholder.supabase.co",
  isSupabaseConfigured ? anonKey : "placeholder-anon-key",
);

/** Map Supabase auth errors to friendly, actionable text. */
export function friendlyAuthError(message: string): string {
  const m = message.toLowerCase();
  if (m.includes("invalid login credentials"))
    return "Wrong email or password. Check both and try again.";
  if (m.includes("email not confirmed"))
    return "Your email isn't confirmed yet — open the confirmation link we sent you, then sign in.";
  if (m.includes("user already registered"))
    return "An account with this email already exists — sign in instead.";
  if (m.includes("password should be at least"))
    return "Password is too short — use at least 6 characters.";
  if (m.includes("rate limit") || m.includes("too many requests"))
    return "Too many attempts — wait a minute and try again.";
  if (m.includes("fetch") || m.includes("network"))
    return "Could not reach the login server — check your internet connection.";
  return message;
}
