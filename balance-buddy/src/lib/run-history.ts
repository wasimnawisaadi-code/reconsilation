// Saved reconciliation history — one row per run in public.reconciliation_runs
// (see supabase-setup.sql). Saving is fire-and-forget: a DB hiccup must never
// break a reconciliation that already succeeded in the browser.

import { supabase, isSupabaseConfigured } from "./supabase";

export type RunMode = "single" | "year" | "template" | "organize";

export type RunRecord = {
  id: string;
  user_id: string;
  mode: RunMode;
  ours_files: string[];
  partner_files: string[];
  totals: Record<string, number>;
  match_rate: number | null;
  fx: { oursCcy: string; partnerCcy: string; rate: number; active: boolean } | null;
  created_at: string;
  /** Joined owner email — populated for admins viewing everyone's runs. */
  profiles?: { email: string } | null;
};

export type NewRun = Omit<RunRecord, "id" | "user_id" | "created_at" | "profiles">;

/** Save a finished run under the signed-in user. Silently no-ops when auth
 *  is not configured or the user is signed out. */
export async function saveRun(run: NewRun): Promise<void> {
  if (!isSupabaseConfigured) return;
  const { data } = await supabase.auth.getUser();
  const uid = data?.user?.id;
  if (!uid) return;
  const { error } = await supabase.from("reconciliation_runs").insert({ ...run, user_id: uid });
  if (error) console.warn("[history] save failed:", error.message);
}

/** Recent runs, newest first. RLS scopes the result: normal users get their
 *  own runs, the admin gets everyone's (with the owner's email joined). */
export async function fetchRuns(): Promise<RunRecord[]> {
  if (!isSupabaseConfigured) return [];
  const { data, error } = await supabase
    .from("reconciliation_runs")
    .select("*, profiles(email)")
    .order("created_at", { ascending: false })
    .limit(50);
  if (error) throw new Error(error.message);
  return (data ?? []) as RunRecord[];
}

export async function deleteRun(id: string): Promise<void> {
  if (!isSupabaseConfigured) return;
  const { error } = await supabase.from("reconciliation_runs").delete().eq("id", id);
  if (error) throw new Error(error.message);
}
