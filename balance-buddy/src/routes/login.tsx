// Simple admin password gate — no signup, no forgot-password flow.
// This app has one operator account; entering the password signs in and
// opens the app. Auth still runs on Supabase underneath (secure, session-
// based), just the UI is reduced to a single password field.
//
// If Supabase env vars are missing, a setup card explains exactly what to do
// instead of showing a broken form.

import { createFileRoute, useNavigate } from "@tanstack/react-router";
import React, { useEffect, useState } from "react";
import { Lock, Eye, EyeOff, LogIn, ArrowLeft, AlertCircle, ShieldCheck } from "lucide-react";
import { BrandLogoVector } from "@/components/Brand";
import { supabase, isSupabaseConfigured, friendlyAuthError } from "@/lib/supabase";
import { useSession } from "@/hooks/useSession";

const NAVY = "#0c2e5f";
const GOLD = "#c9a23a";

/** The one account this app gates behind. */
const ADMIN_EMAIL = "admin@nawisaadireconsilation.com";

export const Route = createFileRoute("/login")({
  head: () => ({
    meta: [
      { title: "Sign in | Nawi Saadi Reconciliation" },
      { name: "description", content: "Sign in to Nawi Saadi AI Ledger Reconciliation." },
    ],
  }),
  component: LoginPage,
});

function LoginPage() {
  const navigate = useNavigate();
  const { session, loading: sessionLoading } = useSession();

  const [password, setPassword] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Already signed in → straight to the app.
  useEffect(() => {
    if (!sessionLoading && session) {
      navigate({ to: "/" });
    }
  }, [session, sessionLoading, navigate]);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setBusy(true);
    try {
      const { error: err } = await supabase.auth.signInWithPassword({
        email: ADMIN_EMAIL,
        password,
      });
      if (err) throw err;
      navigate({ to: "/" });
    } catch (err) {
      setError(friendlyAuthError(err instanceof Error ? err.message : String(err)));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div
      className="min-h-screen flex items-center justify-center px-4 py-10"
      style={{ background: `linear-gradient(135deg, #081d3d 0%, ${NAVY} 55%, #103a73 100%)` }}
    >
      <div className="w-full max-w-md">
        <div className="flex justify-center mb-7">
          <BrandLogoVector light />
        </div>

        {!isSupabaseConfigured ? (
          /* ── Setup card: env vars missing ───────────────────────────── */
          <div className="rounded-3xl bg-white shadow-2xl p-7">
            <div className="flex items-center gap-2.5 mb-3">
              <div
                className="size-9 rounded-xl flex items-center justify-center"
                style={{ background: NAVY }}
              >
                <ShieldCheck className="size-5" style={{ color: GOLD }} />
              </div>
              <h1 className="text-lg font-black text-slate-800">Connect Supabase to enable login</h1>
            </div>
            <ol className="list-decimal ml-5 space-y-2 text-[13px] text-slate-600 leading-relaxed">
              <li>
                Create a free project at{" "}
                <span className="font-bold text-slate-800">supabase.com</span>.
              </li>
              <li>
                In the dashboard open <span className="font-bold">Project Settings → API</span> and
                copy the <span className="font-bold">Project URL</span> and the{" "}
                <span className="font-bold">anon public key</span>.
              </li>
              <li>
                In the project folder, copy <code className="rounded bg-slate-100 px-1">.env.example</code>{" "}
                to <code className="rounded bg-slate-100 px-1">.env.local</code> and paste both values.
              </li>
              <li>Restart the dev server (or redeploy) — this page becomes the login form.</li>
            </ol>
            <a
              href="/"
              className="mt-5 inline-flex items-center gap-1.5 text-xs font-bold"
              style={{ color: NAVY }}
            >
              <ArrowLeft className="size-3.5" /> Continue to the app without login
            </a>
          </div>
        ) : (
          /* ── Password gate ──────────────────────────────────────────── */
          <div className="rounded-3xl bg-white shadow-2xl p-7">
            <h1 className="text-xl font-black text-slate-800 text-center">Enter Password</h1>
            <p className="mt-1 mb-5 text-center text-xs text-slate-500">
              Nawi Saadi AI Ledger Reconciliation — admin access only.
            </p>

            {error && (
              <div className="mb-4 flex items-start gap-2 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2.5 text-xs font-semibold text-rose-700">
                <AlertCircle className="size-4 shrink-0 mt-0.5" /> {error}
              </div>
            )}

            <form onSubmit={submit} className="space-y-3.5">
              <label className="block">
                <div className="relative">
                  <Lock className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-slate-400" />
                  <input
                    type={showPw ? "text" : "password"}
                    autoComplete="current-password"
                    autoFocus
                    required
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    placeholder="Password"
                    className="w-full rounded-xl border border-slate-200 bg-slate-50 py-3 pl-10 pr-10 text-sm font-semibold text-slate-800 outline-none transition-colors focus:border-slate-400 focus:bg-white"
                  />
                  <button
                    type="button"
                    onClick={() => setShowPw((s) => !s)}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                    title={showPw ? "Hide password" : "Show password"}
                  >
                    {showPw ? <EyeOff className="size-4" /> : <Eye className="size-4" />}
                  </button>
                </div>
              </label>

              <button
                type="submit"
                disabled={busy}
                className="mt-1 flex w-full items-center justify-center gap-2 rounded-xl py-3 text-sm font-black shadow-lg transition-all hover:opacity-95 active:scale-[0.99] disabled:opacity-50"
                style={{ background: `linear-gradient(90deg, #d4af37, ${GOLD})`, color: NAVY }}
              >
                {busy ? (
                  <span className="size-4 animate-spin rounded-full border-2 border-[#0c2e5f]/30 border-t-[#0c2e5f]" />
                ) : (
                  <LogIn className="size-4" />
                )}
                {busy ? "Opening…" : "Open"}
              </button>
            </form>

            <div className="mt-5 flex items-center justify-center text-[11px] font-bold text-slate-300">
              <ShieldCheck className="size-3.5 mr-1" /> Secured by Supabase Auth
            </div>
          </div>
        )}

        <p className="mt-6 text-center text-[10px] font-medium tracking-wide text-white/40">
          Nawi Saadi Travel &amp; Tourism — AI Ledger Reconciliation
        </p>
      </div>
    </div>
  );
}
