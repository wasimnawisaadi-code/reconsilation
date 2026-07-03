// Login / Sign-up page — Supabase email + password authentication.
//
// Views:
//   signin  — email + password → session → redirect to the app
//   signup  — create account (Supabase sends a confirmation email by default)
//   forgot  — sends a password-reset email
//   reset   — shown automatically when the user arrives from the reset link
//             (Supabase fires PASSWORD_RECOVERY) → set a new password
//
// If Supabase env vars are missing, a setup card explains exactly what to do
// instead of showing a broken form.

import { createFileRoute, useNavigate } from "@tanstack/react-router";
import React, { useEffect, useState } from "react";
import {
  Mail,
  Lock,
  Eye,
  EyeOff,
  LogIn,
  UserPlus,
  KeyRound,
  ArrowLeft,
  CheckCircle2,
  AlertCircle,
  ShieldCheck,
} from "lucide-react";
import { BrandLogoVector } from "@/components/Brand";
import { supabase, isSupabaseConfigured, friendlyAuthError } from "@/lib/supabase";
import { useSession } from "@/hooks/useSession";

const NAVY = "#0c2e5f";
const GOLD = "#c9a23a";

export const Route = createFileRoute("/login")({
  head: () => ({
    meta: [
      { title: "Sign in | Nawi Saadi Reconciliation" },
      { name: "description", content: "Sign in to Nawi Saadi AI Ledger Reconciliation." },
    ],
  }),
  component: LoginPage,
});

type View = "signin" | "signup" | "forgot" | "reset";

function LoginPage() {
  const navigate = useNavigate();
  const { session, loading: sessionLoading } = useSession();

  const [view, setView] = useState<View>("signin");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);

  // Arriving from the password-reset email: Supabase parses the link hash and
  // fires PASSWORD_RECOVERY — switch to the "set a new password" form.
  useEffect(() => {
    if (!isSupabaseConfigured) return;
    const { data: sub } = supabase.auth.onAuthStateChange((event) => {
      if (event === "PASSWORD_RECOVERY") {
        setView("reset");
        setNotice("Set your new password below.");
        setError(null);
      }
    });
    return () => sub.subscription.unsubscribe();
  }, []);

  // Already signed in (and not mid password-reset) → straight to the app.
  useEffect(() => {
    if (!sessionLoading && session && view !== "reset") {
      navigate({ to: "/" });
    }
  }, [session, sessionLoading, view, navigate]);

  const switchView = (v: View) => {
    setView(v);
    setError(null);
    setNotice(null);
    setConfirm("");
  };

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setNotice(null);

    const mail = email.trim().toLowerCase();
    if (view !== "reset" && !/^\S+@\S+\.\S+$/.test(mail)) {
      setError("Enter a valid email address.");
      return;
    }
    if (view !== "forgot" && password.length < 6) {
      setError("Password must be at least 6 characters.");
      return;
    }
    if ((view === "signup" || view === "reset") && password !== confirm) {
      setError("Passwords don't match — type the same password in both boxes.");
      return;
    }

    setBusy(true);
    try {
      if (view === "signin") {
        const { error: err } = await supabase.auth.signInWithPassword({
          email: mail,
          password,
        });
        if (err) throw err;
        navigate({ to: "/" });
      } else if (view === "signup") {
        const { data, error: err } = await supabase.auth.signUp({
          email: mail,
          password,
          options: { emailRedirectTo: `${window.location.origin}/login` },
        });
        if (err) throw err;
        if (data.session) {
          // Email confirmation is disabled in this project — signed in directly.
          navigate({ to: "/" });
        } else {
          switchView("signin");
          setNotice(
            "Account created! Check your email and click the confirmation link, then sign in.",
          );
        }
      } else if (view === "forgot") {
        const { error: err } = await supabase.auth.resetPasswordForEmail(mail, {
          redirectTo: `${window.location.origin}/login`,
        });
        if (err) throw err;
        setNotice("Reset link sent — check your email and open the link on this device.");
      } else if (view === "reset") {
        const { error: err } = await supabase.auth.updateUser({ password });
        if (err) throw err;
        setNotice("Password updated — taking you to the app…");
        setTimeout(() => navigate({ to: "/" }), 900);
      }
    } catch (err) {
      setError(friendlyAuthError(err instanceof Error ? err.message : String(err)));
    } finally {
      setBusy(false);
    }
  };

  const title =
    view === "signin"
      ? "Welcome back"
      : view === "signup"
        ? "Create your account"
        : view === "forgot"
          ? "Reset your password"
          : "Set a new password";

  const cta =
    view === "signin"
      ? "Sign In"
      : view === "signup"
        ? "Create Account"
        : view === "forgot"
          ? "Send Reset Link"
          : "Save New Password";

  const CtaIcon =
    view === "signin" ? LogIn : view === "signup" ? UserPlus : view === "forgot" ? Mail : KeyRound;

  return (
    <div
      className="min-h-screen flex items-center justify-center px-4 py-10"
      style={{ background: `linear-gradient(135deg, #081d3d 0%, ${NAVY} 55%, #103a73 100%)` }}
    >
      <div className="w-full max-w-md">
        {/* Brand */}
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
          /* ── Auth card ──────────────────────────────────────────────── */
          <div className="rounded-3xl bg-white shadow-2xl p-7">
            <h1 className="text-xl font-black text-slate-800 text-center">{title}</h1>
            <p className="mt-1 mb-5 text-center text-xs text-slate-500">
              {view === "signin" && "Sign in to run AI ledger reconciliation."}
              {view === "signup" && "Free account — email and password is all you need."}
              {view === "forgot" && "We'll email you a link to choose a new password."}
              {view === "reset" && "Choose a strong password you haven't used before."}
            </p>

            {/* Sign in / Sign up tab switch */}
            {(view === "signin" || view === "signup") && (
              <div className="mb-5 grid grid-cols-2 rounded-xl bg-slate-100 p-1 text-xs font-bold">
                {(["signin", "signup"] as const).map((v) => (
                  <button
                    key={v}
                    type="button"
                    onClick={() => switchView(v)}
                    className={`rounded-lg py-2 transition-all ${view === v ? "text-white shadow" : "text-slate-500 hover:text-slate-700"}`}
                    style={view === v ? { background: NAVY } : undefined}
                  >
                    {v === "signin" ? "Sign In" : "Create Account"}
                  </button>
                ))}
              </div>
            )}

            {error && (
              <div className="mb-4 flex items-start gap-2 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2.5 text-xs font-semibold text-rose-700">
                <AlertCircle className="size-4 shrink-0 mt-0.5" /> {error}
              </div>
            )}
            {notice && (
              <div className="mb-4 flex items-start gap-2 rounded-xl border border-emerald-200 bg-emerald-50 px-3 py-2.5 text-xs font-semibold text-emerald-700">
                <CheckCircle2 className="size-4 shrink-0 mt-0.5" /> {notice}
              </div>
            )}

            <form onSubmit={submit} className="space-y-3.5">
              {view !== "reset" && (
                <label className="block">
                  <span className="mb-1 block text-[11px] font-bold uppercase tracking-wider text-slate-500">
                    Email
                  </span>
                  <div className="relative">
                    <Mail className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-slate-400" />
                    <input
                      type="email"
                      autoComplete="email"
                      required
                      value={email}
                      onChange={(e) => setEmail(e.target.value)}
                      placeholder="you@company.com"
                      className="w-full rounded-xl border border-slate-200 bg-slate-50 py-2.5 pl-10 pr-3 text-sm font-semibold text-slate-800 outline-none transition-colors focus:border-slate-400 focus:bg-white"
                    />
                  </div>
                </label>
              )}

              {view !== "forgot" && (
                <label className="block">
                  <span className="mb-1 block text-[11px] font-bold uppercase tracking-wider text-slate-500">
                    {view === "reset" ? "New password" : "Password"}
                  </span>
                  <div className="relative">
                    <Lock className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-slate-400" />
                    <input
                      type={showPw ? "text" : "password"}
                      autoComplete={view === "signin" ? "current-password" : "new-password"}
                      required
                      minLength={6}
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      placeholder="At least 6 characters"
                      className="w-full rounded-xl border border-slate-200 bg-slate-50 py-2.5 pl-10 pr-10 text-sm font-semibold text-slate-800 outline-none transition-colors focus:border-slate-400 focus:bg-white"
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
              )}

              {(view === "signup" || view === "reset") && (
                <label className="block">
                  <span className="mb-1 block text-[11px] font-bold uppercase tracking-wider text-slate-500">
                    Confirm password
                  </span>
                  <div className="relative">
                    <Lock className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-slate-400" />
                    <input
                      type={showPw ? "text" : "password"}
                      autoComplete="new-password"
                      required
                      minLength={6}
                      value={confirm}
                      onChange={(e) => setConfirm(e.target.value)}
                      placeholder="Same password again"
                      className="w-full rounded-xl border border-slate-200 bg-slate-50 py-2.5 pl-10 pr-3 text-sm font-semibold text-slate-800 outline-none transition-colors focus:border-slate-400 focus:bg-white"
                    />
                  </div>
                </label>
              )}

              <button
                type="submit"
                disabled={busy}
                className="mt-1 flex w-full items-center justify-center gap-2 rounded-xl py-3 text-sm font-black shadow-lg transition-all hover:opacity-95 active:scale-[0.99] disabled:opacity-50"
                style={{ background: `linear-gradient(90deg, #d4af37, ${GOLD})`, color: NAVY }}
              >
                {busy ? (
                  <span className="size-4 animate-spin rounded-full border-2 border-[#0c2e5f]/30 border-t-[#0c2e5f]" />
                ) : (
                  <CtaIcon className="size-4" />
                )}
                {busy ? "Please wait…" : cta}
              </button>
            </form>

            {/* Secondary links */}
            <div className="mt-5 flex items-center justify-between text-[11px] font-bold">
              {view === "signin" && (
                <button
                  type="button"
                  onClick={() => switchView("forgot")}
                  className="text-slate-400 transition-colors hover:text-slate-600"
                >
                  Forgot password?
                </button>
              )}
              {(view === "forgot" || view === "reset") && (
                <button
                  type="button"
                  onClick={() => switchView("signin")}
                  className="inline-flex items-center gap-1 text-slate-400 transition-colors hover:text-slate-600"
                >
                  <ArrowLeft className="size-3" /> Back to sign in
                </button>
              )}
              <span className="ml-auto inline-flex items-center gap-1 text-slate-300">
                <ShieldCheck className="size-3.5" /> Secured by Supabase Auth
              </span>
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
