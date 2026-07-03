// The signed-in user's profile row (email + role). `role === "admin"` unlocks
// the admin view of the run history. Requires supabase-setup.sql to have been
// run once in the Supabase SQL Editor.

import { useEffect, useState } from "react";
import type { Session } from "@supabase/supabase-js";
import { supabase, isSupabaseConfigured } from "@/lib/supabase";

export type Profile = { id: string; email: string; role: "user" | "admin" };

export function useProfile(session: Session | null) {
  const [profile, setProfile] = useState<Profile | null>(null);

  useEffect(() => {
    if (!isSupabaseConfigured || !session?.user?.id) {
      setProfile(null);
      return;
    }
    let live = true;
    supabase
      .from("profiles")
      .select("id, email, role")
      .eq("id", session.user.id)
      .maybeSingle()
      .then(({ data, error }) => {
        if (!live) return;
        if (error) {
          // Table missing (SQL not run yet) or transient error — stay a normal user.
          console.warn("[profile] load failed:", error.message);
          setProfile(null);
          return;
        }
        setProfile((data as Profile) ?? null);
      });
    return () => {
      live = false;
    };
  }, [session?.user?.id]);

  return profile;
}
