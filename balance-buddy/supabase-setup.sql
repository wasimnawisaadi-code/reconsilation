-- ============================================================================
-- Nawi Saadi Reconciliation — Supabase database setup
-- Run this ONCE in the Supabase Dashboard → SQL Editor → New query → Run.
-- Safe to re-run: everything is IF NOT EXISTS / OR REPLACE / ON CONFLICT.
-- ============================================================================

-- ─────────────────────────────────────────────────────────────────────────
-- 1. PROFILES — one row per user, carries the role (user | admin)
-- ─────────────────────────────────────────────────────────────────────────
create table if not exists public.profiles (
  id         uuid primary key references auth.users(id) on delete cascade,
  email      text not null,
  role       text not null default 'user' check (role in ('user', 'admin')),
  created_at timestamptz not null default now()
);

alter table public.profiles enable row level security;

-- Auto-create a profile whenever someone signs up.
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, email)
  values (new.id, coalesce(new.email, ''))
  on conflict (id) do nothing;
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- Backfill profiles for users that already exist (your admin account).
insert into public.profiles (id, email)
select id, coalesce(email, '') from auth.users
on conflict (id) do nothing;

-- ► Make your account the ADMIN.
update public.profiles
set role = 'admin'
where email = 'admin@nawisaadireconsilation.com';

-- Helper: is the current user an admin? SECURITY DEFINER avoids RLS recursion.
create or replace function public.is_admin()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1 from public.profiles
    where id = auth.uid() and role = 'admin'
  );
$$;

-- Profiles policies: you can read your own profile; admins can read everyone's.
-- Nobody can change roles from the app (only via SQL / service key).
drop policy if exists "profiles_select" on public.profiles;
create policy "profiles_select" on public.profiles
  for select using (auth.uid() = id or public.is_admin());

-- ─────────────────────────────────────────────────────────────────────────
-- 2. RECONCILIATION RUNS — saved history of every reconciliation
-- ─────────────────────────────────────────────────────────────────────────
create table if not exists public.reconciliation_runs (
  id            uuid primary key default gen_random_uuid(),
  -- References profiles (not auth.users directly) so the app can join the
  -- owner's email in one query; profiles cascade from auth.users anyway.
  user_id       uuid not null references public.profiles(id) on delete cascade,
  mode          text not null check (mode in ('single', 'year', 'template', 'organize')),
  ours_files    text[] not null default '{}',
  partner_files text[] not null default '{}',
  totals        jsonb not null default '{}'::jsonb,  -- matched, amountIssues, onlyOurs, onlyPartner…
  match_rate    numeric,                             -- 0..1
  fx            jsonb,                               -- { oursCcy, partnerCcy, rate, active }
  created_at    timestamptz not null default now()
);

create index if not exists reconciliation_runs_user_idx
  on public.reconciliation_runs (user_id, created_at desc);

alter table public.reconciliation_runs enable row level security;

-- You can save runs only as yourself.
drop policy if exists "runs_insert_own" on public.reconciliation_runs;
create policy "runs_insert_own" on public.reconciliation_runs
  for insert with check (auth.uid() = user_id);

-- You can see your own runs; the admin sees EVERY user's runs.
drop policy if exists "runs_select_own_or_admin" on public.reconciliation_runs;
create policy "runs_select_own_or_admin" on public.reconciliation_runs
  for select using (auth.uid() = user_id or public.is_admin());

-- You can delete your own runs; the admin can delete any.
drop policy if exists "runs_delete_own_or_admin" on public.reconciliation_runs;
create policy "runs_delete_own_or_admin" on public.reconciliation_runs
  for delete using (auth.uid() = user_id or public.is_admin());

-- ============================================================================
-- Done. Verify with:
--   select email, role from public.profiles;
-- You should see admin@nawisaadireconsilation.com with role = admin.
-- ============================================================================
