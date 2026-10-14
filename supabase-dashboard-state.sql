create table if not exists public.dashboard_state (
  key text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);
