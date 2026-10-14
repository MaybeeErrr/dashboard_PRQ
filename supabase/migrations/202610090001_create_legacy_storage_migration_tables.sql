create table if not exists public.dashboard_state (
  key text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

create table if not exists public.legacy_storage_backups (
  id text primary key,
  migration_version integer not null,
  source text not null,
  source_hash text not null,
  key_count integer not null default 0,
  payload jsonb not null,
  created_at timestamptz not null default now()
);

create table if not exists public.legacy_storage_entries (
  id text primary key,
  backup_id text not null references public.legacy_storage_backups(id) on delete cascade,
  migration_version integer not null,
  storage_area text not null check (storage_area in ('localStorage', 'sessionStorage')),
  source_key text not null,
  product text not null,
  classification text not null,
  source_hash text not null,
  raw_value text,
  parsed_value jsonb,
  migrated_at timestamptz not null default now(),
  unique (storage_area, source_key, source_hash)
);

create table if not exists public.reporting_imports (
  id text primary key,
  backup_id text references public.legacy_storage_backups(id) on delete set null,
  product text not null,
  source_key text not null,
  period_key text,
  source_hash text not null,
  row_count integer not null default 0,
  metadata jsonb not null default '{}'::jsonb,
  imported_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.dashboard_records (
  id text primary key,
  import_id text references public.reporting_imports(id) on delete cascade,
  product text not null,
  dataset_kind text not null,
  period_key text,
  business_key text not null,
  row_index integer,
  payload jsonb not null,
  source_hash text not null,
  imported_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.migration_logs (
  id text primary key,
  migration_version integer not null,
  source_key text not null,
  source_hash text not null,
  product text not null,
  row_count integer not null default 0,
  status text not null check (status in ('success', 'skipped', 'conflict', 'error')),
  error_message text,
  details jsonb not null default '{}'::jsonb,
  migrated_at timestamptz not null default now(),
  unique (migration_version, source_key, source_hash)
);

create index if not exists idx_legacy_storage_entries_product on public.legacy_storage_entries(product);
create index if not exists idx_legacy_storage_entries_source_key on public.legacy_storage_entries(source_key);
create index if not exists idx_reporting_imports_product_period on public.reporting_imports(product, period_key);
create unique index if not exists uniq_reporting_imports_identity
  on public.reporting_imports(product, source_key, coalesce(period_key, ''), source_hash);
create index if not exists idx_dashboard_records_product_period on public.dashboard_records(product, period_key);
create unique index if not exists uniq_dashboard_records_identity
  on public.dashboard_records(product, dataset_kind, coalesce(period_key, ''), business_key, source_hash);
create index if not exists idx_dashboard_records_business_key on public.dashboard_records(business_key);
create index if not exists idx_migration_logs_source_key on public.migration_logs(source_key);

alter table public.dashboard_state enable row level security;
alter table public.legacy_storage_backups enable row level security;
alter table public.legacy_storage_entries enable row level security;
alter table public.reporting_imports enable row level security;
alter table public.dashboard_records enable row level security;
alter table public.migration_logs enable row level security;

do $$
begin
  if not exists (
    select 1 from pg_policies
    where schemaname = 'public' and tablename = 'dashboard_state' and policyname = 'service role manages dashboard state'
  ) then
    create policy "service role manages dashboard state"
      on public.dashboard_state
      for all
      using (auth.role() = 'service_role')
      with check (auth.role() = 'service_role');
  end if;

  if not exists (
    select 1 from pg_policies
    where schemaname = 'public' and tablename = 'legacy_storage_backups' and policyname = 'service role manages legacy backups'
  ) then
    create policy "service role manages legacy backups"
      on public.legacy_storage_backups
      for all
      using (auth.role() = 'service_role')
      with check (auth.role() = 'service_role');
  end if;

  if not exists (
    select 1 from pg_policies
    where schemaname = 'public' and tablename = 'legacy_storage_entries' and policyname = 'service role manages legacy entries'
  ) then
    create policy "service role manages legacy entries"
      on public.legacy_storage_entries
      for all
      using (auth.role() = 'service_role')
      with check (auth.role() = 'service_role');
  end if;

  if not exists (
    select 1 from pg_policies
    where schemaname = 'public' and tablename = 'reporting_imports' and policyname = 'service role manages reporting imports'
  ) then
    create policy "service role manages reporting imports"
      on public.reporting_imports
      for all
      using (auth.role() = 'service_role')
      with check (auth.role() = 'service_role');
  end if;

  if not exists (
    select 1 from pg_policies
    where schemaname = 'public' and tablename = 'dashboard_records' and policyname = 'service role manages dashboard records'
  ) then
    create policy "service role manages dashboard records"
      on public.dashboard_records
      for all
      using (auth.role() = 'service_role')
      with check (auth.role() = 'service_role');
  end if;

  if not exists (
    select 1 from pg_policies
    where schemaname = 'public' and tablename = 'migration_logs' and policyname = 'service role manages migration logs'
  ) then
    create policy "service role manages migration logs"
      on public.migration_logs
      for all
      using (auth.role() = 'service_role')
      with check (auth.role() = 'service_role');
  end if;
end $$;
