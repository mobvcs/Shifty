-- 시프티 정규화 테이블 스키마
-- 필요한 경우 Supabase SQL Editor에서 실행하세요.

create table if not exists employees (
  id text primary key,
  team text not null,
  group_key text not null default '',
  name text not null,
  position text not null,
  employment_type text not null default 'REGULAR',
  hire_date date,
  recontract_date date,
  resignation_date date,
  leave_balance numeric(6, 1) not null default 0,
  sort_order integer not null default 0,
  updated_at timestamptz not null default now()
);

create table if not exists shift_codes (
  code text primary key,
  color text not null,
  updated_at timestamptz not null default now()
);

create table if not exists shift_entries (
  employee_id text not null,
  work_date date not null,
  shift_code text not null,
  updated_at timestamptz not null default now(),
  primary key (employee_id, work_date)
);

create table if not exists shift_events (
  event_id text primary key,
  title text not null,
  start_date date not null,
  end_date date not null,
  bg_color text not null default '#fff3b0',
  display_order integer not null default 0,
  updated_at timestamptz not null default now()
);

create table if not exists occ_daily (
  work_date date primary key,
  room_count integer not null default 0,
  occ_percent numeric(5, 2) not null default 0,
  updated_at timestamptz not null default now()
);

create table if not exists occ_meta (
  month_key text primary key,
  last_occ_updated_text text not null default '-',
  updated_at timestamptz not null default now()
);

create table if not exists cell_colors (
  employee_id text not null,
  work_date date not null,
  bg_color text not null,
  updated_at timestamptz not null default now(),
  primary key (employee_id, work_date)
);

create table if not exists leave_configs (
  employee_id text primary key,
  balance numeric(6, 1) not null default 0,
  apply_seniority boolean not null default true,
  updated_at timestamptz not null default now()
);

create table if not exists month_closings (
  month_key text primary key,
  is_closed boolean not null default true,
  carry_map jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

create table if not exists position_colors (
  position text primary key,
  bg_color text not null,
  updated_at timestamptz not null default now()
);

create table if not exists custom_off_days (
  work_date date primary key,
  updated_at timestamptz not null default now()
);

create index if not exists idx_shift_entries_work_date on shift_entries(work_date);
create index if not exists idx_cell_colors_work_date on cell_colors(work_date);
create index if not exists idx_shift_events_date on shift_events(start_date, end_date);
create index if not exists idx_month_closings_is_closed on month_closings(is_closed);
create index if not exists idx_position_colors_bg on position_colors(bg_color);
create index if not exists idx_occ_meta_updated_at on occ_meta(updated_at);
create index if not exists idx_custom_off_days_updated_at on custom_off_days(updated_at);

alter table employees add column if not exists resignation_date date;
alter table shift_events add column if not exists bg_color text not null default '#fff3b0';

