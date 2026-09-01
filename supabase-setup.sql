-- Run this once in your Supabase project's SQL Editor (Supabase dashboard > SQL Editor > New query).
-- It creates a single key/value table that the app uses for all its data
-- (work orders, harvest, mileage, expenses, fermentation, barrels).

create table if not exists app_storage (
  key text primary key,
  value text not null,
  shared boolean not null default false,
  updated_at timestamptz not null default now()
);

-- Keep updated_at current on every write
create or replace function set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists app_storage_set_updated_at on app_storage;
create trigger app_storage_set_updated_at
  before update on app_storage
  for each row
  execute function set_updated_at();

-- Row Level Security: this app has no login system, so every visitor with the
-- site's URL shares one open read/write table (the same trust model the
-- Claude-artifact version used). Anyone with the link can view and edit data.
-- If you want to restrict access later (e.g. require a login), this is the
-- policy to tighten first.
alter table app_storage enable row level security;

drop policy if exists "Allow all access" on app_storage;
create policy "Allow all access"
  on app_storage
  for all
  using (true)
  with check (true);
