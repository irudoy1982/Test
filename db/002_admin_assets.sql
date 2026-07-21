insert into storage.buckets (id, name, public, file_size_limit)
values ('audit-admin-assets', 'audit-admin-assets', false, 26214400)
on conflict (id) do update set
    public = false,
    file_size_limit = excluded.file_size_limit;

create table if not exists public.admin_assets (
    asset_key text primary key check (
        asset_key in ('logo', 'presentation_template', 'vendor_matrix')
    ),
    object_path text not null unique,
    filename text not null,
    content_type text not null,
    size_bytes bigint not null check (size_bytes > 0),
    sha256 text not null,
    details jsonb not null default '{}'::jsonb,
    updated_at timestamptz not null default now(),
    updated_by text not null default 'admin'
);

alter table public.admin_assets enable row level security;
revoke all on public.admin_assets from anon, authenticated;

create or replace function public.set_admin_asset_updated_at()
returns trigger
language plpgsql
as $$
begin
    new.updated_at := now();
    return new;
end;
$$;

drop trigger if exists admin_assets_updated_at on public.admin_assets;
create trigger admin_assets_updated_at
before update on public.admin_assets
for each row execute function public.set_admin_asset_updated_at();
