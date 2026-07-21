create table if not exists public.admin_users (
    username text primary key check (username ~ '^[a-zA-Z0-9._@-]{3,80}$'),
    display_name text not null,
    password_hash text not null,
    role text not null default 'viewer'
        check (role in ('admin', 'editor', 'viewer')),
    active boolean not null default true,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now(),
    updated_by text not null default 'bootstrap'
);

alter table public.admin_users enable row level security;
revoke all on public.admin_users from anon, authenticated;

create or replace function public.set_admin_user_updated_at()
returns trigger
language plpgsql
as $$
begin
    new.updated_at := now();
    return new;
end;
$$;

drop trigger if exists admin_users_updated_at on public.admin_users;
create trigger admin_users_updated_at
before update on public.admin_users
for each row execute function public.set_admin_user_updated_at();
