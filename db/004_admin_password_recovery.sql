alter table public.admin_users
    add column if not exists email text;

create unique index if not exists admin_users_email_unique
    on public.admin_users (lower(email))
    where email is not null and email <> '';

create table if not exists public.admin_password_resets (
    username text primary key,
    code_hash text not null,
    expires_at timestamptz not null,
    attempts integer not null default 0 check (attempts >= 0),
    used boolean not null default false,
    created_at timestamptz not null default now()
);

alter table public.admin_password_resets enable row level security;
revoke all on public.admin_password_resets from anon, authenticated;

create index if not exists admin_password_resets_expiry_idx
    on public.admin_password_resets (expires_at);
