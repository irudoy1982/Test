create table if not exists public.admin_sessions (
    token_hash text primary key,
    username text not null,
    role text not null
        check (role in ('admin', 'editor', 'viewer')),
    display_name text not null,
    expires_at timestamptz not null,
    created_at timestamptz not null default now(),
    last_seen_at timestamptz not null default now()
);

create index if not exists admin_sessions_username_idx
    on public.admin_sessions (username);

create index if not exists admin_sessions_expires_at_idx
    on public.admin_sessions (expires_at);

alter table public.admin_sessions enable row level security;
revoke all on public.admin_sessions from anon, authenticated;
