# X3 Admin Setup

The admin console is available only when the `admin` query parameter is present:

```text
https://testkh.streamlit.app/?admin=1
```

It is intentionally not linked from the customer questionnaire.

## 1. Create the storage

1. Create a Supabase project.
2. Open the SQL editor.
3. Execute the migrations in order:
   - `db/001_crm_admin.sql`
   - `db/002_admin_assets.sql`
   - `db/003_admin_users.sql`
4. Keep Row Level Security enabled. The migration grants no table access to `anon` or `authenticated`.

## 2. Configure Test secrets

Add these values in Streamlit Cloud App settings -> Secrets:

```toml
SUPABASE_URL = "https://PROJECT.supabase.co"
SUPABASE_SERVICE_ROLE_KEY = "SUPABASE_SERVER_SECRET"
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD_HASH = "pbkdf2_sha256$PASSWORD_HASH"
```

Generate the password hash locally:

```powershell
python tools/generate_admin_password.py
```

Never commit `.streamlit/secrets.toml`, Supabase server keys, CRM tokens, or admin passwords.

## 3. Safe activation flow

1. Open the admin console and sign in.
2. Create named administrator/editor/viewer accounts if needed.
3. Choose the customer output format and Telegram delivery options.
4. Download, edit, validate, and publish branding or portfolio assets.
5. Save amoCRM settings and credentials.
6. Run the connection check.
7. Activate amoCRM only after the check succeeds.

The active provider defaults to `off`. CRM failures must not block customer report generation.

## Server migration

The UI reads settings through the storage adapter, so moving to a private server does not require redesigning the admin console. The deployment can keep Supabase or replace the adapter with private PostgreSQL and a server-side secret vault. Only the storage bootstrap credential remains outside the admin console.
