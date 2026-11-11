# Render + Supabase deployment

## Important separation
- Use the dedicated `billers-running-error-system` Supabase project only.
- Do not use the QC Monitoring Supabase connection string.
- Keep `billers_system.db` as the read-only migration source until verification is complete.

## GitHub
1. Create a new private repository for this application.
2. Commit the Flask source, templates, static assets, requirements, and migration script.
3. Do not commit `.env`, database credentials, or `billers_system.db`.

## Render Web Service
- Runtime: Python
- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app`
- Environment variables:
  - `DATABASE_URL`: Supabase Session Pooler URI (port 5432), stored privately in Render.
  - `SECRET_KEY`: long random secret value.
  - `RENDER`: `true`

The Flask app creates the PostgreSQL schema on startup. It does not seed a default Admin account unless both `ADMIN_USERNAME` and `ADMIN_PASSWORD` are explicitly set and the users table is empty. Existing accounts should normally be migrated from SQLite first.

## One-time data migration
1. Keep the original `billers_system.db` on your trusted local computer.
2. Set `DATABASE_URL` in your local terminal/session to the NEW Billers Supabase Session Pooler URI.
3. Install dependencies with `pip install -r requirements.txt`.
4. Run `python migrate_sqlite_to_supabase.py` from this project folder.
5. Confirm the source record counts and type `MIGRATE` when prompted.
6. The script replaces rows in the six application tables in the dedicated Billers project, copies password hashes unchanged, resets identity sequences, and verifies counts before commit.

Do not run the migration against the QC Monitoring project. Do not run it again after the live site has started receiving new data unless you intentionally want to replace the destination records with the SQLite snapshot.

## Local development
Without `DATABASE_URL`, the app continues to use `billers_system.db` with SQLite. This allows local development to continue independently.
