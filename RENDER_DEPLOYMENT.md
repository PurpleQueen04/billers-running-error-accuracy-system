# Render + Supabase deployment

## Important separation
- Use the dedicated Billers Supabase project only.
- Do not use the QC Monitoring Supabase connection string.
- Keep any local SQLite source database backed up and separate from production.

## GitHub
1. Create or use a private repository for this application.
2. Commit the Flask source, templates, static assets, requirements, and deployment documentation.
3. Do not commit `.env`, database credentials, production workbooks, or `billers_system.db`.

## Render Web Service
- Runtime: Python
- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app --timeout 180`
- Environment variables:
  - `DATABASE_URL`: Billers Supabase Session Pooler URI (port 5432), stored privately in Render.
  - `SECRET_KEY`: long random secret value.
  - `RENDER`: `true`

The application does **not** automatically execute schema initialization when Gunicorn starts. This is intentional so deployment does not unexpectedly run database migrations. Before using a new version that requires schema compatibility updates:

1. Take a current Supabase backup.
2. Confirm `DATABASE_URL` points to the dedicated Billers project, never the QC Monitoring project.
3. Run the following one-time command from a trusted environment with the new application code and the production `DATABASE_URL` configured:

   ```bash
   flask --app app init-db
   ```

4. Confirm the command completes successfully before restarting or switching traffic to the new application version.

`init-db` creates missing tables and columns and applies guarded compatibility migrations. It does not drop, truncate, or reset application tables. It also safely copies legacy single-project assignments into the multi-project association table when that legacy column exists, and synchronizes response fields from the response-history table only where the corresponding fields are missing or still at the default pending value. Review the migration code and keep the backup before running it.

## Existing production data

- Do not use a SQLite snapshot migration against a live Supabase database unless you intentionally intend to replace the destination data.
- Do not run any migration against the QC Monitoring Supabase project.
- The updated workbook import flow preserves existing Biller/QC response columns when error records are re-imported.
- Admin imports can target one project or use a `PROJECT` column to route rows. QC imports are restricted to the QC account's assigned active projects, with server-side validation for every row.

## Local development

Without `DATABASE_URL`, the application uses SQLite in `billers_system.db`. Initialize the local schema with `flask --app app init-db`, then run `python app.py`.
