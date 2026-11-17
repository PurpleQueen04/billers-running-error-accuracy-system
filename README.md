# Billers Running Error & Accuracy System

Flask application for Admin, QC, and Biller workflows, with SQLite support for local development and Supabase PostgreSQL support for deployment.

## Local run

```bash
pip install -r requirements.txt
flask --app app init-db
python app.py
```

Open `http://127.0.0.1:5001`. Without `DATABASE_URL`, the app uses the local `billers_system.db` file.

## QC workbook import

QC accounts can upload CSV, XLS, XLSX, and ODS files from the QC Error Verification page. A QC can select one assigned active project or choose all assigned projects and include a `PROJECT` column in the workbook. The server validates the project on every imported row; rows for unassigned or inactive projects are rejected. Admin can select one project or route rows by the workbook's `PROJECT` column.

## Supabase / Render

See `RENDER_DEPLOYMENT.md` for deployment configuration and safe schema initialization. The application does not automatically run database initialization when Gunicorn starts. Run the `init-db` command deliberately against the intended database after taking a backup.

## Security and data safety

- Set a unique `SECRET_KEY` and the correct Billers `DATABASE_URL` in Render environment variables.
- Never commit `.env`, database credentials, production workbooks, or database files.
- `init-db` creates missing tables/columns and applies guarded one-time compatibility migrations. It does not drop, truncate, or reset existing application tables.
- Do not run a SQLite-to-Supabase data migration against a live database unless you intentionally want to replace destination records with a snapshot.
