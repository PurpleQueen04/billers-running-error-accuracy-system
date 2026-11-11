# Billers Running Error & Accuracy System

Flask application for Admin and Biller workflows, with SQLite support for local development and Supabase PostgreSQL support for deployment.

## Local run

```bash
pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5001`. Without `DATABASE_URL`, the app uses `billers_system.db`.

## Supabase / Render

See `RENDER_DEPLOYMENT.md` for the separate Supabase project, Render environment variables, GitHub setup, and one-time SQLite-to-PostgreSQL migration steps.

The migration source database is `billers_system.db`. Keep it backed up and do not commit it to GitHub. The migration script copies existing password hashes unchanged and checks all six table counts before committing.

## Security

Set a unique `SECRET_KEY` in Render. Never commit `.env` or database credentials. Change any known/default credentials after deployment.
