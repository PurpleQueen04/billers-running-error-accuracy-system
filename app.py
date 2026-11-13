
import os
import io
import json
import hashlib
import sqlite3
import secrets
from datetime import datetime, date
from functools import wraps

import pandas as pd

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_file, abort
)

from werkzeug.security import (
    generate_password_hash,
    check_password_hash,
)

from werkzeug.utils import secure_filename

try:
    import psycopg
    from psycopg.rows import dict_row
except ImportError:
    psycopg = None
    dict_row = None


# =========================================================
# APPLICATION CONFIGURATION
# =========================================================

APP_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(APP_DIR, "billers_system.db")

UPLOAD_DIR = os.path.join(APP_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)

app.secret_key = (
    os.environ.get("SECRET_KEY")
    or secrets.token_hex(32)
)

app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=(
        os.environ.get("RENDER", "").lower() == "true"
    ),
    MAX_CONTENT_LENGTH=50 * 1024 * 1024,
)

ALLOWED = {".ods", ".xlsx", ".xls", ".csv"}

ERROR_ALIASES = {
    "date": ["BILLED DATE"],
    "source": ["SOURCE OF ERROR"],
    "audit_type": ["AUDIT TYPE"],
    "name": ["NAME"],
    "user": ["DDC USER ID", "DDC USER"],
    "id": ["ID#"],
    "auditor": ["AUDITOR"],
    "pro": ["PRO#"],
    "zb": ["ZB ERROR"],
    "details": ["ERROR DETAILS"],
    "project": ["PROJECT"],
}

INTERNAL_ACCURACY = "ACCURACY"

EXTERNAL_ACCURACY = [
    "PER AUDITED ACCURACY",
    "PER AUDITED ACCURACY - WOC",
]

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()


# =========================================================
# DATABASE CONNECTION
# =========================================================

class PostgresConnection:
    """Compatibility wrapper for qmark SQL with psycopg 3."""

    def __init__(self, connection):
        self.connection = connection

    @staticmethod
    def _sql(query):
        return query.replace("?", "%s")

    def execute(self, query, params=()):
        return self.connection.execute(
            self._sql(query),
            params,
        )

    def executemany(self, query, params):
        return self.connection.cursor().executemany(
            self._sql(query),
            params,
        )

    def executescript(self, script):
        for statement in script.split(";"):
            if statement.strip():
                self.connection.execute(statement)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if exc_type:
                self.connection.rollback()
            else:
                self.connection.commit()
        finally:
            self.connection.close()

        return False


def db():

    if DATABASE_URL:

        if psycopg is None:
            raise RuntimeError(
                "Install psycopg[binary] to use Supabase PostgreSQL."
            )

        return PostgresConnection(
            psycopg.connect(
                DATABASE_URL,
                row_factory=dict_row,
                connect_timeout=15,
            )
        )

    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")

    return con


# =========================================================
# DATABASE HELPERS
# =========================================================

def is_postgres():
    return bool(DATABASE_URL)


def table_columns(con, table_name):

    if is_postgres():

        rows = con.execute(
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema=current_schema()
              AND table_name=?
            """,
            (table_name,),
        ).fetchall()

        return {row["column_name"] for row in rows}

    rows = con.execute(
        f"PRAGMA table_info({table_name})"
    ).fetchall()

    return {row["name"] for row in rows}


def add_column_if_missing(
    con,
    table_name,
    column_name,
    column_type,
):

    columns = table_columns(con, table_name)

    if column_name in columns:
        return

    con.execute(
        f"""
        ALTER TABLE {table_name}
        ADD COLUMN {column_name} {column_type}
        """
    )


# =========================================================
# DATABASE ROLE MIGRATION
# =========================================================

def migrate_user_roles(con, is_postgres_db=False):
    """
    Add QC and ADMIN to allowed roles.

    Preserve all existing accounts, IDs, password hashes,
    account statuses and creation dates.
    """

    if is_postgres_db:

        constraints = con.execute(
            """
            SELECT c.conname
            FROM pg_constraint c
            JOIN pg_class t
                ON t.oid = c.conrelid
            JOIN pg_namespace n
                ON n.oid = t.relnamespace
            WHERE t.relname = 'users'
              AND n.nspname = current_schema()
              AND c.contype = 'c'
              AND pg_get_constraintdef(c.oid) ILIKE '%role%'
            """
        ).fetchall()

        for item in constraints:

            constraint_name = (
                item["conname"].replace('"', '""')
            )

            con.execute(
                f'ALTER TABLE users DROP CONSTRAINT "{constraint_name}"'
            )

        con.execute(
            """
            ALTER TABLE users
            ADD CONSTRAINT users_role_check
            CHECK (role IN ('admin', 'biller', 'qc'))
            """
        )

        return

    schema = con.execute(
        """
        SELECT sql
        FROM sqlite_master
        WHERE type='table'
          AND name='users'
        """
    ).fetchone()

    if not schema:
        return

    table_sql = (schema["sql"] or "").lower()

    if "'qc'" in table_sql and "'admin'" in table_sql:
        return

    # Preserve existing records while rebuilding the users table.
    con.execute("PRAGMA foreign_keys=OFF")

    try:

        con.execute("BEGIN IMMEDIATE")

        con.execute(
            """
            CREATE TABLE users_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL
                    CHECK(role IN ('admin','biller','qc')),
                display_name TEXT,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            )
            """
        )

        con.execute(
            """
            INSERT INTO users_new
                (id, username, password_hash, role,
                 display_name, active, created_at)
            SELECT
                id, username, password_hash, role,
                display_name, active, created_at
            FROM users
            """
        )

        con.execute("DROP TABLE users")

        con.execute(
            "ALTER TABLE users_new RENAME TO users"
        )

        con.commit()

    except Exception:
        con.rollback()
        raise

    finally:
        con.execute("PRAGMA foreign_keys=ON")


# =========================================================
# INITIALIZE DATABASE
# =========================================================

def init_db():

    if is_postgres():

        ddl = """
        CREATE TABLE IF NOT EXISTS users (
          id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
          username TEXT UNIQUE NOT NULL,
          password_hash TEXT NOT NULL,
          role TEXT NOT NULL
            CHECK(role IN ('admin','biller','qc')),
          display_name TEXT,
          active INTEGER NOT NULL DEFAULT 1,
          created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS error_records (
          record_key TEXT PRIMARY KEY,
          billed_date TEXT,
          week TEXT,
          source TEXT,
          audit_type TEXT,
          biller_name TEXT,
          username TEXT,
          id_number TEXT,
          auditor TEXT,
          pro TEXT,
          zb_field TEXT,
          error_details TEXT,
          source_row INTEGER,
          imported_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_error_user
        ON error_records(username);

        CREATE INDEX IF NOT EXISTS idx_error_date
        ON error_records(billed_date);

        CREATE TABLE IF NOT EXISTS accuracy_records (
          id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
          sheet TEXT NOT NULL,
          username TEXT NOT NULL,
          biller_name TEXT,
          month TEXT,
          weekly TEXT,
          record_date TEXT,
          accuracy_type TEXT NOT NULL,
          accuracy_value DOUBLE PRECISION,
          bills_processed TEXT,
          audited TEXT,
          errors TEXT,
          source_row INTEGER,
          imported_at TEXT NOT NULL,
          UNIQUE(sheet, username, record_date, accuracy_type, source_row)
        );

        CREATE INDEX IF NOT EXISTS idx_accuracy_user
        ON accuracy_records(username, sheet);

        CREATE TABLE IF NOT EXISTS biller_responses (
          record_key TEXT PRIMARY KEY
            REFERENCES error_records(record_key) ON DELETE CASCADE,
          verified_error TEXT
            CHECK(verified_error IN ('Y','N') OR verified_error IS NULL),
          biller_comment TEXT NOT NULL DEFAULT '',
          qc_explanation TEXT NOT NULL DEFAULT '',
          qc_name TEXT NOT NULL DEFAULT '',
          updated_by TEXT,
          updated_at TEXT
        );

        CREATE TABLE IF NOT EXISTS import_history (
          id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
          filename TEXT NOT NULL,
          imported_by TEXT NOT NULL,
          imported_at TEXT NOT NULL,
          error_count INTEGER DEFAULT 0,
          internal_count INTEGER DEFAULT 0,
          external_count INTEGER DEFAULT 0,
          status TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS activity_logs (
          id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
          username TEXT,
          action TEXT,
          details TEXT,
          created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS projects (
          id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
          project_name TEXT UNIQUE NOT NULL,
          name TEXT NOT NULL,
          role TEXT NOT NULL
            CHECK(role IN ('biller','qc')),
          active INTEGER NOT NULL DEFAULT 1,
          created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS user_projects (
          user_id BIGINT NOT NULL
            REFERENCES users(id) ON DELETE CASCADE,
          project_id BIGINT NOT NULL
            REFERENCES projects(id) ON DELETE CASCADE,
          PRIMARY KEY(user_id, project_id)
        );

        CREATE INDEX IF NOT EXISTS idx_user_projects_project
        ON user_projects(project_id);
        """

        with db() as con:

            con.executescript(ddl)

            migrate_user_roles(
                con,
                is_postgres_db=True,
            )

            # Add project references to existing records.
            add_column_if_missing(
                con,
                "error_records",
                "project_id",
                "BIGINT REFERENCES projects(id) ON DELETE SET NULL",
            )

            add_column_if_missing(
                con,
                "accuracy_records",
                "project_id",
                "BIGINT REFERENCES projects(id) ON DELETE SET NULL",
            )

            admin_user = os.environ.get("ADMIN_USERNAME")
            admin_pass = os.environ.get("ADMIN_PASSWORD")

            exists = con.execute(
                "SELECT id FROM users LIMIT 1"
            ).fetchone()

            if not exists and admin_user and admin_pass:

                con.execute(
                    """
                    INSERT INTO users
                        (username,password_hash,role,
                         display_name,created_at)
                    VALUES(?,?,?,?,?)
                    """,
                    (
                        admin_user,
                        generate_password_hash(admin_pass),
                        "admin",
                        "System Administrator",
                        datetime.now().isoformat(timespec="seconds"),
                    ),
                )

        return

    # -----------------------------------------------------
    # SQLITE
    # -----------------------------------------------------

    with db() as con:

        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS users(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              username TEXT UNIQUE NOT NULL,
              password_hash TEXT NOT NULL,
              role TEXT NOT NULL
                CHECK(role IN ('admin','biller','qc')),
              display_name TEXT,
              active INTEGER NOT NULL DEFAULT 1,
              created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS error_records(
              record_key TEXT PRIMARY KEY,
              billed_date TEXT,
              week TEXT,
              source TEXT,
              audit_type TEXT,
              biller_name TEXT,
              username TEXT,
              id_number TEXT,
              auditor TEXT,
              pro TEXT,
              zb_field TEXT,
              error_details TEXT,
              source_row INTEGER,
              imported_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS accuracy_records(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              sheet TEXT NOT NULL,
              username TEXT NOT NULL,
              biller_name TEXT,
              month TEXT,
              weekly TEXT,
              record_date TEXT,
              accuracy_type TEXT NOT NULL,
              accuracy_value REAL,
              bills_processed TEXT,
              audited TEXT,
              errors TEXT,
              source_row INTEGER,
              imported_at TEXT NOT NULL,
              UNIQUE(sheet, username, record_date, accuracy_type, source_row)
            );

            CREATE INDEX IF NOT EXISTS idx_error_user
            ON error_records(username);

            CREATE INDEX IF NOT EXISTS idx_error_date
            ON error_records(billed_date);

            CREATE INDEX IF NOT EXISTS idx_accuracy_user
            ON accuracy_records(username, sheet);

            CREATE TABLE IF NOT EXISTS biller_responses(
              record_key TEXT PRIMARY KEY
                REFERENCES error_records(record_key) ON DELETE CASCADE,
              verified_error TEXT
                CHECK(verified_error IN ('Y','N') OR verified_error IS NULL),
              biller_comment TEXT NOT NULL DEFAULT '',
              qc_explanation TEXT NOT NULL DEFAULT '',
              qc_name TEXT NOT NULL DEFAULT '',
              updated_by TEXT,
              updated_at TEXT
            );

            CREATE TABLE IF NOT EXISTS import_history(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              filename TEXT NOT NULL,
              imported_by TEXT NOT NULL,
              imported_at TEXT NOT NULL,
              error_count INTEGER DEFAULT 0,
              internal_count INTEGER DEFAULT 0,
              external_count INTEGER DEFAULT 0,
              status TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS activity_logs(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              username TEXT,
              action TEXT,
              details TEXT,
              created_at TEXT
            );

            CREATE TABLE IF NOT EXISTS projects(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              project_name TEXT UNIQUE NOT NULL,
              name TEXT NOT NULL,
              role TEXT NOT NULL
                CHECK(role IN ('biller','qc')),
              active INTEGER NOT NULL DEFAULT 1,
              created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS user_projects(
              user_id INTEGER NOT NULL
                REFERENCES users(id) ON DELETE CASCADE,
              project_id INTEGER NOT NULL
                REFERENCES projects(id) ON DELETE CASCADE,
              PRIMARY KEY(user_id, project_id)
            );

            CREATE INDEX IF NOT EXISTS idx_user_projects_project
            ON user_projects(project_id);
            """
        )

        migrate_user_roles(
            con,
            is_postgres_db=False,
        )

        add_column_if_missing(
            con,
            "error_records",
            "project_id",
            "INTEGER REFERENCES projects(id) ON DELETE SET NULL",
        )

        add_column_if_missing(
            con,
            "accuracy_records",
            "project_id",
            "INTEGER REFERENCES projects(id) ON DELETE SET NULL",
        )

        admin_user = os.environ.get("ADMIN_USERNAME")
        admin_pass = os.environ.get("ADMIN_PASSWORD")

        exists = con.execute(
            "SELECT id FROM users LIMIT 1"
        ).fetchone()

        if not exists and admin_user and admin_pass:

            con.execute(
                """
                INSERT INTO users
                    (username,password_hash,role,
                     display_name,created_at)
                VALUES(?,?,?,?,?)
                """,
                (
                    admin_user,
                    generate_password_hash(admin_pass),
                    "admin",
                    "System Administrator",
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )


# =========================================================
# PROJECT ACCESS HELPERS
# =========================================================

def get_user_project_ids(user_id=None):

    user_id = user_id or session.get("user_id")

    if not user_id:
        return []

    with db() as con:

        user = con.execute(
            """
            SELECT role
            FROM users
            WHERE id=?
              AND active=1
            """,
            (user_id,),
        ).fetchone()

        if not user:
            return []

        if user["role"].lower() == "admin":

            rows = con.execute(
                """
                SELECT id
                FROM projects
                WHERE active=1
                ORDER BY project_name
                """
            ).fetchall()

        else:

            rows = con.execute(
                """
                SELECT p.id
                FROM projects p
                INNER JOIN user_projects up
                    ON up.project_id=p.id
                WHERE up.user_id=?
                  AND p.active=1
                ORDER BY p.project_name
                """,
                (user_id,),
            ).fetchall()

    return [row["id"] for row in rows]


def allowed_project_ids():
    """Return all active projects for Admin, assigned projects otherwise."""

    return get_user_project_ids()


def user_has_project(user_id, project_id):

    if not user_id or not project_id:
        return False

    with db() as con:

        user = con.execute(
            """
            SELECT role
            FROM users
            WHERE id=?
              AND active=1
            """,
            (user_id,),
        ).fetchone()

        if not user:
            return False

        if user["role"].lower() == "admin":

            project = con.execute(
                """
                SELECT id
                FROM projects
                WHERE id=?
                  AND active=1
                """,
                (project_id,),
            ).fetchone()

            return bool(project)

        assignment = con.execute(
            """
            SELECT 1
            FROM user_projects
            WHERE user_id=?
              AND project_id=?
            """,
            (user_id, project_id),
        ).fetchone()

        return bool(assignment)


def assign_user_projects(con, user_id, project_ids):

    clean_ids = sorted({
        int(value)
        for value in project_ids
        if str(value).isdigit()
    })

    con.execute(
        "DELETE FROM user_projects WHERE user_id=?",
        (user_id,),
    )

    for project_id in clean_ids:

        project = con.execute(
            """
            SELECT id
            FROM projects
            WHERE id=?
              AND active=1
            """,
            (project_id,),
        ).fetchone()

        if not project:
            raise ValueError(
                "One or more selected projects are invalid."
            )

        con.execute(
            """
            INSERT INTO user_projects(user_id,project_id)
            VALUES(?,?)
            """,
            (user_id, project_id),
        )


# =========================================================
# ACTIVITY LOG
# =========================================================

def log_action(username, action, details=""):

    with db() as con:

        con.execute(
            """
            INSERT INTO activity_logs
                (username,action,details,created_at)
            VALUES(?,?,?,?)
            """,
            (
                username,
                action,
                details,
                datetime.now().isoformat(timespec="seconds"),
            ),
        )


# =========================================================
# AUTHENTICATION HELPERS
# =========================================================

def login_required(fn):

    @wraps(fn)
    def wrapper(*args, **kwargs):

        if not session.get("user_id"):
            return redirect(url_for("login"))

        return fn(*args, **kwargs)

    return wrapper


def admin_required(fn):

    @wraps(fn)
    def wrapper(*args, **kwargs):

        if (session.get("role") or "").lower() != "admin":
            abort(403)

        return fn(*args, **kwargs)

    return wrapper


def role_required(*allowed_roles):

    def decorator(fn):

        @wraps(fn)
        def wrapper(*args, **kwargs):

            role = (session.get("role") or "").lower()

            if role not in allowed_roles:
                abort(403)

            return fn(*args, **kwargs)

        return wrapper

    return decorator


@app.before_request
def csrf_token():

    if "csrf" not in session:
        session["csrf"] = secrets.token_hex(16)


@app.context_processor
def inject_globals():

    return {
        "csrf": session.get("csrf", "")
    }


# =========================================================
# IMPORT HELPERS
# =========================================================

def norm_header(value):

    return " ".join(
        str(value).replace("\n", " ").strip().upper().split()
    )


def find_col(df, aliases):

    cols = {
        norm_header(c): c
        for c in df.columns
    }

    for alias in aliases:

        key = norm_header(alias)

        if key in cols:
            return cols[key]

    return None


def val(row, col):

    if not col:
        return ""

    value = row.get(col, "")

    if pd.isna(value):
        return ""

    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%Y-%m-%d")

    return str(value).strip()


def parse_date(value):

    if not value:
        return ""

    try:

        dt = pd.to_datetime(
            value,
            errors="coerce",
        )

        if pd.isna(dt):
            return str(value).strip()

        return dt.strftime("%Y-%m-%d")

    except Exception:
        return str(value).strip()


def month_from_date(value):

    parsed = parse_date(value)

    if not parsed:
        return ""

    try:
        return pd.to_datetime(parsed).strftime("%Y-%m")
    except Exception:
        return ""


def week_for_billed_date(value):
    """Monday-based week label."""

    if not value:
        return ""

    try:

        dt = pd.to_datetime(
            value,
            errors="coerce",
        )

        if pd.isna(dt):
            return ""

        monday = dt - pd.Timedelta(days=dt.weekday())
        sunday = monday + pd.Timedelta(days=6)

        return (
            f"{monday.strftime('%b %d')} - "
            f"{sunday.strftime('%b %d')}"
        )

    except Exception:
        return ""


def number_value(value):

    raw = str(value or "").strip()

    raw = raw.replace(",", "").replace("%", "")

    if not raw:
        return None

    try:
        return float(raw)
    except (ValueError, TypeError):
        return None


def normalize_accuracy(value):

    result = number_value(value)

    if result is not None and 0 <= result <= 1:
        result *= 100

    return result


def read_source(file_storage, csv_sheet="ERROR"):

    filename = secure_filename(
        file_storage.filename or "source"
    )

    ext = os.path.splitext(filename)[1].lower()

    if ext not in ALLOWED:
        raise ValueError(
            "Unsupported file type. Upload ODS, XLSX, XLS, or CSV."
        )

    raw = file_storage.read()

    if not raw:
        raise ValueError("The uploaded file is empty.")

    if ext == ".csv":

        df = None

        for encoding in (
            "utf-8-sig",
            "cp1252",
            "latin-1",
        ):

            try:

                df = pd.read_csv(
                    io.BytesIO(raw),
                    dtype=object,
                    encoding=encoding,
                )

                break

            except UnicodeDecodeError:
                continue

        if df is None:
            raise ValueError(
                "Unable to read the CSV encoding."
            )

        return {csv_sheet: df}, filename

    engine = "odf" if ext == ".ods" else None

    book = pd.ExcelFile(
        io.BytesIO(raw),
        engine=engine,
    )

    result = {}

    for sheet in ["ERROR", "INTERNAL", "EXTERNAL"]:

        if sheet in book.sheet_names:

            result[sheet] = pd.read_excel(
                book,
                sheet_name=sheet,
                dtype=object,
            )

    if not result:
        raise ValueError(
            "Workbook must contain at least one of these sheets: "
            "ERROR, INTERNAL, EXTERNAL."
        )

    return result, filename


def fingerprint(parts, occurrence):

    payload = json.dumps(
        [str(x or "").strip() for x in parts] + [occurrence],
        ensure_ascii=False,
    )

    return hashlib.sha256(
        payload.encode("utf-8")
    ).hexdigest()


# =========================================================
# PROJECT IMPORT LOOKUP
# =========================================================

def get_import_project_id(con, project_name):

    project_name = str(project_name or "").strip()

    if not project_name:
        return None

    project = con.execute(
        """
        SELECT id
        FROM projects
        WHERE project_name=?
          AND active=1
        """,
        (project_name,),
    ).fetchone()

    if not project:
        raise ValueError(
            f"Project '{project_name}' does not exist or is inactive. "
            "Add the project first or correct the PROJECT column."
        )

    return project["id"]


# =========================================================
# IMPORT WORKBOOK
# =========================================================

def import_workbook(sheets, filename, imported_by):
    """
    Non-destructive import.

    Existing records not included in a new upload are retained.

    Existing Biller and QC responses are preserved.

    Project assignment:
    - If PROJECT column exists, use its exact saved project name.
    - If PROJECT column does not exist, preserve the existing
      project assignment on matching records.
    - Existing historical records are not automatically assigned.
    """

    if "ERROR" in sheets and sheets["ERROR"].empty:
        raise ValueError(
            "The ERROR sheet is empty. Import cancelled to protect "
            "existing error records."
        )

    now = datetime.now().isoformat(timespec="seconds")

    error_df = sheets.get("ERROR", pd.DataFrame())
    internal_df = sheets.get("INTERNAL", pd.DataFrame())
    external_df = sheets.get("EXTERNAL", pd.DataFrame())

    err_count = 0
    int_rows = 0
    ext_rows = 0

    with db() as con:

        # -------------------------------------------------
        # ERROR SHEET
        # -------------------------------------------------

        if "ERROR" in sheets and not error_df.empty:

            cols = {
                key: find_col(error_df, aliases)
                for key, aliases in ERROR_ALIASES.items()
            }

            if not cols["user"] or not cols["pro"]:
                raise ValueError(
                    "ERROR sheet is missing DDC USER ID "
                    "(or DDC USER) and/or PRO#."
                )

            has_project_column = bool(cols["project"])

            seen = {}

            for idx, row in error_df.iterrows():

                billed_date = parse_date(
                    val(row, cols["date"])
                )

                username = val(
                    row,
                    cols["user"],
                ).upper()

                pro = val(
                    row,
                    cols["pro"],
                )

                if not username or not pro:
                    continue

                project_id = None

                if has_project_column:

                    project_id = get_import_project_id(
                        con,
                        val(row, cols["project"]),
                    )

                fields = [
                    billed_date,
                    val(row, cols["source"]),
                    val(row, cols["audit_type"]),
                    val(row, cols["name"]),
                    username,
                    val(row, cols["id"]),
                    val(row, cols["auditor"]),
                    pro,
                    val(row, cols["zb"]),
                    val(row, cols["details"]),
                ]

                # Project is intentionally excluded from the key.
                # This preserves existing biller_responses.
                base = tuple(fields)

                seen[base] = seen.get(base, 0) + 1

                key = fingerprint(
                    fields,
                    seen[base],
                )

                if has_project_column:

                    con.execute(
                        """
                        INSERT INTO error_records
                            (record_key,billed_date,week,source,
                             audit_type,biller_name,username,
                             id_number,auditor,pro,zb_field,
                             error_details,source_row,imported_at,
                             project_id)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)

                        ON CONFLICT(record_key) DO UPDATE SET
                            billed_date=excluded.billed_date,
                            week=excluded.week,
                            source=excluded.source,
                            audit_type=excluded.audit_type,
                            biller_name=excluded.biller_name,
                            username=excluded.username,
                            id_number=excluded.id_number,
                            auditor=excluded.auditor,
                            pro=excluded.pro,
                            zb_field=excluded.zb_field,
                            error_details=excluded.error_details,
                            source_row=excluded.source_row,
                            imported_at=excluded.imported_at,
                            project_id=excluded.project_id
                        """,
                        (
                            key,
                            billed_date,
                            week_for_billed_date(billed_date),
                            fields[1],
                            fields[2],
                            fields[3],
                            username,
                            fields[5],
                            fields[6],
                            pro,
                            fields[8],
                            fields[9],
                            int(idx) + 2,
                            now,
                            project_id,
                        ),
                    )

                else:

                    # Old-format workbook without PROJECT:
                    # do not overwrite an existing project_id.
                    con.execute(
                        """
                        INSERT INTO error_records
                            (record_key,billed_date,week,source,
                             audit_type,biller_name,username,
                             id_number,auditor,pro,zb_field,
                             error_details,source_row,imported_at)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)

                        ON CONFLICT(record_key) DO UPDATE SET
                            billed_date=excluded.billed_date,
                            week=excluded.week,
                            source=excluded.source,
                            audit_type=excluded.audit_type,
                            biller_name=excluded.biller_name,
                            username=excluded.username,
                            id_number=excluded.id_number,
                            auditor=excluded.auditor,
                            pro=excluded.pro,
                            zb_field=excluded.zb_field,
                            error_details=excluded.error_details,
                            source_row=excluded.source_row,
                            imported_at=excluded.imported_at
                        """,
                        (
                            key,
                            billed_date,
                            week_for_billed_date(billed_date),
                            fields[1],
                            fields[2],
                            fields[3],
                            username,
                            fields[5],
                            fields[6],
                            pro,
                            fields[8],
                            fields[9],
                            int(idx) + 2,
                            now,
                        ),
                    )

                err_count += 1

        # -------------------------------------------------
        # INTERNAL AND EXTERNAL ACCURACY
        # -------------------------------------------------

        for sheet_name, df in [
            ("INTERNAL", internal_df),
            ("EXTERNAL", external_df),
        ]:

            if sheet_name not in sheets or df.empty:
                continue

            user_col = find_col(df, ["USER"])
            name_col = find_col(df, ["NAME"])
            project_col = find_col(df, ["PROJECT"])

            if not user_col:
                raise ValueError(
                    f"{sheet_name} sheet is missing the USER column."
                )

            if sheet_name == "INTERNAL":

                date_col = find_col(df, ["DATE"])
                month_col = find_col(df, ["MONTH"])
                week_col = find_col(df, ["WEEKLY"])

                accuracy_cols = [
                    INTERNAL_ACCURACY
                ]

            else:

                date_col = find_col(
                    df,
                    ["DATE STARTED", "DATE"],
                )

                month_col = find_col(df, ["MONTH"])
                week_col = find_col(df, ["WEEKLY"])

                accuracy_cols = EXTERNAL_ACCURACY

            audited_col = find_col(df, ["AUDITED"])
            bills_col = find_col(df, ["BILLS PROCESSED"])
            error_col = find_col(df, ["ERROR"])
            woc_error_col = find_col(df, ["ERROR - WOC"])

            for idx, row in df.iterrows():

                username = val(
                    row,
                    user_col,
                ).upper()

                if not username:
                    continue

                record_date = parse_date(
                    val(row, date_col)
                )

                month_value = val(row, month_col)

                if not month_value:
                    month_value = month_from_date(record_date)

                weekly_value = val(row, week_col)

                if not weekly_value:
                    weekly_value = week_for_billed_date(
                        record_date
                    )

                audited_value = val(row, audited_col)
                bills_value = val(row, bills_col)

                project_id = None

                if project_col:

                    project_id = get_import_project_id(
                        con,
                        val(row, project_col),
                    )

                for accuracy_type in accuracy_cols:

                    acc_col = find_col(
                        df,
                        [accuracy_type],
                    )

                    if not acc_col:
                        continue

                    accuracy = normalize_accuracy(
                        val(row, acc_col)
                    )

                    if sheet_name == "EXTERNAL":

                        if (
                            accuracy_type
                            == "PER AUDITED ACCURACY - WOC"
                        ):

                            errors_value = val(
                                row,
                                woc_error_col,
                            )

                        else:

                            errors_value = val(
                                row,
                                error_col,
                            )

                    else:

                        errors_value = val(
                            row,
                            error_col,
                        )

                    if project_col:

                        con.execute(
                            """
                            INSERT INTO accuracy_records
                                (sheet,username,biller_name,month,
                                 weekly,record_date,accuracy_type,
                                 accuracy_value,bills_processed,
                                 audited,errors,source_row,
                                 imported_at,project_id)
                            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)

                            ON CONFLICT
                                (sheet,username,record_date,
                                 accuracy_type,source_row)
                            DO UPDATE SET
                                biller_name=excluded.biller_name,
                                month=excluded.month,
                                weekly=excluded.weekly,
                                accuracy_value=excluded.accuracy_value,
                                bills_processed=excluded.bills_processed,
                                audited=excluded.audited,
                                errors=excluded.errors,
                                imported_at=excluded.imported_at,
                                project_id=excluded.project_id
                            """,
                            (
                                sheet_name,
                                username,
                                val(row, name_col),
                                month_value,
                                weekly_value,
                                record_date,
                                accuracy_type,
                                accuracy,
                                bills_value,
                                audited_value,
                                errors_value,
                                int(idx) + 2,
                                now,
                                project_id,
                            ),
                        )

                    else:

                        # Old workbook without PROJECT:
                        # preserve existing project assignment.
                        con.execute(
                            """
                            INSERT INTO accuracy_records
                                (sheet,username,biller_name,month,
                                 weekly,record_date,accuracy_type,
                                 accuracy_value,bills_processed,
                                 audited,errors,source_row,
                                 imported_at)
                            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)

                            ON CONFLICT
                                (sheet,username,record_date,
                                 accuracy_type,source_row)
                            DO UPDATE SET
                                biller_name=excluded.biller_name,
                                month=excluded.month,
                                weekly=excluded.weekly,
                                accuracy_value=excluded.accuracy_value,
                                bills_processed=excluded.bills_processed,
                                audited=excluded.audited,
                                errors=excluded.errors,
                                imported_at=excluded.imported_at
                            """,
                            (
                                sheet_name,
                                username,
                                val(row, name_col),
                                month_value,
                                weekly_value,
                                record_date,
                                accuracy_type,
                                accuracy,
                                bills_value,
                                audited_value,
                                errors_value,
                                int(idx) + 2,
                                now,
                            ),
                        )

                if sheet_name == "INTERNAL":
                    int_rows += 1
                else:
                    ext_rows += 1

        con.execute(
            """
            INSERT INTO import_history
                (filename,imported_by,imported_at,error_count,
                 internal_count,external_count,status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                filename,
                imported_by,
                now,
                err_count,
                int_rows,
                ext_rows,
                "SUCCESS",
            ),
        )

    log_action(
        imported_by,
        "SOURCE_IMPORT",
        (
            f"{filename}: ERROR={err_count}, "
            f"INTERNAL={int_rows}, EXTERNAL={ext_rows}"
        ),
    )

    return err_count, int_rows, ext_rows


# =========================================================
# ACCURACY DISPLAY HELPERS
# =========================================================

def get_biller_months(username):

    with db() as con:

        rows = con.execute(
            """
            SELECT DISTINCT month
            FROM accuracy_records
            WHERE UPPER(username)=?
              AND month IS NOT NULL
              AND month <> ''
            ORDER BY month DESC
            """,
            (username.upper(),),
        ).fetchall()

    return [row["month"] for row in rows]


def selected_month_value(username):

    months = get_biller_months(username)

    requested = request.args.get(
        "month",
        "",
    ).strip()

    if requested and requested in months:
        return requested

    if months:
        return months[0]

    return datetime.now().strftime("%Y-%m")


def weighted_accuracy(rows):
    """
    Weighted accuracy:
    (total audited - total errors) / total audited * 100.
    """

    total_audited = 0.0
    total_errors = 0.0

    for row in rows:

        audited = number_value(row["audited"])
        errors = number_value(row["errors"])

        if audited is None:
            continue

        total_audited += audited
        total_errors += errors or 0

    if total_audited <= 0:
        return None

    return (
        (total_audited - total_errors)
        / total_audited
        * 100
    )


def monthly_accuracy_summary(
    rows,
    sheet,
    accuracy_type=None,
):

    selected = [
        row for row in rows
        if row["sheet"] == sheet
        and (
            accuracy_type is None
            or row["accuracy_type"] == accuracy_type
        )
    ]

    return {
        "accuracy": weighted_accuracy(selected),

        "audited": sum(
            number_value(r["audited"]) or 0
            for r in selected
        ),

        "errors": sum(
            number_value(r["errors"]) or 0
            for r in selected
        ),
    }


def format_accuracy(value):

    if value is None:
        return "—"

    return f"{value:.2f}%"


def get_accuracy_rows(
    username,
    sheet,
    month,
    project_ids=None,
):

    with db() as con:

        sql = """
            SELECT *
            FROM accuracy_records
            WHERE UPPER(username)=?
              AND sheet=?
              AND month=?
        """

        params = [
            username.upper(),
            sheet,
            month,
        ]

        if project_ids is not None:

            if not project_ids:
                return []

            placeholders = ",".join(
                ["?"] * len(project_ids)
            )

            sql += f" AND project_id IN ({placeholders})"
            params.extend(project_ids)

        sql += " ORDER BY record_date, source_row"

        rows = con.execute(
            sql,
            tuple(params),
        ).fetchall()

    return rows


def group_accuracy_rows(rows, period):
    """
    Aggregate records into weighted accuracy rows.
    """

    groups = {}

    for row in rows:

        if period == "WEEKLY":

            label = (
                row["weekly"]
                or week_for_billed_date(row["record_date"])
                or "Undated"
            )

        elif period == "MONTHLY":

            label = row["month"] or "Undated"

        else:

            label = row["record_date"] or "Undated"

        if label not in groups:

            groups[label] = {
                "sort_date": row["record_date"] or "9999-99-99",
                "types": {},
            }

        group = groups[label]

        if row["record_date"]:

            group["sort_date"] = min(
                group["sort_date"],
                row["record_date"],
            )

        accuracy_type = row["accuracy_type"]

        if accuracy_type not in group["types"]:

            group["types"][accuracy_type] = {
                "audited": 0.0,
                "errors": 0.0,
                "has_audited": False,
            }

        totals = group["types"][accuracy_type]

        audited = number_value(row["audited"])
        errors = number_value(row["errors"])

        if audited is not None:

            totals["audited"] += audited
            totals["has_audited"] = True
            totals["errors"] += errors or 0.0

    def calculate_group_accuracy(totals):

        if not totals or not totals["has_audited"]:
            return None

        total_audited = totals["audited"]

        if total_audited <= 0:
            return None

        return (
            (total_audited - totals["errors"])
            / total_audited
            * 100
        )

    result = []

    is_internal = (
        bool(rows)
        and rows[0]["sheet"] == "INTERNAL"
    )

    for label, group in groups.items():

        values = group["types"]

        if is_internal:

            internal = calculate_group_accuracy(
                values.get(INTERNAL_ACCURACY)
            )

            result.append({
                "label": label,
                "internal": internal,
                "_sort": group["sort_date"],
            })

        else:

            wc = calculate_group_accuracy(
                values.get("PER AUDITED ACCURACY")
            )

            woc = calculate_group_accuracy(
                values.get("PER AUDITED ACCURACY - WOC")
            )

            result.append({
                "label": label,
                "wc": wc,
                "woc": woc,
                "_sort": group["sort_date"],
            })

    result.sort(
        key=lambda item: (
            item["_sort"],
            item["label"],
        )
    )

    for item in result:
        item.pop("_sort", None)

    return result


def month_year_options(months):

    years = sorted(
        {
            m[:4]
            for m in months
            if len(m) >= 4
        },
        reverse=True,
    )

    return years


# =========================================================
# ACCURACY MONITORING HELPERS
# =========================================================

def get_period_start(value, period):

    parsed = pd.to_datetime(
        value,
        errors="coerce",
    )

    if pd.isna(parsed):
        return None

    if period == "DAILY":
        return parsed.normalize()

    if period == "WEEKLY":
        return (
            parsed.normalize()
            - pd.Timedelta(days=parsed.weekday())
        )

    return parsed.to_period("M").start_time.normalize()


def get_period_end(start, period):

    if period == "DAILY":
        return start + pd.Timedelta(days=1)

    if period == "WEEKLY":
        return start + pd.Timedelta(days=7)

    return start + pd.offsets.MonthBegin(1)


def get_period_rows(rows, period, reference_date):

    reference_start = get_period_start(
        reference_date,
        period,
    )

    if reference_start is None:
        return [], None

    reference_end = get_period_end(
        reference_start,
        period,
    )

    selected = []

    for row in rows:

        record_date = pd.to_datetime(
            row["record_date"],
            errors="coerce",
        )

        if pd.isna(record_date):
            continue

        if (
            record_date >= reference_start
            and record_date < reference_end
        ):
            selected.append(row)

    return selected, reference_start


def previous_period_start(start, period):

    if period == "DAILY":
        return start - pd.Timedelta(days=1)

    if period == "WEEKLY":
        return start - pd.Timedelta(days=7)

    return start - pd.offsets.MonthBegin(1)


def calculate_improvement(
    rows,
    sheet,
    accuracy_type,
    period,
    reference_date,
):

    current_rows, current_start = get_period_rows(
        rows,
        period,
        reference_date,
    )

    if current_start is None:
        return {
            "current": None,
            "previous": None,
            "difference": None,
            "status": "No Data",
            "note": "No accuracy data available.",
        }

    current_selected = [
        row for row in current_rows
        if row["sheet"] == sheet
        and row["accuracy_type"] == accuracy_type
    ]

    current = weighted_accuracy(current_selected)

    previous_start = previous_period_start(
        current_start,
        period,
    )

    previous_end = get_period_end(
        previous_start,
        period,
    )

    previous_rows = []

    for row in rows:

        record_date = pd.to_datetime(
            row["record_date"],
            errors="coerce",
        )

        if pd.isna(record_date):
            continue

        if (
            record_date >= previous_start
            and record_date < previous_end
            and row["sheet"] == sheet
            and row["accuracy_type"] == accuracy_type
        ):
            previous_rows.append(row)

    previous = weighted_accuracy(previous_rows)

    if current is None:

        status = "No Data"
        difference = None
        note = "No accuracy data available for this period."

    elif previous is None:

        status = "New Record"
        difference = None
        note = "No previous-period accuracy available for comparison."

    else:

        difference = current - previous

        if difference > 0:

            status = "Improved"
            note = (
                f"Improved by {difference:.2f} percentage points."
            )

        elif difference < 0:

            status = "Declined"
            note = (
                f"Declined by {abs(difference):.2f} "
                "percentage points."
            )

        else:

            status = "No Change"
            note = "Accuracy remained the same."

    if current is not None and current < 90:

        note += " Still below the 90% target."

    return {
        "current": current,
        "previous": previous,
        "difference": difference,
        "status": status,
        "note": note,
    }


def get_accuracy_monitoring(project_ids, period="MONTHLY"):

    if not project_ids:
        return []

    placeholders = ",".join(
        ["?"] * len(project_ids)
    )

    with db() as con:

        rows = con.execute(
            f"""
            SELECT
                a.*,
                p.project_name
            FROM accuracy_records a
            LEFT JOIN projects p
                ON p.id=a.project_id
            WHERE a.project_id IN ({placeholders})
            ORDER BY a.record_date
            """,
            tuple(project_ids),
        ).fetchall()

    records = {}

    for row in rows:

        project_id = row["project_id"]
        username = row["username"].upper()

        key = (
            project_id,
            username,
        )

        records.setdefault(key, []).append(row)

    monitoring = []

    for (project_id, username), user_rows in records.items():

        biller_name = (
            user_rows[0]["biller_name"]
            or username
        )

        latest_date = max(
            (
                row["record_date"]
                for row in user_rows
                if row["record_date"]
            ),
            default="",
        )

        if not latest_date:
            continue

        internal = calculate_improvement(
            user_rows,
            "INTERNAL",
            INTERNAL_ACCURACY,
            period,
            latest_date,
        )

        external_wc = calculate_improvement(
            user_rows,
            "EXTERNAL",
            "PER AUDITED ACCURACY",
            period,
            latest_date,
        )

        external_woc = calculate_improvement(
            user_rows,
            "EXTERNAL",
            "PER AUDITED ACCURACY - WOC",
            period,
            latest_date,
        )

        project_name = (
            user_rows[0]["project_name"]
            or "Unassigned"
        )

        monitoring.append({
            "project_id": project_id,
            "project_name": project_name,
            "username": username,
            "biller_name": biller_name,
            "internal": internal,
            "external_wc": external_wc,
            "external_woc": external_woc,
        })

    monitoring.sort(
        key=lambda item: (
            item["project_name"],
            item["biller_name"].upper(),
        )
    )

    return monitoring


# =========================================================
# GENERAL ROUTES
# =========================================================

@app.route("/", methods=["GET"])
def index():

    if not session.get("user_id"):
        return redirect(url_for("login"))

    role = (session.get("role") or "").lower()

    if role == "admin":
        return redirect(url_for("admin_dashboard"))

    if role == "qc":
        return redirect(url_for("qc_errors"))

    if role == "biller":
        return redirect(url_for("biller_dashboard"))

    session.clear()

    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        if request.form.get("csrf") != session.get("csrf"):
            abort(400)

        username = request.form.get(
            "username",
            "",
        ).strip()

        password = request.form.get(
            "password",
            "",
        )

        with db() as con:

            user = con.execute(
                """
                SELECT *
                FROM users
                WHERE UPPER(username)=UPPER(?)
                  AND active=1
                """,
                (username,),
            ).fetchone()

        if (
            user
            and user["role"] in ("admin", "biller", "qc")
            and check_password_hash(
                user["password_hash"],
                password,
            )
        ):

            session.clear()

            session.update(
                user_id=user["id"],
                username=user["username"],
                role=user["role"],
                display_name=(
                    user["display_name"]
                    or user["username"]
                ),
                csrf=secrets.token_hex(16),
            )

            log_action(
                user["username"],
                "LOGIN",
            )

            return redirect(url_for("index"))

        flash(
            "Invalid username or password.",
            "error",
        )

    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():

    log_action(
        session.get("username"),
        "LOGOUT",
    )

    session.clear()

    return redirect(url_for("login"))


# =========================================================
# ADMIN DASHBOARD
# =========================================================

@app.route("/admin")
@login_required
@admin_required
def admin_dashboard():

    with db() as con:

        counts = {

            "errors": con.execute(
                """
                SELECT COUNT(*) AS total
                FROM error_records
                """
            ).fetchone()["total"],

            "billers": con.execute(
                """
                SELECT COUNT(*) AS total
                FROM users
                WHERE role='biller'
                  AND active=1
                """
            ).fetchone()["total"],

            "verified": con.execute(
                """
                SELECT COUNT(*) AS total
                FROM biller_responses
                WHERE verified_error='Y'
                """
            ).fetchone()["total"],

            "disputed": con.execute(
                """
                SELECT COUNT(*) AS total
                FROM biller_responses
                WHERE verified_error='N'
                """
            ).fetchone()["total"],

            "pending": con.execute(
                """
                SELECT COUNT(*) AS total
                FROM error_records e
                LEFT JOIN biller_responses r
                    USING(record_key)
                WHERE r.verified_error IS NULL
                """
            ).fetchone()["total"],
        }

        history = con.execute(
            """
            SELECT *
            FROM import_history
            ORDER BY id DESC
            LIMIT 8
            """
        ).fetchall()

        billers = con.execute(
            """
            SELECT id,username,display_name,active
            FROM users
            WHERE role='biller'
            ORDER BY username
            """
        ).fetchall()

        qc_users = con.execute(
            """
            SELECT id,username,display_name,active
            FROM users
            WHERE role='qc'
            ORDER BY username
            """
        ).fetchall()

        admin_users = con.execute(
            """
            SELECT id,username,display_name,active
            FROM users
            WHERE role='admin'
            ORDER BY username
            """
        ).fetchall()

        projects = con.execute(
            """
            SELECT *
            FROM projects
            ORDER BY project_name
            """
        ).fetchall()

        project_assignments = con.execute(
            """
            SELECT
                up.user_id,
                p.id AS project_id,
                p.project_name
            FROM user_projects up
            JOIN projects p
                ON p.id=up.project_id
            ORDER BY p.project_name
            """
        ).fetchall()

    return render_template(
        "admin.html",
        counts=counts,
        history=history,
        billers=billers,
        qc_users=qc_users,
        admin_users=admin_users,
        projects=projects,
        project_assignments=project_assignments,
    )


# =========================================================
# ADMIN SOURCE UPLOAD
# =========================================================

@app.route("/admin/upload", methods=["POST"])
@login_required
@admin_required
def upload_source():

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    file = request.files.get("source_file")

    if not file or not file.filename:

        flash(
            "Choose an ODS, Excel, or CSV source file.",
            "error",
        )

        return redirect(url_for("admin_dashboard"))

    try:

        sheets, filename = read_source(
            file,
            request.form.get("csv_sheet", "ERROR"),
        )

        result = import_workbook(
            sheets,
            filename,
            session["username"],
        )

        flash(
            f"Import successful. ERROR: {result[0]}, "
            f"INTERNAL: {result[1]}, EXTERNAL: {result[2]} rows.",
            "success",
        )

    except Exception as exc:

        flash(
            f"Import failed: {exc}",
            "error",
        )

    return redirect(url_for("admin_dashboard"))


# =========================================================
# ADMIN ACCOUNT MANAGEMENT
# =========================================================

@app.route("/admin/users", methods=["POST"])
@login_required
@admin_required
def create_user():

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    username = request.form.get(
        "username",
        "",
    ).strip().upper()

    display = request.form.get(
        "display_name",
        "",
    ).strip()

    password = request.form.get(
        "password",
        "",
    )

    role = request.form.get(
        "role",
        "biller",
    ).strip().lower()

    project_ids = request.form.getlist("project_ids")

    if role not in ("admin", "biller", "qc"):

        flash(
            "Invalid account role.",
            "error",
        )

        return redirect(url_for("admin_dashboard"))

    if not username or len(password) < 6:

        flash(
            "Username is required and password must be at least 6 characters.",
            "error",
        )

        return redirect(url_for("admin_dashboard"))

    if role != "admin" and not project_ids:

        flash(
            "Select at least one project for this account.",
            "error",
        )

        return redirect(url_for("admin_dashboard"))

    try:

        with db() as con:

            cursor = con.execute(
                """
                INSERT INTO users
                    (username,password_hash,role,
                     display_name,created_at)
                VALUES(?,?,?,?,?)
                """,
                (
                    username,
                    generate_password_hash(password),
                    role,
                    display or username,
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )

            if is_postgres():
                user_id = cursor.fetchone()["id"]
            else:
                user_id = cursor.lastrowid

            if role != "admin":

                assign_user_projects(
                    con,
                    user_id,
                    project_ids,
                )

        log_action(
            session["username"],
            "CREATE_ACCOUNT",
            f"{role.upper()}: {username}",
        )

        flash(
            f"{role.upper()} account {username} created.",
            "success",
        )

    except Exception as exc:

        is_integrity_error = (
            isinstance(exc, sqlite3.IntegrityError)
            or (
                psycopg is not None
                and isinstance(exc, psycopg.IntegrityError)
            )
        )

        if is_integrity_error:

            flash(
                "Username already exists or project assignment is invalid.",
                "error",
            )

        else:
            raise

    return redirect(url_for("admin_dashboard"))


@app.route(
    "/admin/user/<int:user_id>/projects",
    methods=["POST"],
)
@login_required
@admin_required
def update_user_projects(user_id):

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    project_ids = request.form.getlist("project_ids")

    with db() as con:

        user = con.execute(
            """
            SELECT id,username,role
            FROM users
            WHERE id=?
            """,
            (user_id,),
        ).fetchone()

        if not user:
            abort(404)

        if user["role"] == "admin":

            flash(
                "Admin accounts automatically have access to all projects.",
                "info",
            )

            return redirect(url_for("admin_dashboard"))

        if not project_ids:

            flash(
                "Select at least one project.",
                "error",
            )

            return redirect(url_for("admin_dashboard"))

        assign_user_projects(
            con,
            user_id,
            project_ids,
        )

    log_action(
        session["username"],
        "UPDATE_USER_PROJECTS",
        user["username"],
    )

    flash(
        "Project assignments updated.",
        "success",
    )

    return redirect(url_for("admin_dashboard"))


@app.route(
    "/admin/user/<int:user_id>/toggle",
    methods=["POST"],
)
@login_required
@admin_required
def toggle_user(user_id):

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    with db() as con:

        user = con.execute(
            """
            SELECT *
            FROM users
            WHERE id=?
              AND role='biller'
            """,
            (user_id,),
        ).fetchone()

        if user:

            con.execute(
                "UPDATE users SET active=? WHERE id=?",
                (
                    0 if user["active"] else 1,
                    user_id,
                ),
            )

            log_action(
                session["username"],
                "TOGGLE_BILLER",
                user["username"],
            )

    return redirect(url_for("admin_dashboard"))


@app.route("/admin/qc-users", methods=["POST"])
@login_required
@admin_required
def create_qc_user():

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    # Compatibility route for older QC account forms.
    request_role = request.form.get("role", "qc").lower()

    if request_role != "qc":
        abort(400)

    username = request.form.get(
        "username",
        "",
    ).strip().upper()

    display = request.form.get(
        "display_name",
        "",
    ).strip()

    password = request.form.get(
        "password",
        "",
    )

    project_ids = request.form.getlist("project_ids")

    if not username or not display or len(password) < 6:

        flash(
            "QC username, QC name, and a password of at least 6 characters are required.",
            "error",
        )

        return redirect(url_for("admin_dashboard"))

    if not project_ids:

        flash(
            "Select at least one project for this QC account.",
            "error",
        )

        return redirect(url_for("admin_dashboard"))

    try:

        with db() as con:

            cursor = con.execute(
                """
                INSERT INTO users
                    (username,password_hash,role,
                     display_name,created_at)
                VALUES(?,?,?,?,?)
                """,
                (
                    username,
                    generate_password_hash(password),
                    "qc",
                    display,
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )

            if is_postgres():
                user_id = cursor.fetchone()["id"]
            else:
                user_id = cursor.lastrowid

            assign_user_projects(
                con,
                user_id,
                project_ids,
            )

        log_action(
            session["username"],
            "CREATE_QC",
            username,
        )

        flash(
            f"QC account {username} created.",
            "success",
        )

    except Exception as exc:

        is_integrity_error = (
            isinstance(exc, sqlite3.IntegrityError)
            or (
                psycopg is not None
                and isinstance(exc, psycopg.IntegrityError)
            )
        )

        if is_integrity_error:

            flash(
                "That username already exists or project assignment is invalid.",
                "error",
            )

        else:
            raise

    return redirect(url_for("admin_dashboard"))


@app.route(
    "/admin/qc-user/<int:user_id>/toggle",
    methods=["POST"],
)
@login_required
@admin_required
def toggle_qc_user(user_id):

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    with db() as con:

        user = con.execute(
            """
            SELECT *
            FROM users
            WHERE id=?
              AND role='qc'
            """,
            (user_id,),
        ).fetchone()

        if user:

            con.execute(
                "UPDATE users SET active=? WHERE id=?",
                (
                    0 if user["active"] else 1,
                    user_id,
                ),
            )

            log_action(
                session["username"],
                "TOGGLE_QC",
                user["username"],
            )

    return redirect(url_for("admin_dashboard"))


@app.route(
    "/admin/admin-user/<int:user_id>/toggle",
    methods=["POST"],
)
@login_required
@admin_required
def toggle_admin_user(user_id):

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    current_user_id = session.get("user_id")

    if user_id == current_user_id:

        flash(
            "You cannot disable your own Admin account.",
            "error",
        )

        return redirect(url_for("admin_dashboard"))

    with db() as con:

        user = con.execute(
            """
            SELECT *
            FROM users
            WHERE id=?
              AND role='admin'
            """,
            (user_id,),
        ).fetchone()

        if not user:
            abort(404)

        if user["active"]:

            active_admins = con.execute(
                """
                SELECT COUNT(*) AS total
                FROM users
                WHERE role='admin'
                  AND active=1
                """
            ).fetchone()["total"]

            if active_admins <= 1:

                flash(
                    "You cannot disable the last active Admin account.",
                    "error",
                )

                return redirect(url_for("admin_dashboard"))

        con.execute(
            "UPDATE users SET active=? WHERE id=?",
            (
                0 if user["active"] else 1,
                user_id,
            ),
        )

    log_action(
        session["username"],
        "TOGGLE_ADMIN",
        user["username"],
    )

    return redirect(url_for("admin_dashboard"))


# =========================================================
# ADMIN PROJECT MANAGEMENT
# =========================================================

@app.route("/admin/projects", methods=["POST"])
@login_required
@admin_required
def create_project():

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    project_name = request.form.get(
        "project_name",
        "",
    ).strip()

    name = request.form.get(
        "name",
        "",
    ).strip()

    role = request.form.get(
        "role",
        "",
    ).strip().lower()

    if not project_name or not name:

        flash(
            "Project Name and Name are required.",
            "error",
        )

        return redirect(url_for("admin_dashboard"))

    if role not in ("biller", "qc"):

        flash(
            "Select a valid project role: BILLER or QC.",
            "error",
        )

        return redirect(url_for("admin_dashboard"))

    try:

        with db() as con:

            con.execute(
                """
                INSERT INTO projects
                    (project_name,name,role,active,created_at)
                VALUES(?,?,?,?,?)
                """,
                (
                    project_name,
                    name,
                    role,
                    1,
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )

        log_action(
            session["username"],
            "CREATE_PROJECT",
            project_name,
        )

        flash(
            f"Project {project_name} created.",
            "success",
        )

    except Exception as exc:

        is_integrity_error = (
            isinstance(exc, sqlite3.IntegrityError)
            or (
                psycopg is not None
                and isinstance(exc, psycopg.IntegrityError)
            )
        )

        if is_integrity_error:

            flash(
                "That Project Name already exists.",
                "error",
            )

        else:
            raise

    return redirect(url_for("admin_dashboard"))


@app.route(
    "/admin/project/<int:project_id>/toggle",
    methods=["POST"],
)
@login_required
@admin_required
def toggle_project(project_id):

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    with db() as con:

        project = con.execute(
            """
            SELECT *
            FROM projects
            WHERE id=?
            """,
            (project_id,),
        ).fetchone()

        if not project:
            abort(404)

        con.execute(
            "UPDATE projects SET active=? WHERE id=?",
            (
                0 if project["active"] else 1,
                project_id,
            ),
        )

    log_action(
        session["username"],
        "TOGGLE_PROJECT",
        project["project_name"],
    )

    return redirect(url_for("admin_dashboard"))


@app.route(
    "/admin/project/<int:project_id>/edit",
    methods=["POST"],
)
@login_required
@admin_required
def edit_project(project_id):

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    project_name = request.form.get(
        "project_name",
        "",
    ).strip()

    name = request.form.get(
        "name",
        "",
    ).strip()

    role = request.form.get(
        "role",
        "",
    ).strip().lower()

    if not project_name or not name or role not in ("biller", "qc"):

        flash(
            "Enter a valid Project Name, Name, and Role.",
            "error",
        )

        return redirect(url_for("admin_dashboard"))

    try:

        with db() as con:

            project = con.execute(
                """
                SELECT id
                FROM projects
                WHERE id=?
                """,
                (project_id,),
            ).fetchone()

            if not project:
                abort(404)

            con.execute(
                """
                UPDATE projects
                SET project_name=?,
                    name=?,
                    role=?
                WHERE id=?
                """,
                (
                    project_name,
                    name,
                    role,
                    project_id,
                ),
            )

        log_action(
            session["username"],
            "EDIT_PROJECT",
            project_name,
        )

        flash(
            "Project updated.",
            "success",
        )

    except Exception as exc:

        is_integrity_error = (
            isinstance(exc, sqlite3.IntegrityError)
            or (
                psycopg is not None
                and isinstance(exc, psycopg.IntegrityError)
            )
        )

        if is_integrity_error:

            flash(
                "That Project Name is already in use.",
                "error",
            )

        else:
            raise

    return redirect(url_for("admin_dashboard"))


# =========================================================
# ADMIN ACCURACY MONITORING
# =========================================================

@app.route("/admin/accuracy")
@login_required
@admin_required
def admin_accuracy():

    period = request.args.get(
        "period",
        "MONTHLY",
    ).upper()

    if period not in ("DAILY", "WEEKLY", "MONTHLY"):
        period = "MONTHLY"

    project_ids = allowed_project_ids()

    monitoring = get_accuracy_monitoring(
        project_ids,
        period,
    )

    below_90 = [
        item for item in monitoring
        if any(
            metric["current"] is not None
            and metric["current"] < 90
            for metric in (
                item["internal"],
                item["external_wc"],
                item["external_woc"],
            )
        )
    ]

    return render_template(
        "admin_accuracy.html",
        period=period,
        monitoring=monitoring,
        below_90=below_90,
        format_accuracy=format_accuracy,
    )


# =========================================================
# ADMIN ERROR DETAILS
# =========================================================

@app.route("/admin/errors")
@login_required
@admin_required
def admin_errors():

    project_ids = allowed_project_ids()

    if not project_ids:

        rows = []

    else:

        placeholders = ",".join(
            ["?"] * len(project_ids)
        )

        with db() as con:

            rows = con.execute(
                f"""
                SELECT
                    e.*,
                    p.project_name,
                    r.verified_error,
                    r.biller_comment,
                    r.qc_explanation,
                    r.qc_name,
                    r.updated_by,
                    r.updated_at
                FROM error_records e
                LEFT JOIN projects p
                    ON p.id=e.project_id
                LEFT JOIN biller_responses r
                    ON r.record_key=e.record_key
                WHERE e.project_id IN ({placeholders})
                ORDER BY e.billed_date DESC, e.source_row DESC
                """,
                tuple(project_ids),
            ).fetchall()

    return render_template(
        "admin_errors.html",
        rows=rows,
    )


# =========================================================
# ADMIN EDIT ERROR RECORD
# =========================================================

@app.route(
    "/admin/error/<record_key>",
    methods=["GET", "POST"],
)
@login_required
@admin_required
def edit_error(record_key):

    with db() as con:

        record = con.execute(
            """
            SELECT *
            FROM error_records
            WHERE record_key=?
            """,
            (record_key,),
        ).fetchone()

        if not record:
            abort(404)

        if request.method == "POST":

            if request.form.get("csrf") != session.get("csrf"):
                abort(400)

            values = {
                field: request.form.get(field, "").strip()
                for field in [
                    "billed_date",
                    "week",
                    "source",
                    "audit_type",
                    "biller_name",
                    "username",
                    "id_number",
                    "auditor",
                    "pro",
                    "zb_field",
                    "error_details",
                ]
            }

            con.execute(
                """
                UPDATE error_records
                SET billed_date=?,
                    week=?,
                    source=?,
                    audit_type=?,
                    biller_name=?,
                    username=?,
                    id_number=?,
                    auditor=?,
                    pro=?,
                    zb_field=?,
                    error_details=?
                WHERE record_key=?
                """,
                (
                    values["billed_date"],
                    values["week"],
                    values["source"],
                    values["audit_type"],
                    values["biller_name"],
                    values["username"].upper(),
                    values["id_number"],
                    values["auditor"],
                    values["pro"],
                    values["zb_field"],
                    values["error_details"],
                    record_key,
                ),
            )

            log_action(
                session["username"],
                "EDIT_ERROR_RECORD",
                record_key,
            )

            flash(
                "Audit record updated.",
                "success",
            )

            return redirect(url_for("admin_errors"))

    return render_template(
        "edit_error.html",
        record=record,
    )

# =========================================================
# ADMIN - ERROR DETAILS
# =========================================================

@app.route("/admin/errors")
@login_required
@admin_required
def admin_errors():

    search = request.args.get("search", "").strip()
    project_id = request.args.get("project_id", "").strip()

    with db() as con:

        sql = """
            SELECT
                e.*,
                p.project_name,
                br.verified_error,
                br.biller_comment,
                br.qc_explanation,
                br.qc_name
            FROM error_records e
            LEFT JOIN projects p
                ON p.id = e.project_id
            LEFT JOIN biller_responses br
                ON br.record_key = e.record_key
            WHERE 1=1
        """

        params = []

        if project_id:
            sql += " AND e.project_id=?"
            params.append(project_id)

        if search:
            sql += """
                AND (
                    e.biller_name LIKE ?
                    OR e.username LIKE ?
                    OR e.pro LIKE ?
                    OR e.error_details LIKE ?
                )
            """
            term = f"%{search}%"
            params.extend([term, term, term, term])

        sql += """
            ORDER BY e.id DESC
            LIMIT 1000
        """

        records = con.execute(sql, params).fetchall()

        projects = con.execute(
            """
            SELECT id, project_name
            FROM projects
            WHERE active=1
            ORDER BY project_name
            """
        ).fetchall()

    return render_template(
        "errors.html",
        records=records,
        projects=projects,
        search=search,
        selected_project=project_id,
        role="admin",
    )


# =========================================================
# ADMIN - ACCOUNT MANAGEMENT
# =========================================================

@app.route("/admin/accounts")
@login_required
@admin_required
def admin_accounts():

    with db() as con:

        accounts = con.execute(
            """
            SELECT
                u.id,
                u.username,
                u.display_name,
                u.role,
                u.active,
                GROUP_CONCAT(p.project_name, ', ') AS projects
            FROM users u
            LEFT JOIN user_projects up
                ON up.user_id = u.id
            LEFT JOIN projects p
                ON p.id = up.project_id
            GROUP BY
                u.id,
                u.username,
                u.display_name,
                u.role,
                u.active
            ORDER BY u.role, u.username
            """
        ).fetchall()

        projects = con.execute(
            """
            SELECT id, project_name
            FROM projects
            WHERE active=1
            ORDER BY project_name
            """
        ).fetchall()

    return render_template(
        "accounts.html",
        accounts=accounts,
        projects=projects,
    )


@app.route("/admin/account/create", methods=["POST"])
@login_required
@admin_required
def create_account():

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    display_name = request.form.get(
        "display_name", ""
    ).strip()

    username = request.form.get(
        "username", ""
    ).strip().lower()

    password = request.form.get("password", "")

    role = request.form.get("role", "").strip().lower()

    project_ids = request.form.getlist("project_ids")

    if not display_name or not username or not password:
        flash("Please complete all required account fields.", "error")
        return redirect(url_for("admin_accounts"))

    if len(password) < 6:
        flash("Password must be at least 6 characters.", "error")
        return redirect(url_for("admin_accounts"))

    if role not in ("admin", "biller", "qc"):
        flash("Invalid account role.", "error")
        return redirect(url_for("admin_accounts"))

    if role != "admin" and not project_ids:
        flash("Please assign at least one project.", "error")
        return redirect(url_for("admin_accounts"))

    with db() as con:

        existing = con.execute(
            """
            SELECT id
            FROM users
            WHERE LOWER(username)=?
            """,
            (username,),
        ).fetchone()

        if existing:
            flash("Username already exists.", "error")
            return redirect(url_for("admin_accounts"))

        if role == "admin":
            project_ids = []

        valid_projects = []

        for project_id in project_ids:
            project = con.execute(
                """
                SELECT id
                FROM projects
                WHERE id=? AND active=1
                """,
                (project_id,),
            ).fetchone()

            if project:
                valid_projects.append(int(project_id))

        if role != "admin" and not valid_projects:
            flash("Please select valid active projects.", "error")
            return redirect(url_for("admin_accounts"))

        cursor = con.execute(
            """
            INSERT INTO users (
                username,
                password_hash,
                role,
                display_name,
                active,
                created_at
            )
            VALUES (?, ?, ?, ?, 1, CURRENT_TIMESTAMP)
            """,
            (
                username,
                generate_password_hash(password),
                role,
                display_name,
            ),
        )

        user_id = cursor.lastrowid

        for project_id in valid_projects:
            con.execute(
                """
                INSERT INTO user_projects (
                    user_id,
                    project_id
                )
                VALUES (?, ?)
                """,
                (user_id, project_id),
            )

        log_action(
            session["username"],
            "CREATE_ACCOUNT",
            f"{username} ({role})",
        )

    flash("Account created successfully.", "success")
    return redirect(url_for("admin_accounts"))


@app.route(
    "/admin/account/<int:user_id>/projects",
    methods=["POST"],
)
@login_required
@admin_required
def update_account_projects(user_id):

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    project_ids = request.form.getlist("project_ids")

    with db() as con:

        user = con.execute(
            """
            SELECT id, username, role
            FROM users
            WHERE id=?
            """,
            (user_id,),
        ).fetchone()

        if not user:
            abort(404)

        if user["role"] == "admin":
            flash("Admin accounts automatically have access to all projects.", "info")
            return redirect(url_for("admin_accounts"))

        valid_projects = []

        for project_id in project_ids:
            project = con.execute(
                """
                SELECT id
                FROM projects
                WHERE id=? AND active=1
                """,
                (project_id,),
            ).fetchone()

            if project:
                valid_projects.append(int(project_id))

        if not valid_projects:
            flash("An account must have at least one assigned project.", "error")
            return redirect(url_for("admin_accounts"))

        con.execute(
            """
            DELETE FROM user_projects
            WHERE user_id=?
            """,
            (user_id,),
        )

        for project_id in valid_projects:
            con.execute(
                """
                INSERT INTO user_projects (
                    user_id,
                    project_id
                )
                VALUES (?, ?)
                """,
                (user_id, project_id),
            )

        log_action(
            session["username"],
            "UPDATE_ACCOUNT_PROJECTS",
            user["username"],
        )

    flash("Account project assignments updated.", "success")
    return redirect(url_for("admin_accounts"))


@app.route(
    "/admin/account/<int:user_id>/toggle",
    methods=["POST"],
)
@login_required
@admin_required
def toggle_account(user_id):

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    with db() as con:

        user = con.execute(
            """
            SELECT id, username, role, active
            FROM users
            WHERE id=?
            """,
            (user_id,),
        ).fetchone()

        if not user:
            abort(404)

        if user_id == session.get("user_id"):
            flash("You cannot deactivate your own account.", "error")
            return redirect(url_for("admin_accounts"))

        new_status = 0 if user["active"] else 1

        if user["role"] == "admin" and new_status == 0:

            active_admins = con.execute(
                """
                SELECT COUNT(*) AS total
                FROM users
                WHERE role='admin'
                  AND active=1
                """
            ).fetchone()["total"]

            if active_admins <= 1:
                flash("You cannot deactivate the last active Admin.", "error")
                return redirect(url_for("admin_accounts"))

        con.execute(
            """
            UPDATE users
            SET active=?
            WHERE id=?
            """,
            (new_status, user_id),
        )

        log_action(
            session["username"],
            "TOGGLE_ACCOUNT",
            user["username"],
        )

    flash("Account status updated.", "success")
    return redirect(url_for("admin_accounts"))


# Legacy endpoint aliases, if older templates still use these names.
@app.route("/admin/create-biller", methods=["POST"])
@login_required
@admin_required
def create_biller():
    return create_account()


@app.route("/admin/create-qc", methods=["POST"])
@login_required
@admin_required
def create_qc():
    return create_account()


# =========================================================
# ADMIN - PROJECT MANAGEMENT
# =========================================================

@app.route("/admin/projects")
@login_required
@admin_required
def admin_projects():

    with db() as con:

        projects = con.execute(
            """
            SELECT
                p.*,
                GROUP_CONCAT(
                    u.display_name || ' (' || UPPER(u.role) || ')',
                    ', '
                ) AS assigned_accounts
            FROM projects p
            LEFT JOIN user_projects up
                ON up.project_id = p.id
            LEFT JOIN users u
                ON u.id = up.user_id
            GROUP BY p.id
            ORDER BY p.project_name
            """
        ).fetchall()

    return render_template(
        "projects.html",
        projects=projects,
    )


@app.route("/admin/project/create", methods=["POST"])
@login_required
@admin_required
def create_project():

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    project_name = request.form.get(
        "project_name", ""
    ).strip()

    name = request.form.get(
        "name", ""
    ).strip()

    role = request.form.get(
        "role", ""
    ).strip().lower()

    if not project_name or not name:
        flash("Project Name and Name are required.", "error")
        return redirect(url_for("admin_projects"))

    if role not in ("biller", "qc"):
        flash("Please select a valid Role.", "error")
        return redirect(url_for("admin_projects"))

    with db() as con:

        existing = con.execute(
            """
            SELECT id
            FROM projects
            WHERE LOWER(project_name)=LOWER(?)
            """,
            (project_name,),
        ).fetchone()

        if existing:
            flash("Project name already exists.", "error")
            return redirect(url_for("admin_projects"))

        con.execute(
            """
            INSERT INTO projects (
                project_name,
                name,
                role,
                active,
                created_at
            )
            VALUES (?, ?, ?, 1, CURRENT_TIMESTAMP)
            """,
            (project_name, name, role),
        )

        log_action(
            session["username"],
            "CREATE_PROJECT",
            project_name,
        )

    flash("Project created successfully.", "success")
    return redirect(url_for("admin_projects"))


@app.route(
    "/admin/project/<int:project_id>/edit",
    methods=["POST"],
)
@login_required
@admin_required
def edit_project(project_id):

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    project_name = request.form.get(
        "project_name", ""
    ).strip()

    name = request.form.get(
        "name", ""
    ).strip()

    role = request.form.get(
        "role", ""
    ).strip().lower()

    if not project_name or not name or role not in ("biller", "qc"):
        flash("Please provide valid project details.", "error")
        return redirect(url_for("admin_projects"))

    with db() as con:

        project = con.execute(
            """
            SELECT id
            FROM projects
            WHERE id=?
            """,
            (project_id,),
        ).fetchone()

        if not project:
            abort(404)

        duplicate = con.execute(
            """
            SELECT id
            FROM projects
            WHERE LOWER(project_name)=LOWER(?)
              AND id<>?
            """,
            (project_name, project_id),
        ).fetchone()

        if duplicate:
            flash("Another project already uses that name.", "error")
            return redirect(url_for("admin_projects"))

        con.execute(
            """
            UPDATE projects
            SET project_name=?,
                name=?,
                role=?
            WHERE id=?
            """,
            (project_name, name, role, project_id),
        )

        log_action(
            session["username"],
            "EDIT_PROJECT",
            project_name,
        )

    flash("Project updated successfully.", "success")
    return redirect(url_for("admin_projects"))


@app.route(
    "/admin/project/<int:project_id>/toggle",
    methods=["POST"],
)
@login_required
@admin_required
def toggle_project(project_id):

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    with db() as con:

        project = con.execute(
            """
            SELECT id, project_name, active
            FROM projects
            WHERE id=?
            """,
            (project_id,),
        ).fetchone()

        if not project:
            abort(404)

        new_status = 0 if project["active"] else 1

        con.execute(
            """
            UPDATE projects
            SET active=?
            WHERE id=?
            """,
            (new_status, project_id),
        )

        log_action(
            session["username"],
            "TOGGLE_PROJECT",
            project["project_name"],
        )

    flash("Project status updated.", "success")
    return redirect(url_for("admin_projects"))


# =========================================================
# BILLER - DASHBOARD AND ERROR RECORDS
# =========================================================

@app.route("/biller")
@login_required
@biller_required
def biller_dashboard():

    user_id = session["user_id"]
    username = session["username"]

    with db() as con:

        project_ids = get_user_project_ids(con, user_id)

        if not project_ids:
            records = []
        else:
            placeholders = ",".join("?" for _ in project_ids)

            records = con.execute(
                f"""
                SELECT
                    e.*,
                    p.project_name,
                    br.verified_error,
                    br.biller_comment,
                    br.qc_explanation,
                    br.qc_name
                FROM error_records e
                LEFT JOIN projects p
                    ON p.id=e.project_id
                LEFT JOIN biller_responses br
                    ON br.record_key=e.record_key
                WHERE UPPER(e.username)=UPPER(?)
                  AND e.project_id IN ({placeholders})
                ORDER BY e.id DESC
                """,
                [username] + project_ids,
            ).fetchall()

        projects = con.execute(
            """
            SELECT p.id, p.project_name
            FROM projects p
            INNER JOIN user_projects up
                ON up.project_id=p.id
            WHERE up.user_id=?
              AND p.active=1
            ORDER BY p.project_name
            """,
            (user_id,),
        ).fetchall()

    return render_template(
        "biller.html",
        records=records,
        projects=projects,
    )


@app.route("/biller/errors")
@login_required
@biller_required
def biller_errors():
    return biller_dashboard()


# =========================================================
# QC - ERROR RECORDS
# QC ONLY SEES VERIFIED ERROR = N
# =========================================================

@app.route("/qc")
@login_required
@qc_required
def qc_dashboard():

    user_id = session["user_id"]

    with db() as con:

        project_ids = get_user_project_ids(con, user_id)

        if not project_ids:
            records = []
        else:
            placeholders = ",".join("?" for _ in project_ids)

            records = con.execute(
                f"""
                SELECT
                    e.*,
                    p.project_name,
                    br.verified_error,
                    br.biller_comment,
                    br.qc_explanation,
                    br.qc_name
                FROM error_records e
                LEFT JOIN projects p
                    ON p.id=e.project_id
                INNER JOIN biller_responses br
                    ON br.record_key=e.record_key
                WHERE e.project_id IN ({placeholders})
                  AND UPPER(COALESCE(br.verified_error, ''))='N'
                ORDER BY e.id DESC
                """,
                project_ids,
            ).fetchall()

    return render_template(
        "qc.html",
        records=records,
    )


@app.route("/qc/errors")
@login_required
@qc_required
def qc_errors():
    return qc_dashboard()


# =========================================================
# SAVE BILLER / QC RESPONSE
# =========================================================

@app.route(
    "/response/<record_key>",
    methods=["POST"],
)
@login_required
def save_response(record_key):

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    role = session.get("role", "").lower()
    user_id = session["user_id"]

    with db() as con:

        record = con.execute(
            """
            SELECT *
            FROM error_records
            WHERE record_key=?
            """,
            (record_key,),
        ).fetchone()

        if not record:
            abort(404)

        if role == "biller":

            project_ids = get_user_project_ids(con, user_id)

            if (
                record["project_id"] not in project_ids
                or str(record["username"]).lower()
                != str(session["username"]).lower()
            ):
                abort(403)

            verified_error = request.form.get(
                "verified_error", ""
            ).strip().upper()

            biller_comment = request.form.get(
                "biller_comment", ""
            ).strip()

            if verified_error not in ("Y", "N"):
                flash("Please select Verified Error Y or N.", "error")
                return redirect(request.referrer or url_for("biller_dashboard"))

            existing = con.execute(
                """
                SELECT record_key
                FROM biller_responses
                WHERE record_key=?
                """,
                (record_key,),
            ).fetchone()

            if existing:
                con.execute(
                    """
                    UPDATE biller_responses
                    SET verified_error=?,
                        biller_comment=?
                    WHERE record_key=?
                    """,
                    (
                        verified_error,
                        biller_comment,
                        record_key,
                    ),
                )
            else:
                con.execute(
                    """
                    INSERT INTO biller_responses (
                        record_key,
                        verified_error,
                        biller_comment
                    )
                    VALUES (?, ?, ?)
                    """,
                    (
                        record_key,
                        verified_error,
                        biller_comment,
                    ),
                )

            log_action(
                session["username"],
                "BILLER_RESPONSE",
                record_key,
            )

            flash("Your response has been saved.", "success")

        elif role == "qc":

            project_ids = get_user_project_ids(con, user_id)

            if record["project_id"] not in project_ids:
                abort(403)

            response = con.execute(
                """
                SELECT verified_error
                FROM biller_responses
                WHERE record_key=?
                """,
                (record_key,),
            ).fetchone()

            if not response or str(response["verified_error"]).upper() != "N":
                abort(403)

            qc_explanation = request.form.get(
                "qc_explanation", ""
            ).strip()

            if not qc_explanation:
                flash("QC Explanation is required.", "error")
                return redirect(request.referrer or url_for("qc_dashboard"))

            qc_name = session.get("display_name") or session["username"]

            con.execute(
                """
                UPDATE biller_responses
                SET qc_explanation=?,
                    qc_name=?
                WHERE record_key=?
                """,
                (
                    qc_explanation,
                    qc_name,
                    record_key,
                ),
            )

            log_action(
                session["username"],
                "QC_EXPLANATION",
                record_key,
            )

            flash("QC Explanation has been saved.", "success")

        else:
            abort(403)

    if role == "qc":
        return redirect(url_for("qc_dashboard"))

    return redirect(url_for("biller_dashboard"))


# =========================================================
# ADMIN - EXPORT RESPONSES
# =========================================================

@app.route("/admin/export/responses")
@login_required
@admin_required
def export_responses():

    with db() as con:

        rows = con.execute(
            """
            SELECT
                e.record_key,
                p.project_name,
                e.billed_date,
                e.week,
                e.source,
                e.audit_type,
                e.biller_name,
                e.username,
                e.id_number,
                e.auditor,
                e.pro,
                e.zb_field,
                e.error_details,
                br.verified_error,
                br.biller_comment,
                br.qc_name,
                br.qc_explanation
            FROM error_records e
            LEFT JOIN projects p
                ON p.id=e.project_id
            LEFT JOIN biller_responses br
                ON br.record_key=e.record_key
            ORDER BY e.id
            """
        ).fetchall()

    df = pd.DataFrame([dict(row) for row in rows])

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Error Responses",
        )

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="Error_Responses.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


# =========================================================
# ADMIN - ACTIVITY LOGS
# =========================================================

@app.route("/admin/logs")
@login_required
@admin_required
def admin_logs():

    with db() as con:

        logs = con.execute(
            """
            SELECT *
            FROM activity_logs
            ORDER BY id DESC
            LIMIT 1000
            """
        ).fetchall()

    return render_template(
        "logs.html",
        logs=logs,
    )


# =========================================================
# ERROR HANDLERS
# =========================================================

@app.errorhandler(403)
def forbidden(error):
    return render_template(
        "error.html",
        message="You do not have permission to access this page.",
    ), 403


@app.errorhandler(404)
def not_found(error):
    return render_template(
        "error.html",
        message="The requested page or record was not found.",
    ), 404


# =========================================================
# APPLICATION STARTUP
# =========================================================

with app.app_context():
    init_db()


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=5001,
        debug=False,
    )

