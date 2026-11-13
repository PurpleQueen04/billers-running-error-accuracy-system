
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
# DATABASE ROLE MIGRATION
# =========================================================

def migrate_user_roles(con, is_postgres=False):
    """
    Add QC to the allowed user roles.

    PostgreSQL:
        Replace the existing role CHECK constraint.

    SQLite:
        Rebuild only the users table while preserving all
        existing user IDs, usernames, password hashes,
        roles, display names, statuses and creation dates.

    No error records or biller responses are deleted.
    """

    if is_postgres:

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
            constraint_name = item["conname"]

            safe_name = constraint_name.replace('"', '""')

            con.execute(
                f'ALTER TABLE users DROP CONSTRAINT "{safe_name}"'
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

    if "'qc'" in table_sql or '"qc"' in table_sql:
        return

    # This table has no foreign-key references from the
    # application's existing schema.
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

    if DATABASE_URL:

        ddl = """
        CREATE TABLE IF NOT EXISTS users (
          id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
          username TEXT UNIQUE NOT NULL,
          password_hash TEXT NOT NULL,
          role TEXT NOT NULL CHECK(role IN ('admin','biller')),
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
        """

        with db() as con:

            con.executescript(ddl)

            # Add QC role without deleting existing accounts.
            migrate_user_roles(
                con,
                is_postgres=True,
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

    # SQLITE DATABASE

    with db() as con:

        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS users(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              username TEXT UNIQUE NOT NULL,
              password_hash TEXT NOT NULL,
              role TEXT NOT NULL
                CHECK(role IN ('admin','biller')),
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
            """
        )

        # Add QC role without deleting existing accounts.
        migrate_user_roles(
            con,
            is_postgres=False,
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

        df = pd.read_csv(
            io.BytesIO(raw),
            dtype=object,
            encoding="utf-8-sig",
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
# IMPORT WORKBOOK
# =========================================================

def import_workbook(sheets, filename, imported_by):
    """
    Non-destructive import.

    Existing records not included in a new upload are retained.
    Existing Biller and QC responses are preserved.
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

            seen = {}

            for idx, row in error_df.iterrows():

                billed_date = parse_date(
                    val(row, cols["date"])
                )

                username = val(
                    row,
                    cols["user"],
                ).upper()

                pro = val(row, cols["pro"])

                if not username or not pro:
                    continue

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

                base = tuple(fields)

                seen[base] = seen.get(base, 0) + 1

                key = fingerprint(
                    fields,
                    seen[base],
                )

                con.execute(
                    """
                    INSERT INTO error_records
                        (record_key,billed_date,week,source,audit_type,
                         biller_name,username,id_number,auditor,pro,
                         zb_field,error_details,source_row,imported_at)
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

                    con.execute(
                        """
                        INSERT INTO accuracy_records
                            (sheet,username,biller_name,month,weekly,
                             record_date,accuracy_type,accuracy_value,
                             bills_processed,audited,errors,source_row,
                             imported_at)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)

                        ON CONFLICT
                            (sheet,username,record_date,accuracy_type,source_row)
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

    return [
        row["month"]
        for row in rows
    ]


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


def get_accuracy_rows(username, sheet, month):

    with db() as con:

        rows = con.execute(
            """
            SELECT *
            FROM accuracy_records
            WHERE UPPER(username)=?
              AND sheet=?
              AND month=?
            ORDER BY record_date, source_row
            """,
            (
                username.upper(),
                sheet,
                month,
            ),
        ).fetchall()

    return rows


def group_accuracy_rows(rows, period):
    """
    Aggregate records into one weighted-accuracy row per date or week.
    """

    groups = {}

    for row in rows:

        if period == "WEEKLY":

            label = (
                row["weekly"]
                or week_for_billed_date(row["record_date"])
                or "Undated"
            )

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
            SELECT id, username, display_name, active
            FROM users
            WHERE role='qc'
            ORDER BY username
            """
        ).fetchall()

    return render_template(
        "admin.html",
        counts=counts,
        history=history,
        **{
            "billers": billers,
            "qc_users": qc_users,
        },
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
# ADMIN EDIT ERROR RECORD
# =========================================================

@app.route("/admin/error/<record_key>", methods=["GET", "POST"])
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

            return redirect(url_for("admin_dashboard"))

    return render_template(
        "edit_error.html",
        record=record,
    )


@app.route("/admin/errors")
@login_required
@admin_required
def admin_errors():

    with db() as con:

        rows = con.execute(
            """
            SELECT *
            FROM error_records
            ORDER BY source_row DESC
            LIMIT 250
            """
        ).fetchall()

    return render_template(
        "admin_errors.html",
        rows=rows,
    )


# =========================================================
# ADMIN CREATE BILLER
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

    if not username or len(password) < 10:

        flash(
            "Username is required and password must be "
            "at least 10 characters.",
            "error",
        )

    else:

        try:

            with db() as con:

                con.execute(
                    """
                    INSERT INTO users
                        (username,password_hash,role,
                         display_name,created_at)
                    VALUES(?,?,?,?,?)
                    """,
                    (
                        username,
                        generate_password_hash(password),
                        "biller",
                        display or username,
                        datetime.now().isoformat(timespec="seconds"),
                    ),
                )

            log_action(
                session["username"],
                "CREATE_BILLER",
                username,
            )

            flash(
                f"Biller account {username} created.",
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
                    "That username already exists.",
                    "error",
                )

            else:
                raise

    return redirect(url_for("admin_dashboard"))


# =========================================================
# ADMIN TOGGLE BILLER
# =========================================================

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


# =========================================================
# ADMIN CREATE QC ACCOUNT
# =========================================================

@app.route("/admin/qc-users", methods=["POST"])
@login_required
@admin_required
def create_qc_user():

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

    if not username or not display or len(password) < 10:

        flash(
            "QC username, QC name, and a password of at least "
            "10 characters are required.",
            "error",
        )

        return redirect(url_for("admin_dashboard"))

    try:

        with db() as con:

            con.execute(
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
                "That username already exists.",
                "error",
            )

        else:
            raise

    return redirect(url_for("admin_dashboard"))


# =========================================================
# ADMIN ENABLE / DISABLE QC ACCOUNT
# =========================================================

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


# =========================================================
# BILLER DASHBOARD
# =========================================================

@app.route("/biller")
@login_required
def biller_dashboard():

    if session.get("role") != "biller":
        abort(403)

    username = session["username"].upper()

    months = get_biller_months(username)

    selected_month = selected_month_value(username)

    with db() as con:

        accuracy = con.execute(
            """
            SELECT *
            FROM accuracy_records
            WHERE UPPER(username)=?
            ORDER BY record_date DESC, source_row DESC
            """,
            (username,),
        ).fetchall()

        response_counts = con.execute(
            """
            SELECT
                COUNT(*) AS total,

                COALESCE(SUM(
                    CASE WHEN r.verified_error='Y'
                    THEN 1 ELSE 0 END
                ),0) AS yes_count,

                COALESCE(SUM(
                    CASE WHEN r.verified_error='N'
                    THEN 1 ELSE 0 END
                ),0) AS no_count

            FROM error_records e

            LEFT JOIN biller_responses r
                USING(record_key)

            WHERE UPPER(e.username)=?
            """,
            (username,),
        ).fetchone()

    selected_rows = [
        row for row in accuracy
        if row["month"] == selected_month
    ]

    internal_summary = monthly_accuracy_summary(
        selected_rows,
        "INTERNAL",
    )

    external_wc_summary = monthly_accuracy_summary(
        selected_rows,
        "EXTERNAL",
        "PER AUDITED ACCURACY",
    )

    external_woc_summary = monthly_accuracy_summary(
        selected_rows,
        "EXTERNAL",
        "PER AUDITED ACCURACY - WOC",
    )

    return render_template(
        "biller.html",
        selected_month=selected_month,
        months=months,
        years=month_year_options(months),
        internal_summary=internal_summary,
        external_wc_summary=external_wc_summary,
        external_woc_summary=external_woc_summary,
        response_counts=response_counts,
        format_accuracy=format_accuracy,
    )


# =========================================================
# BILLER INTERNAL ACCURACY
# =========================================================

@app.route("/biller/internal")
@login_required
def biller_internal_accuracy():

    if session.get("role") != "biller":
        abort(403)

    username = session["username"].upper()

    months = get_biller_months(username)

    selected_month = selected_month_value(username)

    period = request.args.get(
        "period",
        "DAILY",
    ).upper()

    if period not in ("DAILY", "WEEKLY"):
        period = "DAILY"

    rows = get_accuracy_rows(
        username,
        "INTERNAL",
        selected_month,
    )

    summary = monthly_accuracy_summary(
        rows,
        "INTERNAL",
    )

    grouped = group_accuracy_rows(
        rows,
        period,
    )

    return render_template(
        "internal_accuracy.html",
        selected_month=selected_month,
        months=months,
        years=month_year_options(months),
        period=period,
        summary=summary,
        rows=grouped,
        format_accuracy=format_accuracy,
    )


# =========================================================
# BILLER EXTERNAL ACCURACY
# =========================================================

@app.route("/biller/external")
@login_required
def biller_external_accuracy():

    if session.get("role") != "biller":
        abort(403)

    username = session["username"].upper()

    months = get_biller_months(username)

    selected_month = selected_month_value(username)

    period = request.args.get(
        "period",
        "DAILY",
    ).upper()

    if period not in ("DAILY", "WEEKLY"):
        period = "DAILY"

    rows = get_accuracy_rows(
        username,
        "EXTERNAL",
        selected_month,
    )

    wc_summary = monthly_accuracy_summary(
        rows,
        "EXTERNAL",
        "PER AUDITED ACCURACY",
    )

    woc_summary = monthly_accuracy_summary(
        rows,
        "EXTERNAL",
        "PER AUDITED ACCURACY - WOC",
    )

    grouped = group_accuracy_rows(
        rows,
        period,
    )

    return render_template(
        "external_accuracy.html",
        selected_month=selected_month,
        months=months,
        years=month_year_options(months),
        period=period,
        wc_summary=wc_summary,
        woc_summary=woc_summary,
        rows=grouped,
        format_accuracy=format_accuracy,
    )


# =========================================================
# BILLER ERROR RECORDS
# =========================================================

@app.route("/biller/errors")
@login_required
def biller_errors():

    if session.get("role") != "biller":
        abort(403)

    username = session["username"].upper()

    with db() as con:

        errors = con.execute(
            """
            SELECT
                e.*,
                r.verified_error,
                r.biller_comment,
                r.qc_explanation,
                r.qc_name,
                r.updated_at

            FROM error_records e

            LEFT JOIN biller_responses r
                USING(record_key)

            WHERE UPPER(e.username)=?

            ORDER BY e.billed_date DESC, e.source_row DESC
            """,
            (username,),
        ).fetchall()

        response_counts = con.execute(
            """
            SELECT
                COUNT(*) AS total,

                COALESCE(SUM(
                    CASE WHEN r.verified_error='Y'
                    THEN 1 ELSE 0 END
                ),0) AS yes_count,

                COALESCE(SUM(
                    CASE WHEN r.verified_error='N'
                    THEN 1 ELSE 0 END
                ),0) AS no_count

            FROM error_records e

            LEFT JOIN biller_responses r
                USING(record_key)

            WHERE UPPER(e.username)=?
            """,
            (username,),
        ).fetchone()

    return render_template(
        "errors.html",
        errors=errors,
        response_counts=response_counts,
    )


# =========================================================
# QC ERROR RECORDS
# =========================================================

@app.route("/qc/errors")
@login_required
@role_required("qc")
def qc_errors():

    with db() as con:

        errors = con.execute(
            """
            SELECT
                e.*,
                r.verified_error,
                r.biller_comment,
                r.qc_explanation,
                r.qc_name,
                r.updated_at

            FROM error_records e

            INNER JOIN biller_responses r
                ON r.record_key = e.record_key

            WHERE r.verified_error='N'

            ORDER BY e.billed_date DESC, e.source_row DESC
            """
        ).fetchall()

    return render_template(
        "errors.html",
        errors=errors,
    )


# =========================================================
# SAVE ROLE-BASED RESPONSE
# =========================================================

@app.route(
    "/biller/respond/<record_key>",
    methods=["POST"],
)
@login_required
def save_response(record_key):

    role = (session.get("role") or "").lower()

    if role not in ("biller", "qc", "admin"):
        abort(403)

    if request.form.get("csrf") != session.get("csrf"):
        abort(400)

    username = session.get("username", "")

    now = datetime.now().isoformat(timespec="seconds")

    # -----------------------------------------------------
    # BILLER
    # -----------------------------------------------------

    if role == "biller":

        with db() as con:

            record = con.execute(
                """
                SELECT record_key
                FROM error_records
                WHERE record_key=?
                  AND UPPER(username)=?
                """,
                (
                    record_key,
                    username.upper(),
                ),
            ).fetchone()

            if not record:
                abort(404)

            verified = request.form.get(
                "verified_error",
                "",
            ).strip().upper()

            if verified not in ("", "Y", "N"):
                abort(400)

            biller_comment = request.form.get(
                "biller_comment",
                "",
            ).strip()

            # IMPORTANT:
            # Biller cannot update QC explanation or QC name.
            # Existing QC values remain unchanged.

            con.execute(
                """
                INSERT INTO biller_responses
                    (record_key,verified_error,biller_comment,
                     updated_by,updated_at)
                VALUES(?,?,?,?,?)

                ON CONFLICT(record_key) DO UPDATE SET
                    verified_error=excluded.verified_error,
                    biller_comment=excluded.biller_comment,
                    updated_by=excluded.updated_by,
                    updated_at=excluded.updated_at
                """,
                (
                    record_key,
                    verified or None,
                    biller_comment,
                    username,
                    now,
                ),
            )

        log_action(
            username,
            "SAVE_BILLER_RESPONSE",
            record_key,
        )

        flash(
            "Your response has been saved.",
            "success",
        )

        return redirect(
            url_for("biller_errors")
            + "#record-"
            + record_key
        )

    # -----------------------------------------------------
    # QC
    # -----------------------------------------------------

    if role == "qc":

        explanation = request.form.get(
            "qc_explanation",
            "",
        ).strip()

        if not explanation:

            flash(
                "QC Explanation is required.",
                "error",
            )

            return redirect(url_for("qc_errors"))

        with db() as con:

            # QC can only update a record whose current
            # Verified Error value is N.

            record = con.execute(
                """
                SELECT
                    e.record_key,
                    r.verified_error

                FROM error_records e

                INNER JOIN biller_responses r
                    ON r.record_key=e.record_key

                WHERE e.record_key=?
                  AND r.verified_error='N'
                """,
                (record_key,),
            ).fetchone()

            if not record:
                abort(403)

            # QC cannot change:
            # - Verified Error
            # - Biller Comment
            #
            # QC Name is always assigned from the logged-in account.

            qc_name = (
                session.get("display_name")
                or username
            )

            con.execute(
                """
                UPDATE biller_responses

                SET qc_explanation=?,
                    qc_name=?,
                    updated_by=?,
                    updated_at=?

                WHERE record_key=?
                  AND verified_error='N'
                """,
                (
                    explanation,
                    qc_name,
                    username,
                    now,
                    record_key,
                ),
            )

        log_action(
            username,
            "SAVE_QC_EXPLANATION",
            record_key,
        )

        flash(
            "QC Explanation has been submitted.",
            "success",
        )

        return redirect(
            url_for("qc_errors")
            + "#record-"
            + record_key
        )

    # -----------------------------------------------------
    # ADMIN
    # -----------------------------------------------------

    if role == "admin":

        with db() as con:

            record = con.execute(
                """
                SELECT record_key
                FROM error_records
                WHERE record_key=?
                """,
                (record_key,),
            ).fetchone()

            if not record:
                abort(404)

            verified = request.form.get(
                "verified_error",
                "",
            ).strip().upper()

            if verified not in ("", "Y", "N"):
                abort(400)

            con.execute(
                """
                INSERT INTO biller_responses
                    (record_key,verified_error,biller_comment,
                     qc_explanation,qc_name,updated_by,updated_at)
                VALUES(?,?,?,?,?,?,?)

                ON CONFLICT(record_key) DO UPDATE SET
                    verified_error=excluded.verified_error,
                    biller_comment=excluded.biller_comment,
                    qc_explanation=excluded.qc_explanation,
                    qc_name=excluded.qc_name,
                    updated_by=excluded.updated_by,
                    updated_at=excluded.updated_at
                """,
                (
                    record_key,
                    verified or None,
                    request.form.get(
                        "biller_comment",
                        "",
                    ).strip(),
                    request.form.get(
                        "qc_explanation",
                        "",
                    ).strip(),
                    request.form.get(
                        "qc_name",
                        "",
                    ).strip(),
                    username,
                    now,
                ),
            )

        log_action(
            username,
            "ADMIN_EDIT_RESPONSE",
            record_key,
        )

        flash(
            "Response updated successfully.",
            "success",
        )

        return redirect(url_for("admin_dashboard"))

    abort(403)


# =========================================================
# ADMIN EXPORT RESPONSES
# =========================================================

@app.route("/admin/export-responses")
@login_required
@admin_required
def export_responses():

    with db() as con:

        export_sql = """
          SELECT
            e.billed_date AS DATE,
            e.week AS WEEK,
            e.source AS SOURCE_OF_ERROR,
            e.username AS USER,
            e.pro AS PRO,
            e.zb_field AS ZB_FIELD,
            e.error_details AS ERROR_DETAILS,
            e.auditor AS AUDITOR,
            e.audit_type AS AUDIT_TYPE,
            r.verified_error AS VERIFIED_ERROR,
            r.biller_comment AS BILLER_COMMENT,
            r.qc_explanation AS QC_EXPLANATION,
            r.qc_name AS QC_NAME,
            r.updated_by AS UPDATED_BY,
            r.updated_at AS UPDATED_AT

          FROM error_records e

          LEFT JOIN biller_responses r
            USING(record_key)

          ORDER BY e.username,e.billed_date,e.pro
        """

        cursor = con.execute(export_sql)

        rows = cursor.fetchall()

        columns = [
            d.name if hasattr(d, "name") else d[0]
            for d in cursor.description
        ]

        df = pd.DataFrame(
            [
                dict(row)
                if not isinstance(row, dict)
                else row
                for row in rows
            ],
            columns=columns,
        )

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Biller Responses",
        )

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="Biller_Running_Error_Responses.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


# =========================================================
# ADMIN ACTIVITY LOGS
# =========================================================

@app.route("/admin/logs")
@login_required
@admin_required
def logs():

    with db() as con:

        rows = con.execute(
            """
            SELECT *
            FROM activity_logs
            ORDER BY id DESC
            LIMIT 200
            """
        ).fetchall()

    return render_template(
        "logs.html",
        rows=rows,
    )


# =========================================================
# ERROR HANDLERS
# =========================================================

@app.errorhandler(403)
def forbidden(_):

    return render_template(
        "error.html",
        code=403,
        message="You do not have permission to access this page.",
    ), 403


@app.errorhandler(404)
def not_found(_):

    return render_template(
        "error.html",
        code=404,
        message="The requested record or page could not be found.",
    ), 404


# =========================================================
# STARTUP
# =========================================================

init_db()


if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5001,
        debug=False,
    )
