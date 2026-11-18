
import os
import io
import json
import csv
import hashlib
import sqlite3
import secrets
import zipfile
import xml.etree.ElementTree as ET

from datetime import datetime, date, timedelta
from functools import wraps

import pandas as pd

from flask import (
    Flask,
    render_template,
    request,
    Response,
    redirect,
    url_for,
    session,
    flash,
    send_file,
    abort,
    g
)

from werkzeug.security import (
    generate_password_hash,
    check_password_hash
)

from werkzeug.utils import secure_filename
from werkzeug.exceptions import HTTPException

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

UPLOAD_FOLDER = os.path.join(APP_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

# -------------------------------------------------
# REQUIRE POSTGRESQL IN PRODUCTION
# -------------------------------------------------

if (
    os.environ.get("RENDER", "").lower() == "true"
    and not DATABASE_URL
):
    raise RuntimeError(
        "DATABASE_URL is required in production. "
        "SQLite fallback is not allowed on Render."
    )
    
app = Flask(__name__)

SECRET_KEY = os.environ.get("SECRET_KEY")

if not SECRET_KEY and os.environ.get("RENDER", "").lower() == "true":
    raise RuntimeError(
        "SECRET_KEY environment variable is required in production."
    )

app.secret_key = SECRET_KEY or secrets.token_hex(32)

app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=(
        os.environ.get("RENDER", "").lower() == "true"
    )
)

ALLOWED_EXTENSIONS = {
    ".ods",
    ".xlsx",
    ".xls",
    ".csv"
}

INTERNAL_ACCURACY = "ACCURACY"

EXTERNAL_ACCURACY = [
    "PER AUDITED ACCURACY",
    "PER AUDITED ACCURACY - WOC"
]

ERROR_ALIASES = {
    "billed_date": [
        "BILLED DATE"
    ],
    "source": [
        "SOURCE",
        "ORIGIN"
    ],
    "audit_type": [
        "AUDIT TYPE",
        "TYPE"
    ],
    "biller_name": [
        "BILLER NAME",
        "BILLER",
        "NAME"
    ],
    "username": [
        "USERNAME",
        "USER",
        "USER NAME"
    ],
    "id_number": [
        "ID NUMBER",
        "ID",
        "EMPLOYEE ID"
    ],
    "auditor": [
        "AUDITOR",
        "AUDIT BY"
    ],
    "pro": [
        "PRO",
        "PRO NUMBER",
        "PRO#"
    ],
    "zb_field": [
        "ZB FIELD",
        "FIELD"
    ],
    "error_details": [
        "ERROR DETAILS",
        "ERROR",
        "DETAILS"
    ],
    "project": [
        "PROJECT",
        "PROJECT NAME"
    ]
}


# =========================================================
# DATABASE CONNECTION
# Supports SQLite and PostgreSQL / Supabase
# =========================================================

def is_postgres():
    return bool(DATABASE_URL)



def convert_placeholders(sql):
    """
    Convert SQLite-style ? placeholders to PostgreSQL %s.

    Ignore question marks inside:
    - Single-quoted strings
    - Double-quoted identifiers
    - SQL line comments
    - SQL block comments
    """

    if not is_postgres():
        return sql

    result = []
    i = 0
    length = len(sql)

    in_single_quote = False
    in_double_quote = False
    in_line_comment = False
    in_block_comment = False

    while i < length:

        char = sql[i]
        next_char = sql[i + 1] if i + 1 < length else ""

        if in_line_comment:
            result.append(char)

            if char == "\n":
                in_line_comment = False

            i += 1
            continue

        if in_block_comment:
            result.append(char)

            if char == "*" and next_char == "/":
                result.append(next_char)
                i += 2
                in_block_comment = False
                continue

            i += 1
            continue

        if in_single_quote:
            result.append(char)

            if char == "'" and next_char == "'":
                result.append(next_char)
                i += 2
                continue

            if char == "'":
                in_single_quote = False

            i += 1
            continue

        if in_double_quote:
            result.append(char)

            if char == '"' and next_char == '"':
                result.append(next_char)
                i += 2
                continue

            if char == '"':
                in_double_quote = False

            i += 1
            continue

        if char == "'":
            in_single_quote = True
            result.append(char)

        elif char == '"':
            in_double_quote = True
            result.append(char)

        elif char == "-" and next_char == "-":
            in_line_comment = True
            result.append(char)
            result.append(next_char)
            i += 2
            continue

        elif char == "/" and next_char == "*":
            in_block_comment = True
            result.append(char)
            result.append(next_char)
            i += 2
            continue

        elif char == "?":
            result.append("%s")

        else:
            result.append(char)

        i += 1

    return "".join(result)


class DatabaseCursor:

    def __init__(self, cursor, postgres=False):
        self.cursor = cursor
        self.postgres = postgres

    def execute(self, sql, params=None):
        sql = convert_placeholders(sql)

        if params is None:
            self.cursor.execute(sql)
        else:
            self.cursor.execute(sql, params)

        return self

    def executemany(self, sql, params):
        sql = convert_placeholders(sql)
        self.cursor.executemany(sql, params)
        return self

    def fetchone(self):
        return self.cursor.fetchone()

    def fetchall(self):
        return self.cursor.fetchall()

    @property
    def rowcount(self):
        return self.cursor.rowcount

    @property
    def lastrowid(self):
        if self.postgres:
            return None

        return getattr(
            self.cursor,
            "lastrowid",
            None
        )
        
    def close(self):
        self.cursor.close()


class DatabaseConnection:

    def __init__(self, connection, postgres=False):
        self.connection = connection
        self.postgres = postgres

    def execute(self, sql, params=None):
        cursor = self.connection.cursor()

        wrapped = DatabaseCursor(
            cursor,
            postgres=self.postgres
        )

        wrapped.execute(sql, params)

        return wrapped

    def executemany(self, sql, params):
        cursor = self.connection.cursor()

        wrapped = DatabaseCursor(
            cursor,
            postgres=self.postgres
        )

        wrapped.executemany(sql, params)

        return wrapped

    def commit(self):
        self.connection.commit()

    def rollback(self):
        self.connection.rollback()

    def close(self):
        self.connection.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):

        try:
            if exc_type is None:
                self.commit()
            else:
                self.rollback()
    
        finally:
            self.close()
    
        return False


def db():
    """
    Return a database connection.

    Supabase uses PostgreSQL.
    Local development uses SQLite.
    """

    if is_postgres():

        if psycopg is None:
            raise RuntimeError(
                "PostgreSQL driver is missing. "
                "Install psycopg[binary]."
            )

        connection = psycopg.connect(
            DATABASE_URL,
            row_factory=dict_row,
            connect_timeout=15
        )

        return DatabaseConnection(
            connection,
            postgres=True
        )

    connection = sqlite3.connect(
        DB_PATH,
        timeout=30
    )

    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")

    return DatabaseConnection(
        connection,
        postgres=False
    )


# =========================================================
# DATABASE SCHEMA HELPERS
# =========================================================

def table_exists(con, table_name):

    if is_postgres():

        row = con.execute(
            """
            SELECT EXISTS (
                SELECT 1
                FROM information_schema.tables
                WHERE table_schema = 'public'
                AND table_name = ?
            ) AS found
            """,
            (table_name,)
        ).fetchone()

        return bool(row["found"])

    row = con.execute(
        """
        SELECT name
        FROM sqlite_master
        WHERE type = 'table'
        AND name = ?
        """,
        (table_name,)
    ).fetchone()

    return row is not None

def table_columns(con, table_name):

    if not table_name or not table_name.isidentifier():
        raise ValueError("Invalid table name.")

    if not table_exists(con, table_name):
        return set()

    if is_postgres():

        rows = con.execute(
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'public'
            AND table_name = ?
            """,
            (table_name,)
        ).fetchall()

        return {
            row["column_name"]
            for row in rows
        }

    rows = con.execute(
        f'PRAGMA table_info("{table_name}")'
    ).fetchall()

    return {
        row["name"]
        for row in rows
    }

def add_column_if_missing(
    con,
    table_name,
    column_name,
    column_definition
):

    allowed_tables = {
        "users",
        "projects",
        "user_projects",
        "error_records",
        "accuracy_records",
        "biller_responses",
        "import_history",
        "activity_logs",
        "schema_migrations"
    }

    if table_name not in allowed_tables:
        raise ValueError(
            f"Invalid database table: {table_name}"
        )

    if not table_name.isidentifier():
        raise ValueError("Invalid table name.")

    if not column_name.isidentifier():
        raise ValueError("Invalid column name.")

    if not column_definition:
        raise ValueError("Column definition is required.")

    columns = table_columns(con, table_name)

    if column_name in columns:
        return

    con.execute(
        f'''
        ALTER TABLE "{table_name}"
        ADD COLUMN "{column_name}" {column_definition}
        '''
    )

# =========================================================
# GENERAL HELPERS
# =========================================================

def normalize_text(value):

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    return " ".join(
        str(value).strip().split()
    )


def norm_header(value):

    return normalize_text(value).upper()


def number_value(value):

    if value is None:
        return None

    text_value = normalize_text(value)

    if not text_value:
        return None

    text_value = (
        text_value
        .replace(",", "")
        .replace("%", "")
        .replace("₱", "")
        .replace("$", "")
    )

    try:
        return float(text_value)
    except (TypeError, ValueError):
        return None


def normalize_accuracy(value):

    result = number_value(value)

    if result is not None and 0 <= result <= 1:
        result *= 100

    return result


def parse_date(value):
    """Parse common workbook date values without expensive pandas inference."""
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    # pandas Timestamp and numpy date-like values expose to_pydatetime().
    to_python = getattr(value, "to_pydatetime", None)
    if callable(to_python):
        try:
            converted = to_python()
            if isinstance(converted, datetime):
                return converted.date()
            if isinstance(converted, date):
                return converted
        except (TypeError, ValueError, OverflowError):
            pass

    # Excel serial date values may appear when a column was not auto-converted.
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            if 1 <= float(value) <= 2958465:
                return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
        except (TypeError, ValueError, OverflowError):
            return None

    text_value = normalize_text(value)
    if not text_value:
        return None

    # ISO strings, including timestamps and timezone suffixes.
    try:
        return datetime.fromisoformat(text_value.replace("Z", "+00:00")).date()
    except (TypeError, ValueError, OverflowError):
        pass

    formats = (
        "%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y", "%d/%m/%y",
        "%Y/%m/%d", "%m-%d-%Y", "%m-%d-%y", "%d-%m-%Y",
        "%d-%m-%y", "%d-%b-%Y", "%d-%B-%Y", "%b %d, %Y",
        "%B %d, %Y", "%b %d %Y", "%B %d %Y", "%b-%d-%Y",
        "%B-%d-%Y", "%d %b %Y", "%d %B %Y"
    )
    for fmt in formats:
        try:
            parsed = datetime.strptime(text_value, fmt).date()
            if 1900 <= parsed.year <= 9999:
                return parsed
        except (TypeError, ValueError, OverflowError):
            continue
    return None


def month_from_date(value):

    parsed = parse_date(value)

    if not parsed:
        return ""

    return parsed.strftime("%Y-%m")


def week_for_billed_date(value):

    parsed = parse_date(value)

    if not parsed:
        return ""

    return (
        f"{parsed.isocalendar().year}-"
        f"W{parsed.isocalendar().week:02d}"
    )


def allowed_file(filename):

    if not filename or "." not in filename:
        return False

    extension = os.path.splitext(filename)[1].lower()

    return extension in ALLOWED_EXTENSIONS


def fingerprint(parts, occurrence=0):

    payload = json.dumps(
        [
            normalize_text(part)
            for part in parts
        ] + [occurrence],
        ensure_ascii=False,
        default=str
    )

    return hashlib.sha256(
        payload.encode("utf-8")
    ).hexdigest()


# =========================================================
# COLUMN LOOKUP HELPERS
# =========================================================

def find_col(columns, aliases):

    normalized = {
        norm_header(column): column
        for column in columns
    }

    for alias in aliases:

        key = norm_header(alias)

        if key in normalized:
            return normalized[key]

    return None


def val(row, column, default=""):

    if not column:
        return default

    try:
        value = row[column]
    except (KeyError, IndexError, TypeError):
        return default

    if value is None:
        return default

    return value


# =========================================================
# DATABASE INITIALIZATION
# =========================================================

def init_db():

    with db() as con:

        # -------------------------------------------------
        # USERS
        # -------------------------------------------------

        if is_postgres():
            con.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                    username TEXT NOT NULL UNIQUE,
                    password_hash TEXT NOT NULL,
                    role TEXT NOT NULL DEFAULT 'biller',
                    display_name TEXT,
                    active INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT
                )
            """)
        else:
            con.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT NOT NULL UNIQUE,
                    password_hash TEXT NOT NULL,
                    role TEXT NOT NULL DEFAULT 'biller',
                    display_name TEXT,
                    active INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT
                )
            """)

        # -------------------------------------------------
        # PROJECTS
        # -------------------------------------------------

        if is_postgres():
            con.execute("""
                CREATE TABLE IF NOT EXISTS projects (
                    id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                    project_name TEXT NOT NULL UNIQUE,
                    name TEXT,
                    role TEXT,
                    active INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT
                )
            """)
        else:
            con.execute("""
                CREATE TABLE IF NOT EXISTS projects (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_name TEXT NOT NULL UNIQUE,
                    name TEXT,
                    role TEXT,
                    active INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT
                )
            """)

        # -------------------------------------------------
        # USER / PROJECT ASSOCIATION
        # One user can have multiple projects.
        # -------------------------------------------------

        if is_postgres():

            con.execute("""
                CREATE TABLE IF NOT EXISTS user_projects (
                    user_id INTEGER NOT NULL
                        REFERENCES users(id) ON DELETE CASCADE,
                    project_id INTEGER NOT NULL
                        REFERENCES projects(id) ON DELETE CASCADE,
                    PRIMARY KEY (user_id, project_id)
                )
            """)

        else:

            con.execute("""
                CREATE TABLE IF NOT EXISTS user_projects (
                    user_id INTEGER NOT NULL
                        REFERENCES users(id) ON DELETE CASCADE,
                    project_id INTEGER NOT NULL
                        REFERENCES projects(id) ON DELETE CASCADE,
                    PRIMARY KEY (user_id, project_id)
                )
            """)

        # -------------------------------------------------
        # ERROR RECORDS
        # -------------------------------------------------

        if is_postgres():
            con.execute("""
                CREATE TABLE IF NOT EXISTS error_records (
                    id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                    record_key TEXT NOT NULL UNIQUE,
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
                    verified_error TEXT DEFAULT 'N',
                    biller_comment TEXT,
                    qc_explanation TEXT,
                    qc_name TEXT,
                    updated_at TEXT,
                    project_id INTEGER
                )
            """)
        else:
            con.execute("""
                CREATE TABLE IF NOT EXISTS error_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    record_key TEXT NOT NULL UNIQUE,
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
                    verified_error TEXT DEFAULT 'N',
                    biller_comment TEXT,
                    qc_explanation TEXT,
                    qc_name TEXT,
                    updated_at TEXT,
                    project_id INTEGER
                )
            """)

        # -------------------------------------------------
        # ACCURACY RECORDS
        # -------------------------------------------------

        if is_postgres():
            con.execute("""
                CREATE TABLE IF NOT EXISTS accuracy_records (
                    id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                    sheet TEXT,
                    username TEXT,
                    biller_name TEXT,
                    month TEXT,
                    weekly TEXT,
                    record_date TEXT,
                    accuracy_type TEXT,
                    accuracy_value DOUBLE PRECISION,
                    bills_processed TEXT,
                    audited TEXT,
                    errors TEXT,
                    source_row INTEGER,
                    imported_at TEXT,
                    project_id INTEGER
                )
            """)
        else:
            con.execute("""
                CREATE TABLE IF NOT EXISTS accuracy_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    sheet TEXT,
                    username TEXT,
                    biller_name TEXT,
                    month TEXT,
                    weekly TEXT,
                    record_date TEXT,
                    accuracy_type TEXT,
                    accuracy_value REAL,
                    bills_processed TEXT,
                    audited TEXT,
                    errors TEXT,
                    source_row INTEGER,
                    imported_at TEXT,
                    project_id INTEGER
                )
            """)

        # -------------------------------------------------
        # BILLER RESPONSES
        # -------------------------------------------------

        if is_postgres():
            con.execute("""
                CREATE TABLE IF NOT EXISTS biller_responses (
                    id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                    record_key TEXT,
                    verified_error TEXT,
                    biller_comment TEXT,
                    qc_explanation TEXT,
                    qc_name TEXT,
                    updated_by TEXT,
                    updated_at TEXT,
                    UNIQUE(record_key)
                )
            """)
        else:
            con.execute("""
                CREATE TABLE IF NOT EXISTS biller_responses (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    record_key TEXT,
                    verified_error TEXT,
                    biller_comment TEXT,
                    qc_explanation TEXT,
                    qc_name TEXT,
                    updated_by TEXT,
                    updated_at TEXT,
                    UNIQUE(record_key)
                )
            """)
            
        # -------------------------------------------------
        # IMPORT HISTORY
        # -------------------------------------------------

        if is_postgres():
            con.execute("""
                CREATE TABLE IF NOT EXISTS import_history (
                    id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                    filename TEXT,
                    uploaded_by TEXT,
                    imported_by TEXT,
                    records_imported INTEGER DEFAULT 0,
                    imported_at TEXT
                )
            """)
        else:
            con.execute("""
                CREATE TABLE IF NOT EXISTS import_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    filename TEXT,
                    uploaded_by TEXT,
                    imported_by TEXT,
                    records_imported INTEGER DEFAULT 0,
                    imported_at TEXT
                )
            """)

        # -------------------------------------------------
        # ACTIVITY LOGS
        # -------------------------------------------------

        if is_postgres():
            con.execute("""
                CREATE TABLE IF NOT EXISTS activity_logs (
                    id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                    username TEXT,
                    action TEXT,
                    details TEXT,
                    created_at TEXT
                )
            """)
        else:
            con.execute("""
                CREATE TABLE IF NOT EXISTS activity_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT,
                    action TEXT,
                    details TEXT,
                    created_at TEXT
                )
            """)

        
        # -------------------------------------------------
        # SAFE ADDITION OF MISSING COLUMNS
        # -------------------------------------------------

        required_columns = {

            "users": {
                "display_name": "TEXT",
                "active": "INTEGER DEFAULT 1",
                "created_at": "TEXT"
            },

            "projects": {
                "name": "TEXT",
                "role": "TEXT",
                "active": "INTEGER DEFAULT 1",
                "created_at": "TEXT"
            },

            "error_records": {
                "billed_date": "TEXT",
                "week": "TEXT",
                "source": "TEXT",
                "audit_type": "TEXT",
                "biller_name": "TEXT",
                "username": "TEXT",
                "id_number": "TEXT",
                "auditor": "TEXT",
                "pro": "TEXT",
                "zb_field": "TEXT",
                "error_details": "TEXT",
                "source_row": "INTEGER",
                "imported_at": "TEXT",
                "verified_error": "TEXT DEFAULT 'N'",
                "biller_comment": "TEXT",
                "qc_explanation": "TEXT",
                "qc_name": "TEXT",
                "updated_at": "TEXT",
                "updated_by": "TEXT"
            },

            "accuracy_records": {
                "sheet": "TEXT",
                "username": "TEXT",
                "biller_name": "TEXT",
                "month": "TEXT",
                "weekly": "TEXT",
                "record_date": "TEXT",
                "accuracy_type": "TEXT",
                "accuracy_value": "REAL",
                "bills_processed": "TEXT",
                "audited": "TEXT",
                "errors": "TEXT",
                "source_row": "INTEGER",
                "imported_at": "TEXT"
            },

            "biller_responses": {
                "record_key": "TEXT",
                "verified_error": "TEXT",
                "biller_comment": "TEXT",
                "qc_explanation": "TEXT",
                "qc_name": "TEXT",
                "updated_by": "TEXT",
                "updated_at": "TEXT"
            },

            "import_history": {
                "filename": "TEXT",
                "uploaded_by": "TEXT",
                "imported_by": "TEXT",
                "records_imported": "INTEGER DEFAULT 0",
                "imported_at": "TEXT"
            },

            "activity_logs": {
                "username": "TEXT",
                "action": "TEXT",
                "details": "TEXT",
                "created_at": "TEXT"
            }
        }

        for table_name, columns in required_columns.items():

            for column_name, definition in columns.items():

                add_column_if_missing(
                    con,
                    table_name,
                    column_name,
                    definition
                )

        # -------------------------------------------------
        # PROJECT ID COLUMNS
        # Existing records remain NULL if they have no project.
        # -------------------------------------------------

        add_column_if_missing(
            con,
            "error_records",
            "project_id",
            "INTEGER"
        )

        add_column_if_missing(
            con,
            "accuracy_records",
            "project_id",
            "INTEGER"
        )
        
        # -------------------------------------------------
        # ONE-TIME SYNC OF EXISTING BILLER RESPONSES
        # Preserve historical Biller and QC responses.
        # -------------------------------------------------

        con.execute("""
            CREATE TABLE IF NOT EXISTS schema_migrations (
                migration_name TEXT PRIMARY KEY,
                applied_at TEXT
            )
        """)

        
        # -------------------------------------------------
        # SAFE LEGACY PROJECT ASSIGNMENT BACKFILL
        # Copy old single-project assignments into the
        # many-to-many table without deleting/replacing rows.
        # -------------------------------------------------
        assignment_migration = "copy_legacy_assigned_project_id_v1"
        assignment_done = con.execute(
            "SELECT migration_name FROM schema_migrations WHERE migration_name = ?",
            (assignment_migration,)
        ).fetchone()
        if not assignment_done and "assigned_project_id" in table_columns(con, "users"):
            con.execute("""
                INSERT INTO user_projects (user_id, project_id)
                SELECT u.id, u.assigned_project_id
                FROM users u
                INNER JOIN projects p ON p.id = u.assigned_project_id
                WHERE u.assigned_project_id IS NOT NULL
                ON CONFLICT (user_id, project_id) DO NOTHING
            """)
        if not assignment_done:
            con.execute(
                "INSERT INTO schema_migrations (migration_name, applied_at) VALUES (?, ?)",
                (assignment_migration, datetime.now().isoformat(timespec="seconds"))
            )

        migration_name = "copy_biller_responses_to_error_records_v1"

        migration = con.execute(
            """
            SELECT migration_name
            FROM schema_migrations
            WHERE migration_name = ?
            """,
            (migration_name,)
        ).fetchone()

        if not migration:

            con.execute("""
                UPDATE error_records
                SET
                    verified_error = CASE
                        WHEN EXISTS (
                            SELECT 1 FROM biller_responses br
                            WHERE br.record_key = error_records.record_key
                              AND NULLIF(TRIM(br.verified_error), '') IS NOT NULL
                        ) AND (
                            error_records.verified_error IS NULL
                            OR TRIM(error_records.verified_error) = ''
                            OR (UPPER(TRIM(error_records.verified_error)) = 'N'
                                AND EXISTS (
                                    SELECT 1 FROM biller_responses br
                                    WHERE br.record_key = error_records.record_key
                                      AND UPPER(TRIM(COALESCE(br.verified_error, ''))) = 'Y'
                                ))
                        ) THEN (
                            SELECT br.verified_error FROM biller_responses br
                            WHERE br.record_key = error_records.record_key
                        )
                        ELSE error_records.verified_error
                    END,
                    biller_comment = COALESCE(NULLIF(error_records.biller_comment, ''),
                        (SELECT br.biller_comment FROM biller_responses br WHERE br.record_key = error_records.record_key)),
                    qc_explanation = COALESCE(NULLIF(error_records.qc_explanation, ''),
                        (SELECT br.qc_explanation FROM biller_responses br WHERE br.record_key = error_records.record_key)),
                    qc_name = COALESCE(NULLIF(error_records.qc_name, ''),
                        (SELECT br.qc_name FROM biller_responses br WHERE br.record_key = error_records.record_key)),
                    updated_at = COALESCE(NULLIF(error_records.updated_at, ''),
                        (SELECT br.updated_at FROM biller_responses br WHERE br.record_key = error_records.record_key)),
                    updated_by = COALESCE(NULLIF(error_records.updated_by, ''),
                        (SELECT br.updated_by FROM biller_responses br WHERE br.record_key = error_records.record_key))
                WHERE EXISTS (
                    SELECT 1 FROM biller_responses br
                    WHERE br.record_key = error_records.record_key
                )
            """)

            con.execute(
                """
                INSERT INTO schema_migrations (
                    migration_name,
                    applied_at
                )
                VALUES (?, ?)
                """,
                (
                    migration_name,
                    datetime.now().isoformat(
                        timespec="seconds"
                    )
                )
            )


# =========================================================
# PROJECT ASSIGNMENT HELPERS
# =========================================================

def get_user_project_ids(user_id=None):

    if user_id is None:
        user_id = session.get("user_id")

    if not user_id:
        return []

    with db() as con:

        user = con.execute(
            """
            SELECT id, role, active
            FROM users
            WHERE id = ?
            """,
            (user_id,)
        ).fetchone()

        if not user or not user["active"]:
            return []

        role = normalize_text(user["role"]).lower()

        if role == "admin":

            rows = con.execute(
                """
                SELECT id
                FROM projects
                WHERE active = 1
                ORDER BY project_name
                """
            ).fetchall()
        else:

            rows = con.execute(
                """
                SELECT p.id
                FROM projects p
                INNER JOIN user_projects up
                    ON up.project_id = p.id
                WHERE up.user_id = ?
                AND p.active = 1
                ORDER BY p.project_name
                """,
                (user_id,)
            ).fetchall()

        return sorted(
            int(row["id"])
            for row in rows
        )


def allowed_project_ids():
    return get_user_project_ids()


def user_has_project(user_id, project_id):

    try:
        user_id = int(user_id)
        project_id = int(project_id)
    except (TypeError, ValueError):
        return False

    with db() as con:

        user = con.execute(
            """
            SELECT role, active
            FROM users
            WHERE id = ?
            """,
            (user_id,)
        ).fetchone()

        if not user or not user["active"]:
            return False

        if normalize_text(user["role"]).lower() == "admin":

            project = con.execute(
                """
                SELECT id
                FROM projects
                WHERE id = ?
                AND active = 1
                """,
                (project_id,)
            ).fetchone()

            return project is not None

        project = con.execute(
            """
            SELECT id
            FROM projects
            WHERE id = ?
            AND active = 1
            """,
            (project_id,)
        ).fetchone()

        if not project:
            return False

        assignment = con.execute(
            """
            SELECT 1
            FROM user_projects
            WHERE user_id = ?
            AND project_id = ?
            """,
            (user_id, project_id)
        ).fetchone()

        return assignment is not None


def assign_user_projects(user_id, project_ids):

    try:
        user_id = int(user_id)
    except (TypeError, ValueError):
        raise ValueError("Invalid user ID.")

    clean_ids = set()

    for project_id in project_ids or []:

        try:
            clean_ids.add(int(project_id))
        except (TypeError, ValueError):
            continue

    with db() as con:

        user = con.execute(
            """
            SELECT id, role
            FROM users
            WHERE id = ?
            """,
            (user_id,)
        ).fetchone()

        if not user:
            raise ValueError("User account not found.")

        role = normalize_text(user["role"]).lower()

        if role == "admin":
            return

        if not clean_ids:
            raise ValueError(
                "Select at least one project."
            )

        placeholders = ",".join(
            ["?"] * len(clean_ids)
        )

        valid_projects = con.execute(
            f"""
            SELECT id
            FROM projects
            WHERE active = 1
            AND id IN ({placeholders})
            """,
            tuple(sorted(clean_ids))
        ).fetchall()

        valid_ids = {
            int(row["id"])
            for row in valid_projects
        }

        if valid_ids != clean_ids:
            raise ValueError(
                "One or more selected projects are invalid or inactive."
            )

        con.execute(
            """
            DELETE FROM user_projects
            WHERE user_id = ?
            """,
            (user_id,)
        )

        con.executemany(
            """
            INSERT INTO user_projects (
                user_id,
                project_id
            )
            VALUES (?, ?)
            """,
            [
                (user_id, project_id)
                for project_id in sorted(clean_ids)
            ]
        )



# =========================================================
# AUTHENTICATION AND ROLE DECORATORS
# =========================================================


def get_active_session_user():

    user_id = session.get("user_id")

    if not user_id:
        return None

    with db() as con:

        user = con.execute(
            """
            SELECT
                id,
                username,
                role,
                display_name,
                active
            FROM users
            WHERE id = ?
            """,
            (user_id,)
        ).fetchone()

    if not user or not user["active"]:
        session.clear()
        return None

    current_role = normalize_text(
        user["role"]
    ).lower()

    if current_role not in ("admin", "qc", "biller"):
        session.clear()
        return None

    session["username"] = user["username"]
    session["role"] = current_role

    session["display_name"] = (
        user["display_name"]
        or user["username"]
    )

    return user

def login_required(function):

    @wraps(function)
    def wrapper(*args, **kwargs):

        user = get_active_session_user()

        if not user:
            flash(
                "Your session has expired or your account is inactive. Please log in again.",
                "warning"
            )

            return redirect(url_for("login"))

        return function(*args, **kwargs)

    return wrapper


def role_required(*allowed_roles):

    normalized_roles = {
        normalize_text(role).lower()
        for role in allowed_roles
    }

    def decorator(function):

        @wraps(function)
        def wrapper(*args, **kwargs):

            user = get_active_session_user()

            if not user:
                flash(
                    "Your session has expired or your account is inactive. Please log in again.",
                    "warning"
                )

                return redirect(url_for("login"))

            current_role = normalize_text(
                user["role"]
            ).lower()

            if current_role not in normalized_roles:
                abort(403)

            return function(*args, **kwargs)

        return wrapper

    return decorator


def admin_required(function):

    return role_required("admin")(function)


def biller_required(function):

    return role_required("biller")(function)


def qc_required(function):

    return role_required("qc")(function)


# =========================================================
# ACTIVITY LOGGING
# =========================================================

def log_action(username, action, details=""):

    with db() as con:

        con.execute(
            """
            INSERT INTO activity_logs (
                username,
                action,
                details,
                created_at
            )
            VALUES (?, ?, ?, ?)
            """,
            (
                normalize_text(username),
                normalize_text(action),
                normalize_text(details),
                datetime.now().isoformat(timespec="seconds")
            )
        )


# =========================================================
# CSRF PROTECTION
# =========================================================

@app.before_request
def csrf_token():

    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_hex(32)


def validate_csrf():

    submitted = request.form.get("csrf_token", "")
    stored = session.get("csrf_token", "")

    if not submitted:
        app.logger.error("CSRF ERROR: Submitted token is missing")
        abort(400)

    if not stored:
        app.logger.error("CSRF ERROR: Session token is missing")
        abort(400)

    if not secrets.compare_digest(submitted, stored):
        app.logger.error("CSRF ERROR: Tokens do not match")
        abort(400)


@app.context_processor
def inject_globals():

    return {
        "csrf_token": session.get("csrf_token", "")
    }
    
# =========================================================
# WORKBOOK READING
# =========================================================

class ODSWorkbookSource:
    """A seekable ODS upload parsed one sheet and one bounded chunk at a time."""

    CHUNK_SIZE = 500
    SUPPORTED_SHEETS = {"ERROR", "INTERNAL", "EXTERNAL"}

    def __init__(self, stream):
        self.stream = stream

    @staticmethod
    def _cell_text(cell):
        office = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
        text_ns = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
        value = (
            cell.get(office + "string-value")
            or cell.get(office + "date-value")
            or cell.get(office + "time-value")
            or cell.get(office + "boolean-value")
            or cell.get(office + "value")
        )
        paragraphs = []
        for paragraph in cell.iter(text_ns + "p"):
            paragraphs.append("".join(paragraph.itertext()))
        paragraph_text = "\n".join(part for part in paragraphs if part != "")
        if paragraph_text != "":
            return paragraph_text
        return value if value is not None else ""

    @classmethod
    def _row_values(cls, row_element):
        table_ns = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
        cell_tags = {
            table_ns + "table-cell",
            table_ns + "covered-table-cell",
        }
        values = []
        for cell in list(row_element):
            if cell.tag not in cell_tags:
                continue
            repeat_text = cell.get(table_ns + "number-columns-repeated", "1")
            try:
                repeat = max(1, min(int(repeat_text), 16384))
            except (TypeError, ValueError):
                repeat = 1
            value = cls._cell_text(cell)
            span_text = cell.get(table_ns + "number-columns-spanned", "1")
            try:
                span = max(1, min(int(span_text), 16384))
            except (TypeError, ValueError):
                span = 1
            values.extend([value] * repeat)
            if span > 1:
                values.extend([""] * (span - 1))
        # Trailing empty cells are not useful to the import and can be
        # repeated by spreadsheet formatting across thousands of columns.
        while values and values[-1] == "":
            values.pop()
        return values

    def iter_chunks(self):
        """Yield (sheet_name, DataFrame-or-None); None marks an empty supported sheet."""
        table_ns = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
        self.stream.seek(0)
        try:
            archive = zipfile.ZipFile(self.stream)
        except (zipfile.BadZipFile, OSError) as exc:
            raise ValueError("The ODS workbook is invalid or corrupted.") from exc

        with archive:
            if "content.xml" not in archive.namelist():
                raise ValueError("The ODS workbook is missing content.xml.")
            with archive.open("content.xml") as xml_stream:
                stack = []
                current_sheet = None
                headers = None
                physical_row = 0
                rows = []
                row_indexes = []
                yielded_for_sheet = False

                try:
                    iterator = ET.iterparse(xml_stream, events=("start", "end"))
                    for event, element in iterator:
                        if event == "start":
                            stack.append(element)
                            if element.tag == table_ns + "table":
                                raw_name = element.get(table_ns + "name", "")
                                normalized = normalize_text(raw_name).upper()
                                current_sheet = normalized if normalized in self.SUPPORTED_SHEETS else None
                                headers = None
                                physical_row = 0
                                rows = []
                                row_indexes = []
                                yielded_for_sheet = False
                            continue

                        if element.tag == table_ns + "table-row" and current_sheet:
                            row_values = self._row_values(element)
                            repeat_text = element.get(table_ns + "number-rows-repeated", "1")
                            try:
                                row_repeat = max(1, min(int(repeat_text), 1000000))
                            except (TypeError, ValueError):
                                row_repeat = 1

                            if headers is None:
                                headers = row_values
                                physical_row = row_repeat
                            else:
                                has_data = any(value not in (None, "") for value in row_values)
                                if not has_data:
                                    physical_row += row_repeat
                                else:
                                    for _ in range(row_repeat):
                                        physical_row += 1
                                        padded = row_values[:len(headers)]
                                        if len(padded) < len(headers):
                                            padded.extend([""] * (len(headers) - len(padded)))
                                        rows.append(padded)
                                        # Pandas importers use source row = index + 2.
                                        row_indexes.append(physical_row - 2)
                                        if len(rows) >= self.CHUNK_SIZE:
                                            frame = pd.DataFrame(rows, columns=headers, index=row_indexes, dtype=object)
                                            yield current_sheet, frame
                                            yielded_for_sheet = True
                                            rows = []
                                            row_indexes = []

                            if len(stack) >= 2:
                                try:
                                    stack[-2].remove(element)
                                except ValueError:
                                    pass
                            element.clear()

                        elif element.tag == table_ns + "table" and current_sheet:
                            if rows:
                                frame = pd.DataFrame(rows, columns=headers or [], index=row_indexes, dtype=object)
                                yield current_sheet, frame
                                yielded_for_sheet = True
                            if not yielded_for_sheet:
                                yield current_sheet, None
                            current_sheet = None
                            headers = None
                            rows = []
                            row_indexes = []
                            if len(stack) >= 2:
                                try:
                                    stack[-2].remove(element)
                                except ValueError:
                                    pass
                            element.clear()

                        if stack:
                            stack.pop()
                except ET.ParseError as exc:
                    raise ValueError("The ODS workbook contains invalid XML and could not be read.") from exc


def read_source(file_storage, csv_sheet="ERROR"):

    filename = secure_filename(
        file_storage.filename or "source"
    )

    extension = os.path.splitext(filename)[1].lower()

    if extension not in ALLOWED_EXTENSIONS:
        raise ValueError("Unsupported file type.")

    # ODS files are parsed directly from Flask's spooled upload stream.
    # Do not duplicate the complete compressed file in a second BytesIO buffer.
    if extension == ".ods":
        stream = file_storage.stream
        stream.seek(0)
        if not zipfile.is_zipfile(stream):
            raise ValueError("The uploaded ODS file is invalid or corrupted.")
        stream.seek(0)
        return ODSWorkbookSource(stream)

    file_bytes = file_storage.read()
    if not file_bytes:
        raise ValueError("The uploaded file is empty.")
    buffer = io.BytesIO(file_bytes)

    if extension == ".csv":
        last_error = None
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                buffer.seek(0)
                dataframe = pd.read_csv(buffer, encoding=encoding, dtype=object)
                sheet_name = normalize_text(csv_sheet).upper()
                if sheet_name not in {"ERROR", "INTERNAL", "EXTERNAL"}:
                    raise ValueError("Invalid CSV sheet selection.")
                return {sheet_name: dataframe}
            except UnicodeDecodeError as exc:
                last_error = exc
        raise ValueError(f"Unable to read CSV encoding: {last_error}")

    workbook = pd.ExcelFile(buffer)
    result = {}
    available_sheets = {
        normalize_text(name).upper(): name
        for name in workbook.sheet_names
    }
    for requested in ("ERROR", "INTERNAL", "EXTERNAL"):
        actual_name = available_sheets.get(requested)
        if actual_name is not None:
            result[requested] = pd.read_excel(
                workbook,
                sheet_name=actual_name,
                dtype=object,
                na_filter=False
            )
    if not result:
        raise ValueError(
            "No supported sheets found. Expected ERROR, INTERNAL, or EXTERNAL."
        )
    return result


# =========================================================
# PROJECT LOOKUP
# =========================================================

def get_import_project_id(con, project_name):

    project_name = normalize_text(project_name)

    if not project_name:
        return None

    row = con.execute(
        """
        SELECT id
        FROM projects
        WHERE LOWER(TRIM(project_name)) = LOWER(TRIM(?))
        AND active = 1
        """,
        (project_name,)
    ).fetchone()

    if not row:
        raise ValueError(
            f"Project '{project_name}' does not exist "
            "or is inactive. Please add or activate it "
            "under Admin > Project first."
        )

    return int(row["id"])


def resolve_workbook_project_id(
    con,
    project_name,
    default_project_id=None,
    allowed_project_ids=None
):
    """Resolve a row's project and enforce the uploader's project scope."""
    if normalize_text(project_name):
        project_id = get_import_project_id(con, project_name)
    elif default_project_id is not None:
        project_id = int(default_project_id)
        active = con.execute(
            "SELECT id FROM projects WHERE id = ? AND active = 1",
            (project_id,)
        ).fetchone()
        if not active:
            raise ValueError("The selected project does not exist or is inactive.")
    elif allowed_project_ids is not None and len(allowed_project_ids) == 1:
        project_id = int(allowed_project_ids[0])
    else:
        raise ValueError(
            "A PROJECT column is required when importing into multiple projects. "
            "Or select one project before uploading."
        )

    if allowed_project_ids is not None and int(project_id) not in {
        int(value) for value in allowed_project_ids
    }:
        raise ValueError(
            f"Project '{normalize_text(project_name) or project_id}' is not assigned to your QC account."
        )
    return int(project_id)


# =========================================================
# DATAFRAME HELPERS
# =========================================================

def dataframe_value(row, column, default=""):

    if not column:
        return default

    try:
        value = row[column]
    except (KeyError, TypeError):
        return default

    return normalize_text(value)


def get_error_columns(dataframe):

    columns = list(dataframe.columns)

    return {
        key: find_col(columns, aliases)
        for key, aliases in ERROR_ALIASES.items()
    }


# =========================================================
# ACCURACY COLUMN MAPPING
# =========================================================

def get_accuracy_columns(dataframe):

    columns = list(dataframe.columns)

    return {
        "username": find_col(
            columns,
            ["USER", "USERNAME", "USER NAME"]
        ),

        "biller_name": find_col(
            columns,
            ["NAME", "BILLER NAME", "BILLER"]
        ),

        "date": find_col(
            columns,
            ["DATE"]
        ),
        
        "month": find_col(
            columns,
            ["MONTH"]
        ),

        "weekly": find_col(
            columns,
            ["WEEKLY", "WEEK"]
        ),

        "audited": find_col(
            columns,
            ["AUDITED"]
        ),

        "bills_processed": find_col(
            columns,
            ["BILLS PROCESSED"]
        ),

        "internal_errors": find_col(
            columns,
            ["ERROR"]
        ),

        "external_woc_errors": find_col(
            columns,
            ["ERROR - WOC"]
        ),

        "internal_accuracy": find_col(
            columns,
            ["ACCURACY"]
        ),

        "external_wc_accuracy": find_col(
            columns,
            ["PER AUDITED ACCURACY"]
        ),

        "external_woc_accuracy": find_col(
            columns,
            ["PER AUDITED ACCURACY - WOC"]
        ),

        "project": find_col(
            columns,
            ["PROJECT", "PROJECT NAME"]
        )
    }



# =========================================================
# ERROR RECORD IMPORT - BATCH UPSERT
# =========================================================

def import_error_sheet(
    con,
    dataframe,
    default_project_id=None,
    allowed_project_ids=None,
    occurrences=None,
    imported_at=None
):
    if dataframe is None or dataframe.empty:
        return 0

    if imported_at is None:
        imported_at = datetime.now().isoformat(timespec="microseconds")

    columns = get_error_columns(dataframe)
    if not columns["username"] or not columns["pro"]:
        raise ValueError("ERROR sheet must contain USERNAME and PRO columns.")

    imported = 0
    if occurrences is None:
        occurrences = {}
    project_cache = {}
    batch = []
    upsert_sql = """
        INSERT INTO error_records (
            record_key, billed_date, week, source, audit_type, biller_name,
            username, id_number, auditor, pro, zb_field, error_details,
            project_id, verified_error, imported_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'N', ?)
        ON CONFLICT (record_key) DO UPDATE SET
            billed_date = EXCLUDED.billed_date,
            week = EXCLUDED.week,
            source = EXCLUDED.source,
            audit_type = EXCLUDED.audit_type,
            biller_name = EXCLUDED.biller_name,
            username = EXCLUDED.username,
            id_number = EXCLUDED.id_number,
            auditor = EXCLUDED.auditor,
            pro = EXCLUDED.pro,
            zb_field = EXCLUDED.zb_field,
            error_details = EXCLUDED.error_details,
            project_id = COALESCE(error_records.project_id, EXCLUDED.project_id),
            imported_at = EXCLUDED.imported_at
    """

    for _, row in dataframe.iterrows():
        username = dataframe_value(row, columns["username"])
        pro = dataframe_value(row, columns["pro"])
        if not username or not pro:
            continue

        project_name = dataframe_value(row, columns["project"])
        cache_key = normalize_text(project_name).casefold() or f"default:{default_project_id}"
        if cache_key not in project_cache:
            project_cache[cache_key] = resolve_workbook_project_id(
                con, project_name, default_project_id, allowed_project_ids
            )
        project_id = project_cache[cache_key]

        billed_date = parse_date(val(row, columns["billed_date"]))
        billed_date_text = billed_date.isoformat() if billed_date else ""
        source = dataframe_value(row, columns["source"])
        audit_type = dataframe_value(row, columns["audit_type"])
        biller_name = dataframe_value(row, columns["biller_name"])
        id_number = dataframe_value(row, columns["id_number"])
        auditor = dataframe_value(row, columns["auditor"])
        zb_field = dataframe_value(row, columns["zb_field"])
        error_details = dataframe_value(row, columns["error_details"])
        week = week_for_billed_date(billed_date_text)

        fingerprint_parts = [
            billed_date_text, source, audit_type, biller_name, username,
            id_number, auditor, pro, zb_field, error_details
        ]
        base_fingerprint = fingerprint(fingerprint_parts, occurrence=0)
        occurrence = occurrences.get(base_fingerprint, 0)
        occurrences[base_fingerprint] = occurrence + 1
        record_key = fingerprint(fingerprint_parts, occurrence=occurrence)

        batch.append((
            record_key, billed_date_text, week, source, audit_type,
            biller_name, username, id_number, auditor, pro, zb_field,
            error_details, project_id,
            imported_at
        ))
        imported += 1
        if len(batch) >= 500:
            con.executemany(upsert_sql, batch)
            batch.clear()

    if batch:
        con.executemany(upsert_sql, batch)
    return imported


# =========================================================
# ACCURACY RECORD IMPORT - BATCH UPSERT
# INTERNAL:
#   ACCURACY
#
# EXTERNAL:
#   PER AUDITED ACCURACY       = WC
#   PER AUDITED ACCURACY - WOC = WOC
# =========================================================

def import_accuracy_sheet(
    con,
    dataframe,
    sheet_name,
    default_project_id=None,
    allowed_project_ids=None,
    imported_at=None
):
    if dataframe is None or dataframe.empty:
        return 0

    sheet_name = normalize_text(sheet_name).upper()
    if sheet_name not in ("INTERNAL", "EXTERNAL"):
        raise ValueError("Invalid accuracy sheet.")

    columns = get_accuracy_columns(dataframe)
    if not columns["username"]:
        raise ValueError(f"{sheet_name} sheet must contain USER.")
    if not columns["biller_name"]:
        raise ValueError(f"{sheet_name} sheet must contain NAME.")
    if not columns["audited"]:
        raise ValueError(f"{sheet_name} sheet must contain AUDITED.")

    if sheet_name == "INTERNAL":
        accuracy_types = [(INTERNAL_ACCURACY, columns["internal_accuracy"], columns["internal_errors"])]
    else:
        accuracy_types = [
            (EXTERNAL_ACCURACY[0], columns["external_wc_accuracy"], columns["internal_errors"]),
            (EXTERNAL_ACCURACY[1], columns["external_woc_accuracy"], columns["external_woc_errors"])
        ]

    # -----------------------------------------------------
    # PRELOAD EXISTING ROWS IN ONE QUERY
    # The old implementation ran a SELECT for every workbook
    # row, which caused long-running imports and Gunicorn
    # worker timeouts on larger ODS files.
    # -----------------------------------------------------
    source_row_numbers = [int(index) + 2 for index in dataframe.index]
    first_source_row = min(source_row_numbers) if source_row_numbers else 2
    last_source_row = max(source_row_numbers) if source_row_numbers else 1

    existing_by_key = {}
    existing_blank_by_key = {}
    latest_blank_import_by_project = {}

    existing_rows = con.execute(
        """
        SELECT
            id,
            username,
            record_date,
            accuracy_type,
            source_row,
            project_id,
            imported_at
        FROM accuracy_records
        WHERE sheet = ? AND source_row BETWEEN ? AND ?
        ORDER BY id DESC
        """,
        (sheet_name, first_source_row, last_source_row)
    ).fetchall()

    # Identify the latest import batch containing blank dates
    # for each project.
    for existing in existing_rows:
        if normalize_text(existing["record_date"]):
            continue

        project_key = (
            int(existing["project_id"])
            if existing["project_id"] is not None
            else None
        )

        batch_time = normalize_text(existing["imported_at"])

        if not batch_time:
            continue

        current_latest = latest_blank_import_by_project.get(project_key)

        if current_latest is None or batch_time > current_latest:
            latest_blank_import_by_project[project_key] = batch_time

    # Build exact matches and safe fallback matches.
    for existing in existing_rows:
        username_key = normalize_text(existing["username"])
        date_key = normalize_text(existing["record_date"])
        accuracy_type_key = normalize_text(existing["accuracy_type"])
        source_row_key = int(existing["source_row"])

        project_key = (
            int(existing["project_id"])
            if existing["project_id"] is not None
            else None
        )

        existing_id = int(existing["id"])

        exact_key = (
            username_key,
            date_key,
            accuracy_type_key,
            source_row_key,
            project_key
        )

        existing_by_key.setdefault(exact_key, existing_id)

        batch_time = normalize_text(existing["imported_at"])

        # Only use blank-date records from the latest blank-date
        # import batch for that project.
        if (
            not date_key
            and batch_time
            == latest_blank_import_by_project.get(project_key)
        ):
            fallback_key = (
                username_key,
                accuracy_type_key,
                source_row_key,
                project_key
            )

            existing_blank_by_key.setdefault(
                fallback_key,
                existing_id
            )

    imported = 0
    project_cache = {}

    if imported_at is None:
        imported_at = datetime.now().isoformat(timespec="microseconds")

    updates = []
    inserts = []

    for dataframe_index, row in dataframe.iterrows():
        source_row = int(dataframe_index) + 2
        username = dataframe_value(row, columns["username"])
        biller_name = dataframe_value(row, columns["biller_name"])
        if not username:
            continue

        project_name = dataframe_value(row, columns["project"])
        cache_key = normalize_text(project_name).casefold() or f"default:{default_project_id}"
        if cache_key not in project_cache:
            project_cache[cache_key] = resolve_workbook_project_id(
                con, project_name, default_project_id, allowed_project_ids
            )
        project_id = project_cache[cache_key]

        source_month = dataframe_value(row, columns["month"])
        record_date = parse_accuracy_date(
            val(row, columns["date"]), source_month
        )
        date_text = record_date.isoformat() if record_date else ""
        month = record_date.strftime("%Y-%m") if record_date else source_month
        weekly = week_for_billed_date(record_date) if record_date else dataframe_value(row, columns["weekly"])
        audited = number_value(val(row, columns["audited"])) or 0
        bills_processed = number_value(val(row, columns["bills_processed"]))

        for accuracy_type, accuracy_column, error_column in accuracy_types:
            if not accuracy_column:
                continue
            accuracy_value = normalize_accuracy(val(row, accuracy_column))
            if accuracy_value is None or not 0 <= accuracy_value <= 100:
                continue
            errors = number_value(val(row, error_column)) or 0

            
            project_key = (
                int(project_id)
                if project_id is not None
                else None
            )

            key = (
                username,
                date_text,
                accuracy_type,
                source_row,
                project_key
            )

            values = (
                biller_name,
                month,
                weekly,
                date_text,
                accuracy_value,
                bills_processed,
                audited,
                errors,
                imported_at
            )

            existing_id = existing_by_key.get(key)

            # If the exact date match is missing, check the
            # latest blank-date import batch for this project.
            if existing_id is None and date_text:
                fallback_key = (
                    username,
                    accuracy_type,
                    source_row,
                    project_key
                )

                existing_id = existing_blank_by_key.get(
                    fallback_key
                )

            if existing_id is not None:
                updates.append((*values, existing_id))
            else:
                inserts.append((
                    sheet_name,
                    username,
                    biller_name,
                    month,
                    weekly,
                    date_text,
                    accuracy_type,
                    accuracy_value,
                    bills_processed,
                    audited,
                    errors,
                    source_row,
                    imported_at,
                    project_id
                ))

            imported += 1


    # Execute database writes in batches instead of issuing one
    # SELECT + UPDATE/INSERT cycle per row.
    
    update_sql = """
        UPDATE accuracy_records SET
            biller_name = ?,
            month = ?,
            weekly = ?,
            record_date = ?,
            accuracy_value = ?,
            bills_processed = ?,
            audited = ?,
            errors = ?,
            imported_at = ?
        WHERE id = ?
    """
    insert_sql = """
        INSERT INTO accuracy_records (
            sheet, username, biller_name, month, weekly, record_date,
            accuracy_type, accuracy_value, bills_processed, audited,
            errors, source_row, imported_at, project_id
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    batch_size = 500
    for offset in range(0, len(updates), batch_size):
        con.executemany(update_sql, updates[offset:offset + batch_size])
    for offset in range(0, len(inserts), batch_size):
        con.executemany(insert_sql, inserts[offset:offset + batch_size])

    return imported


# =========================================================
# COMPLETE WORKBOOK IMPORT
# =========================================================

def import_workbook(
    con,
    sheets,
    filename,
    uploaded_by,
    default_project_id=None,
    allowed_project_ids=None
):
    total_imported = 0

    sheet_import_counts = {
        "ERROR": 0,
        "INTERNAL": 0,
        "EXTERNAL": 0
    }

    workbook_imported_at = datetime.now().isoformat(timespec="microseconds")

    if isinstance(sheets, ODSWorkbookSource):
        seen_supported_sheets = set()
        error_occurrences = {}

        # ODS parsing and database writes happen sheet-by-sheet in chunks.
        for sheet_name, dataframe in sheets.iter_chunks():
            seen_supported_sheets.add(sheet_name)

            if dataframe is None or dataframe.empty:
                continue

            if sheet_name == "ERROR":
                sheet_imported = import_error_sheet(
                    con,
                    dataframe,
                    default_project_id,
                    allowed_project_ids,
                    occurrences=error_occurrences,
                    imported_at=workbook_imported_at
                )

                sheet_import_counts["ERROR"] += sheet_imported
                total_imported += sheet_imported

            elif sheet_name in ("INTERNAL", "EXTERNAL"):
                sheet_imported = import_accuracy_sheet(
                    con,
                    dataframe,
                    sheet_name,
                    default_project_id,
                    allowed_project_ids,
                    imported_at=workbook_imported_at
                )

                sheet_import_counts[sheet_name] += sheet_imported
                total_imported += sheet_imported

        if not seen_supported_sheets:
            raise ValueError(
                "No supported sheets found. Expected ERROR, INTERNAL, or EXTERNAL."
            )

    else:
        error_dataframe = sheets.get("ERROR")
        internal_dataframe = sheets.get("INTERNAL")
        external_dataframe = sheets.get("EXTERNAL")

        if error_dataframe is not None:
            sheet_import_counts["ERROR"] += import_error_sheet(
                con,
                error_dataframe,
                default_project_id,
                allowed_project_ids,
                imported_at=workbook_imported_at
            )

        if internal_dataframe is not None:
            sheet_import_counts["INTERNAL"] += import_accuracy_sheet(
                con,
                internal_dataframe,
                "INTERNAL",
                default_project_id,
                allowed_project_ids,
                imported_at=workbook_imported_at
            )

        if external_dataframe is not None:
            sheet_import_counts["EXTERNAL"] += import_accuracy_sheet(
                con,
                external_dataframe,
                "EXTERNAL",
                default_project_id,
                allowed_project_ids,
                imported_at=workbook_imported_at
            )

        total_imported = sum(sheet_import_counts.values())
        
    # -----------------------------------------------------
    # IMPORT HISTORY
    # Detect actual database columns to support both
    # older schemas and the expanded Supabase schema.
    # -----------------------------------------------------

    imported_at = datetime.now().isoformat(timespec="seconds")

    if is_postgres():
        history_columns = {
            row["column_name"]
            for row in con.execute(
                """
                SELECT column_name
                FROM information_schema.columns
                WHERE table_schema = current_schema()
                  AND table_name = 'import_history'
                """
            ).fetchall()
        }
    else:
        history_columns = {
            row["name"]
            for row in con.execute(
                "PRAGMA table_info(import_history)"
            ).fetchall()
        }

    history_data = {
        "filename": secure_filename(filename or "source"),
        "uploaded_by": normalize_text(uploaded_by),
        "imported_by": normalize_text(uploaded_by),
        "records_imported": total_imported,
        "imported_at": imported_at,
        "status": "SUCCESS"
    }

    insert_columns = [
        column
        for column in history_data
        if column in history_columns
    ]

    if not insert_columns:
        raise RuntimeError(
            "No compatible columns found in import_history."
        )

    columns_sql = ", ".join(insert_columns)
    placeholders = ", ".join(["?"] * len(insert_columns))

    values = tuple(
        history_data[column]
        for column in insert_columns
    )

    con.execute(
        f"""
        INSERT INTO import_history ({columns_sql})
        VALUES ({placeholders})
        """,
        values
    )

    # -----------------------------------------------------
    # ACTIVITY LOG
    # -----------------------------------------------------

    con.execute(
        """
        INSERT INTO activity_logs
            (username, action, details, created_at)
        VALUES (?, ?, ?, ?)
        """,
        (
            normalize_text(uploaded_by),
            "WORKBOOK IMPORT",
        (
            f"{filename}: {total_imported} rows imported or updated | "
            f"ERROR: {sheet_import_counts['ERROR']} | "
            f"INTERNAL: {sheet_import_counts['INTERNAL']} | "
            f"EXTERNAL: {sheet_import_counts['EXTERNAL']}"
        ),
            imported_at
        )
    )

    return total_imported


# =========================================================
# ACCURACY FORMATTING
# =========================================================

def format_accuracy(value):

    if value is None:
        return "—"

    try:
        return f"{float(value):.2f}%"
    except (TypeError, ValueError):
        return "—"


def weighted_accuracy(rows):

    total_audited = 0.0
    total_errors = 0.0

    for row in rows:

        audited = number_value(row["audited"])
        errors = number_value(row["errors"])

        if audited is not None:
            total_audited += audited

        if errors is not None:
            total_errors += errors

    if total_audited <= 0:
        return None

    result = (
        (total_audited - total_errors)
        / total_audited
    ) * 100

    return max(0.0, min(100.0, result))


def accuracy_totals(rows):

    total_audited = sum(
        number_value(row["audited"]) or 0
        for row in rows
    )

    total_errors = sum(
        number_value(row["errors"]) or 0
        for row in rows
    )

    return {
        "audited": total_audited,
        "errors": total_errors,
        "accuracy": weighted_accuracy(rows)
    }


# =========================================================
# MONTH FILTER HELPERS
# =========================================================

def get_biller_months(username, project_ids=None):

    project_ids = list(project_ids or [])

    if not project_ids:
        return []

    placeholders = ",".join(
        ["?"] * len(project_ids)
    )

    with db() as con:

        rows = con.execute(
            f"""
            SELECT DISTINCT month
            FROM accuracy_records
            WHERE username = ?
            AND project_id IN ({placeholders})
            AND month IS NOT NULL
            AND TRIM(month) <> ''
            ORDER BY month DESC
            """,
            (username, *project_ids)
        ).fetchall()

    return [
        row["month"]
        for row in rows
    ]


def selected_month_value(months, requested=None):

    if requested and requested in months:
        return requested

    return months[0] if months else ""


def month_year_options(project_ids=None):

    project_ids = list(project_ids or [])

    if not project_ids:
        return []

    placeholders = ",".join(
        ["?"] * len(project_ids)
    )

    with db() as con:

        rows = con.execute(
            f"""
            SELECT DISTINCT month
            FROM accuracy_records
            WHERE project_id IN ({placeholders})
            AND month IS NOT NULL
            AND TRIM(month) <> ''
            ORDER BY month DESC
            """,
            tuple(project_ids)
        ).fetchall()

    return [
        row["month"]
        for row in rows
    ]


# =========================================================
# ACCURACY RECORD RETRIEVAL
# =========================================================

def get_accuracy_rows(
    project_ids,
    username=None,
    month=None,
    sheet=None,
    accuracy_type=None
):

    project_ids = list(project_ids or [])

    if not project_ids:
        return []

    placeholders = ",".join(
        ["?"] * len(project_ids)
    )

    conditions = [
        f"ar.project_id IN ({placeholders})"
    ]

    params = list(project_ids)

    if username:
        conditions.append("ar.username = ?")
        params.append(username)

    if month:
        conditions.append("ar.month = ?")
        params.append(month)

    if sheet:
        conditions.append("ar.sheet = ?")
        params.append(sheet)

    if accuracy_type:
        conditions.append("ar.accuracy_type = ?")
        params.append(accuracy_type)

    where_clause = " AND ".join(conditions)

    with db() as con:

        rows = con.execute(
            f"""
            SELECT
                ar.id,
                ar.sheet,
                ar.username,
                ar.biller_name,
                ar.month,
                ar.weekly,
                ar.record_date,
                ar.accuracy_type,
                ar.accuracy_value,
                ar.bills_processed,
                ar.audited,
                ar.errors,
                ar.source_row,
                ar.imported_at,
                ar.project_id,
                p.project_name

            FROM accuracy_records ar

            LEFT JOIN projects p
                ON p.id = ar.project_id

            WHERE {where_clause}

            ORDER BY
                ar.record_date DESC,
                ar.username,
                ar.accuracy_type
            """,
            tuple(params)
        ).fetchall()

    return [dict(row) for row in rows]


# =========================================================
# GROUP ACCURACY ROWS
# =========================================================

def group_accuracy_rows(
    rows,
    group_by="username"
):

    allowed_group_fields = {
        "username",
        "project_id",
        "accuracy_type",
        "month",
        "weekly",
        "record_date"
    }

    if group_by not in allowed_group_fields:
        raise ValueError("Invalid grouping field.")

    grouped = {}

    for row in rows:

        key = row.get(group_by)

        grouped.setdefault(key, []).append(row)

    return grouped


# =========================================================
# PERIOD DATE HELPERS
# =========================================================

def get_period_start(value, period="MONTHLY"):

    parsed = parse_date(value)

    if not parsed:
        return None

    period = normalize_text(period).upper()

    if period == "DAILY":
        return parsed

    if period == "WEEKLY":

        return parsed - timedelta(
            days=parsed.weekday()
        )

    if period == "MONTHLY":

        return parsed.replace(day=1)

    raise ValueError(
        "Period must be DAILY, WEEKLY, or MONTHLY."
    )


def get_period_end(value, period="MONTHLY"):

    start = get_period_start(value, period)

    if not start:
        return None

    period = normalize_text(period).upper()

    if period == "DAILY":
        return start

    if period == "WEEKLY":
        return start + timedelta(days=6)

    if period == "MONTHLY":

        if start.month == 12:
            next_month = date(
                start.year + 1,
                1,
                1
            )
        else:
            next_month = date(
                start.year,
                start.month + 1,
                1
            )

        return next_month - timedelta(days=1)

    raise ValueError(
        "Period must be DAILY, WEEKLY, or MONTHLY."
    )


def previous_period_start(start, period="MONTHLY"):

    if not start:
        return None

    period = normalize_text(period).upper()

    if period == "DAILY":
        return start - timedelta(days=1)

    if period == "WEEKLY":
        return start - timedelta(days=7)

    if period == "MONTHLY":

        if start.month == 1:
            return date(
                start.year - 1,
                12,
                1
            )

        return date(
            start.year,
            start.month - 1,
            1
        )

    raise ValueError(
        "Period must be DAILY, WEEKLY, or MONTHLY."
    )


# =========================================================
# GET ACCURACY ROWS FOR A DATE PERIOD
# =========================================================

def get_period_rows(
    project_ids,
    start_date,
    end_date,
    username=None,
    accuracy_type=None
):

    project_ids = list(project_ids or [])

    if not project_ids:
        return []

    start_date = parse_date(start_date)
    end_date = parse_date(end_date)

    if not start_date or not end_date:
        return []

    placeholders = ",".join(
        ["?"] * len(project_ids)
    )

    conditions = [
        f"ar.project_id IN ({placeholders})",
        "ar.record_date >= ?",
        "ar.record_date <= ?"
    ]

    params = [
        *project_ids,
        start_date.isoformat(),
        end_date.isoformat()
    ]

    if username:
        conditions.append("ar.username = ?")
        params.append(username)

    if accuracy_type:
        conditions.append("ar.accuracy_type = ?")
        params.append(accuracy_type)

    where_clause = " AND ".join(conditions)

    with db() as con:

        rows = con.execute(
            f"""
            SELECT
                ar.*,
                p.project_name

            FROM accuracy_records ar

            LEFT JOIN projects p
                ON p.id = ar.project_id

            WHERE {where_clause}

            ORDER BY ar.record_date
            """,
            tuple(params)
        ).fetchall()

    return [dict(row) for row in rows]


# =========================================================
# IMPROVEMENT CALCULATION
# =========================================================

def calculate_improvement(
    current_accuracy,
    previous_accuracy
):

    if current_accuracy is None:
        return {
            "change": None,
            "status": "No Data",
            "note": "No current accuracy data available."
        }

    if previous_accuracy is None:
        return {
            "change": None,
            "status": "New Record",
            "note": (
                "New Record. No previous period "
                "is available for comparison."
            )
        }

    change = (
        current_accuracy - previous_accuracy
    )

    if change > 0.000001:

        status = "Improved"

        note = (
            f"Improved by {change:.2f}%"
        )

    elif change < -0.000001:

        status = "Declined"

        note = (
            f"Declined by {abs(change):.2f}%"
        )

    else:

        status = "No Change"

        change = 0.0

        note = "No change from the previous period."

    if current_accuracy < 90:

        note += ", but still below 90%."

    else:

        note += "."

    return {
        "change": change,
        "status": status,
        "note": note
    }


# =========================================================
# ACCURACY MONITORING
# =========================================================

def get_accuracy_monitoring(
    project_ids,
    period="MONTHLY"
):

    project_ids = list(project_ids or [])

    period = normalize_text(period).upper()

    if period not in (
        "DAILY",
        "WEEKLY",
        "MONTHLY"
    ):
        period = "MONTHLY"

    if not project_ids:
        return []

    placeholders = ",".join(
        ["?"] * len(project_ids)
    )

    with db() as con:

        rows = con.execute(
            f"""
            SELECT
                ar.project_id,
                p.project_name,
                ar.username,
                ar.biller_name,
                ar.accuracy_type,
                ar.record_date,
                ar.audited,
                ar.errors,
                ar.accuracy_value

            FROM accuracy_records ar

            LEFT JOIN projects p
                ON p.id = ar.project_id

            WHERE ar.project_id IN ({placeholders})
            AND ar.record_date IS NOT NULL
            AND TRIM(ar.record_date) <> ''

            ORDER BY
                ar.project_id,
                ar.username,
                ar.accuracy_type,
                ar.record_date
            """,
            tuple(project_ids)
        ).fetchall()

    source_rows = [dict(row) for row in rows]

    # -----------------------------------------------------
    # Group by project, biller, and accuracy type
    # -----------------------------------------------------

    grouped = {}

    for row in source_rows:

        key = (
            row["project_id"],
            row["username"],
            row["accuracy_type"]
        )

        grouped.setdefault(key, []).append(row)

    monitoring = []

    for key, biller_rows in grouped.items():

        project_id, username, accuracy_type = key

        valid_dates = [
            parse_date(row["record_date"])
            for row in biller_rows
        ]

        valid_dates = [
            value
            for value in valid_dates
            if value is not None
        ]

        if not valid_dates:
            continue

        latest_date = max(valid_dates)

        current_start = get_period_start(
            latest_date,
            period
        )

        current_end = get_period_end(
            latest_date,
            period
        )

        previous_start = previous_period_start(
            current_start,
            period
        )

        previous_end = current_start - timedelta(days=1)

        # -------------------------------------------------
        # Current period records
        # -------------------------------------------------

        current_rows = [
            row
            for row in biller_rows
            if (
                parse_date(row["record_date"])
                and current_start
                <= parse_date(row["record_date"])
                <= current_end
            )
        ]

        # -------------------------------------------------
        # Previous period records
        # -------------------------------------------------

        previous_rows = [
            row
            for row in biller_rows
            if (
                parse_date(row["record_date"])
                and previous_start
                <= parse_date(row["record_date"])
                <= previous_end
            )
        ]

        current_accuracy = weighted_accuracy(
            current_rows
        )

        previous_accuracy = weighted_accuracy(
            previous_rows
        )

        totals = accuracy_totals(
            current_rows
        )

        improvement = calculate_improvement(
            current_accuracy,
            previous_accuracy
        )

        first = biller_rows[0]

        monitoring.append({

            "project_id": project_id,

            "project_name": (
                first.get("project_name")
                or "Unassigned"
            ),

            "username": username,

            "biller_name": (
                first.get("biller_name")
                or username
            ),

            "accuracy_type": accuracy_type,

            "period": period,

            "period_start": current_start.isoformat(),

            "period_end": current_end.isoformat(),

            "previous_accuracy": previous_accuracy,

            "current_accuracy": current_accuracy,

            "audited": totals["audited"],

            "errors": totals["errors"],

            "change": improvement["change"],

            "status": improvement["status"],

            "note": improvement["note"]

        })

    monitoring.sort(
        key=lambda item: (
            item["project_name"] or "",
            item["biller_name"] or "",
            item["accuracy_type"] or ""
        )
    )

    return monitoring


# =========================================================
# BELOW 90% ACCURACY FILTER
# =========================================================

def get_below_90_accuracy(monitoring):

    result = {
        INTERNAL_ACCURACY: [],
        EXTERNAL_ACCURACY[0]: [],
        EXTERNAL_ACCURACY[1]: []
    }

    for row in monitoring:

        accuracy_type = row["accuracy_type"]

        accuracy = row["current_accuracy"]

        if accuracy is None or accuracy >= 90:
            continue

        if accuracy_type in result:
            result[accuracy_type].append(row)

    return result


# =========================================================
# INDEX / HOME
# =========================================================

@app.route("/")
def index():

    if not session.get("user_id"):
        return redirect(url_for("login"))

    role = normalize_text(
        session.get("role")
    ).lower()

    if role == "admin":
        return redirect(url_for("admin_dashboard"))

    if role == "qc":
        return redirect(url_for("qc_dashboard"))

    if role == "biller":
        return redirect(url_for("biller_dashboard"))

    session.clear()

    flash("Invalid account role. Please log in again.", "danger")

    return redirect(url_for("login"))


# =========================================================
# LOGIN
# =========================================================

@app.route("/login", methods=["GET", "POST"])
def login():

    if session.get("user_id"):
        return redirect(url_for("index"))

    if request.method == "POST":

        validate_csrf()

        username = normalize_text(
            request.form.get("username")
        )

        password = request.form.get("password", "")

        if not username or not password:
            flash(
                "Please enter your username and password.",
                "warning"
            )

            return redirect(url_for("login"))

        with db() as con:

            user = con.execute(
                """
                SELECT
                    id,
                    username,
                    password_hash,
                    role,
                    display_name,
                    active
                FROM users
                WHERE LOWER(username) = LOWER(?)
                """,
                (username,)
            ).fetchone()

        if not user or not user["active"]:

            flash(
                "Invalid username or password.",
                "danger"
            )

            return redirect(url_for("login"))

        if not check_password_hash(
            user["password_hash"],
            password
        ):

            flash(
                "Invalid username or password.",
                "danger"
            )

            return redirect(url_for("login"))

        role = normalize_text(
            user["role"]
        ).lower()

        if role not in ("admin", "biller", "qc"):

            flash(
                "This account has an invalid role.",
                "danger"
            )

            return redirect(url_for("login"))

        session.clear()

        session["user_id"] = user["id"]
        session["username"] = user["username"]
        session["role"] = role
        session["display_name"] = (
            user["display_name"]
            or user["username"]
        )
        session["csrf_token"] = secrets.token_hex(32)

        log_action(
            user["username"],
            "LOGIN",
            f"{role.upper()} logged in"
        )

        flash(
            f"Welcome, {session['display_name']}!",
            "success"
        )

        return redirect(url_for("index"))

    return render_template("login.html")



# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout", methods=["POST"])
@login_required
def logout():

    validate_csrf()

    username = session.get("username", "")

    log_action(
        username,
        "LOGOUT",
        "User logged out"
    )

    session.clear()

    flash(
        "You have been logged out.",
        "success"
    )

    return redirect(url_for("login"))




# =========================================================
# ADMIN DASHBOARD
# =========================================================

@app.route("/admin")
@login_required
@admin_required
def admin_dashboard():

    with db() as con:

        # -------------------------------------------------
        # DASHBOARD COUNTS
        # -------------------------------------------------

        total_errors_row = con.execute(
            """
            SELECT COUNT(*) AS total
            FROM error_records
            """
        ).fetchone()

        active_billers_row = con.execute(
            """
            SELECT COUNT(*) AS total
            FROM users
            WHERE LOWER(role) = 'biller'
            AND active = 1
            """
        ).fetchone()

        verified_row = con.execute(
            """
            SELECT COUNT(*) AS total
            FROM error_records
            WHERE UPPER(TRIM(COALESCE(verified_error, '')))
                = 'Y'
            """
        ).fetchone()

        disputed_row = con.execute(
            """
            SELECT COUNT(*) AS total
            FROM error_records
            WHERE UPPER(TRIM(COALESCE(verified_error, '')))
                = 'N'
            """
        ).fetchone()

        pending_row = con.execute(
            """
            SELECT COUNT(*) AS total
            FROM error_records
            WHERE verified_error IS NULL
            OR TRIM(verified_error) = ''
            """
        ).fetchone()

        imports_row = con.execute(
            """
            SELECT COUNT(*) AS total
            FROM import_history
            """
        ).fetchone()

        counts = {
            "imports": int(imports_row["total"] or 0),
            "errors": int(total_errors_row["total"] or 0),
            "billers": int(active_billers_row["total"] or 0),
            "verified": int(verified_row["total"] or 0),
            "disputed": int(disputed_row["total"] or 0),
            "pending": int(pending_row["total"] or 0)
        }

        # -------------------------------------------------
        # RECENT IMPORTS
        # -------------------------------------------------

        history = con.execute(
            """
            SELECT
                id,
                filename,
                uploaded_by,
                records_imported,
                imported_at
            FROM import_history
            ORDER BY id DESC
            LIMIT 10
            """
        ).fetchall()

    # -----------------------------------------------------
    # BILLER ACCURACY MONITORING
    # -----------------------------------------------------

    period_type = normalize_text(
        request.args.get("period")
    ).lower() or "monthly"

    selected_project = normalize_text(
        request.args.get("project_id")
    ) or "all"

    selected_period = normalize_text(
        request.args.get("date")
    )

    below_from_date = normalize_text(
        request.args.get("below_from")
    )

    below_to_date = normalize_text(
        request.args.get("below_to")
    )

    accuracy_context = build_accuracy_context(
        allowed_project_ids=None,
        selected_project=selected_project,
        period_type=period_type,
        selected_period=selected_period,
        below_from_date=below_from_date,
        below_to_date=below_to_date
    )

    # -----------------------------------------------------
    # RENDER ADMIN DASHBOARD
    # -----------------------------------------------------

    return render_template(
        "admin.html",
        counts=counts,
        history=history,
        **accuracy_context
    )

# =========================================================
# ADMIN WORKBOOK UPLOAD
# =========================================================

@app.route("/admin/upload", methods=["POST"])
@login_required
@admin_required
def upload_workbook():
    validate_csrf()
    uploaded_file = request.files.get("source_file")
    if not uploaded_file or not uploaded_file.filename:
        flash("Please select a workbook or CSV file.", "warning")
        return redirect(url_for("admin_dashboard"))

    filename = secure_filename(uploaded_file.filename)
    if not allowed_file(filename):
        flash("Unsupported file type. Allowed: ODS, XLSX, XLS, and CSV.", "danger")
        return redirect(url_for("admin_dashboard"))

    csv_sheet = normalize_text(request.form.get("csv_sheet", "ERROR")).upper()
    if csv_sheet not in ("ERROR", "INTERNAL", "EXTERNAL"):
        flash("Invalid CSV sheet selection.", "danger")
        return redirect(url_for("admin_dashboard"))

    selected_project = normalize_text(request.form.get("project_id")) or "all"
    default_project_id = None
    if selected_project != "all":
        try:
            default_project_id = int(selected_project)
        except ValueError:
            flash("Invalid project selection.", "danger")
            return redirect(url_for("admin_dashboard"))
        with db() as con:
            valid = con.execute(
                "SELECT id FROM projects WHERE id = ? AND active = 1",
                (default_project_id,)
            ).fetchone()
        if not valid:
            flash("Selected project does not exist or is inactive.", "danger")
            return redirect(url_for("admin_dashboard"))

    try:
        sheets = read_source(uploaded_file, csv_sheet=csv_sheet)
        if not sheets:
            raise ValueError("No supported data was found in the file.")
        with db() as con:
            imported = import_workbook(
                con, sheets, filename, session.get("username", ""),
                default_project_id=default_project_id,
                allowed_project_ids=None
            )
        flash(f"Import completed. {imported} records imported or updated.", "success")
    except ValueError as exc:
        flash(f"Import failed: {str(exc)}", "danger")
    except Exception:
        app.logger.exception("Admin workbook import failed")
        flash("Workbook import failed. No successful import was confirmed. Please check the workbook format and server logs.", "danger")
    return redirect(url_for("admin_dashboard"))


@app.route("/qc/upload", methods=["POST"])
@login_required
@qc_required
def qc_upload_workbook():
    validate_csrf()
    allowed_ids = get_user_project_ids(session.get("user_id"))
    if not allowed_ids:
        flash("Your QC account has no assigned active project.", "warning")
        return redirect(url_for("qc_dashboard"))

    uploaded_file = request.files.get("source_file")
    if not uploaded_file or not uploaded_file.filename:
        flash("Please select a workbook or CSV file.", "warning")
        return redirect(url_for("qc_dashboard"))

    filename = secure_filename(uploaded_file.filename)
    if not allowed_file(filename):
        flash("Unsupported file type. Allowed: ODS, XLSX, XLS, and CSV.", "danger")
        return redirect(url_for("qc_dashboard"))

    csv_sheet = normalize_text(request.form.get("csv_sheet", "ERROR")).upper()
    if csv_sheet not in ("ERROR", "INTERNAL", "EXTERNAL"):
        flash("Invalid import type.", "danger")
        return redirect(url_for("qc_dashboard"))

    selected_project = normalize_text(request.form.get("project_id")) or ""
    default_project_id = None
    if selected_project and selected_project != "all":
        try:
            default_project_id = int(selected_project)
        except ValueError:
            flash("Invalid project selection.", "danger")
            return redirect(url_for("qc_dashboard"))
        if default_project_id not in allowed_ids:
            abort(403)
    elif selected_project == "all":
        default_project_id = None
    else:
        flash("Please select an assigned project, or choose All Assigned Projects and include a PROJECT column.", "warning")
        return redirect(url_for("qc_dashboard"))

    try:
        sheets = read_source(uploaded_file, csv_sheet=csv_sheet)
        if not sheets:
            raise ValueError("No supported data was found in the file.")
        with db() as con:
            imported = import_workbook(
                con, sheets, filename, session.get("username", ""),
                default_project_id=default_project_id,
                allowed_project_ids=allowed_ids
            )
        flash(f"Import completed. {imported} records imported or updated.", "success")
    except ValueError as exc:
        flash(f"Import failed: {str(exc)}", "danger")
    except Exception:
        app.logger.exception("QC workbook import failed")
        flash("Workbook import failed. No successful import was confirmed. Please check the workbook format and server logs.", "danger")
    return redirect(url_for("qc_dashboard"))

# =========================================================
# ADMIN ACCOUNT MANAGEMENT
# =========================================================

@app.route("/admin/accounts")
@login_required
@admin_required
def admin_accounts():

    accounts = []

    with db() as con:

        users = con.execute(
            """
            SELECT
                id,
                username,
                display_name,
                role,
                active,
                created_at
            FROM users
            ORDER BY
                CASE LOWER(role)
                    WHEN 'admin' THEN 1
                    WHEN 'qc' THEN 2
                    ELSE 3
                END,
                display_name,
                username
            """
        ).fetchall()

        
        projects = con.execute(
            """
            SELECT
                id,
                project_name,
                active
            FROM projects
            ORDER BY project_name
            """
        ).fetchall()

        # Load all project assignments at once.
        assignments = con.execute(
            """
            SELECT
                up.user_id,
                p.id AS project_id,
                p.project_name
            FROM user_projects up
            INNER JOIN projects p
                ON p.id = up.project_id
            ORDER BY p.project_name
            """
        ).fetchall()

    assignment_map = {}

    for assignment in assignments:

        user_id = int(assignment["user_id"])

        assignment_map.setdefault(
            user_id,
            []
        ).append({
            "id": int(assignment["project_id"]),
            "name": assignment["project_name"]
        })

    for user in users:

        user_id = int(user["id"])

        assigned = assignment_map.get(
            user_id,
            []
        )

        assigned_ids = [
            item["id"]
            for item in assigned
        ]

        assigned_names = [
            item["name"]
            for item in assigned
        ]

        accounts.append({
            "id": user_id,
            "username": user["username"],
            "display_name": (
                user["display_name"]
                or user["username"]
            ),
            "role": normalize_text(
                user["role"]
            ).upper(),
            "active": bool(user["active"]),
            "created_at": user["created_at"],
            "projects": (
                ", ".join(assigned_names)
                if assigned_names
                else (
                    "ALL PROJECTS"
                    if normalize_text(user["role"]).lower() == "admin"
                    else "No project assigned"
                )
            ),
            "assigned_project_ids": assigned_ids
        })

    return render_template(
        "accounts.html",
        accounts=accounts,
        projects=projects
    )


# =========================================================
# CREATE ACCOUNT
# BILLER / QC / ADMIN
# =========================================================

@app.route(
    "/admin/account/create",
    methods=["POST"]
)
@login_required
@admin_required
def create_account():

    validate_csrf()

    display_name = normalize_text(
        request.form.get("display_name")
        or request.form.get("name")
    )

    username = normalize_text(
        request.form.get("username")
    )

    role = normalize_text(
        request.form.get("role")
    ).lower()

    password = request.form.get(
        "password",
        ""
    )

    project_ids = (
        request.form.getlist("project_ids")
        or request.form.getlist("projects")
        or request.form.getlist("project_id")
    )

    if not display_name or not username or not password:

        flash(
            "Please complete all required account fields.",
            "warning"
        )

        return redirect(url_for("admin_accounts"))

    if len(password) < 6:

        flash(
            "Password must contain at least 6 characters.",
            "warning"
        )

        return redirect(url_for("admin_accounts"))

    if role not in (
        "admin",
        "biller",
        "qc"
    ):

        flash(
            "Invalid account role.",
            "danger"
        )

        return redirect(url_for("admin_accounts"))

    clean_project_ids = set()

    try:
        for project_id in project_ids:
            clean_project_ids.add(int(project_id))

    except (TypeError, ValueError):
        flash(
            "Invalid project selection.",
            "danger"
        )
        return redirect(url_for("admin_accounts"))

    if role in ("biller", "qc") and not clean_project_ids:

        flash(
            "Select at least one project for this account.",
            "warning"
        )

        return redirect(url_for("admin_accounts"))

    password_hash = generate_password_hash(password)

    try:

        with db() as con:

            duplicate = con.execute(
                """
                SELECT id
                FROM users
                WHERE LOWER(username) = LOWER(?)
                """,
                (username,)
            ).fetchone()

            if duplicate:

                flash(
                    "That username already exists.",
                    "danger"
                )

                return redirect(url_for("admin_accounts"))

            if role in ("biller", "qc"):

                placeholders = ",".join(
                    ["?"] * len(clean_project_ids)
                )

                valid_projects = con.execute(
                    f"""
                    SELECT id
                    FROM projects
                    WHERE active = 1
                    AND id IN ({placeholders})
                    """,
                    tuple(sorted(clean_project_ids))
                ).fetchall()

                valid_ids = {
                    int(row["id"])
                    for row in valid_projects
                }

                if valid_ids != clean_project_ids:

                    raise ValueError(
                        "One or more selected projects are invalid or inactive."
                    )

            # RETURNING id works with PostgreSQL
            # and modern SQLite.
            inserted = con.execute(
                """
                INSERT INTO users (
                    username,
                    password_hash,
                    role,
                    display_name,
                    active,
                    created_at
                )
                VALUES (?, ?, ?, ?, 1, ?)
                RETURNING id
                """,
                (
                    username,
                    password_hash,
                    role,
                    display_name,
                    datetime.now().isoformat(timespec="seconds")
                )
            ).fetchone()

            if not inserted:

                raise RuntimeError(
                    "Unable to retrieve the new account ID."
                )

            new_user_id = int(inserted["id"])

            # Admin accounts automatically have access
            # to all active projects.
            if role in ("biller", "qc"):

                con.executemany(
                    """
                    INSERT INTO user_projects (
                        user_id,
                        project_id
                    )
                    VALUES (?, ?)
                    """,
                    [
                        (new_user_id, project_id)
                        for project_id in sorted(clean_project_ids)
                    ]
                )

            con.execute(
                """
                INSERT INTO activity_logs (
                    username,
                    action,
                    details,
                    created_at
                )
                VALUES (?, ?, ?, ?)
                """,
                (
                    session.get("username", ""),
                    "CREATE ACCOUNT",
                    f"Created {role.upper()} account: {username}",
                    datetime.now().isoformat(timespec="seconds")
                )
            )

        flash(
            f"{role.upper()} account created successfully.",
            "success"
        )

    except ValueError as exc:

        flash(
            str(exc),
            "danger"
        )

    except Exception:

        app.logger.exception(
            "Account creation failed"
        )

        flash(
            "Unable to create account. "
            "Please check the account details.",
            "danger"
        )

    return redirect(url_for("admin_accounts"))


# =========================================================
# EDIT ACCOUNT
# UPDATE ACCOUNT DETAILS AND PROJECT ASSIGNMENTS
# =========================================================

@app.route(
    "/admin/account/edit",
    methods=["POST"]
)
@login_required
@admin_required
def edit_account():

    validate_csrf()

    user_id = request.form.get("user_id", "").strip()

    display_name = normalize_text(
        request.form.get("display_name")
    )

    username = normalize_text(
        request.form.get("username")
    )

    role = normalize_text(
        request.form.get("role")
    ).lower()

    password = request.form.get("password", "")

    submitted_project_ids = request.form.getlist(
        "project_ids"
    )

    if not user_id.isdigit():

        flash("Invalid account selected.", "danger")
        return redirect(url_for("admin_accounts"))

    user_id = int(user_id)

    if not display_name or not username:

        flash(
            "Please complete all required account fields.",
            "warning"
        )
        return redirect(url_for("admin_accounts"))

    if role not in ("admin", "biller", "qc"):

        flash("Invalid account role.", "danger")
        return redirect(url_for("admin_accounts"))

    if password and len(password) < 6:

        flash(
            "Password must contain at least 6 characters.",
            "warning"
        )
        return redirect(url_for("admin_accounts"))

    clean_project_ids = set()

    try:

        for project_id in submitted_project_ids:
            clean_project_ids.add(int(project_id))

    except (TypeError, ValueError):

        flash(
            "Invalid project selection.",
            "danger"
        )
        return redirect(url_for("admin_accounts"))

    try:

        with db() as con:

            # ---------------------------------------------
            # PREVENT CONCURRENT ADMIN ROLE CHANGES
            # ---------------------------------------------

            if is_postgres():

                con.execute(
                    """
                    SELECT pg_advisory_xact_lock(746281)
                    """
                )

            # ---------------------------------------------
            # CHECK TARGET ACCOUNT
            # ---------------------------------------------

            target = con.execute(
                """
                SELECT id, username, role, active
                FROM users
                WHERE id = ?
                """,
                (user_id,)
            ).fetchone()

            if not target:

                raise ValueError(
                    "The selected account does not exist."
                )

            old_role = normalize_text(
                target["role"]
            ).lower()

            # -------------------------------------------------
            # PREVENT ADMIN SELF-DEMOTION
            # -------------------------------------------------
            
            current_user_id = session.get("user_id")
            
            if (
                old_role == "admin"
                and role != "admin"
                and str(current_user_id) == str(user_id)
            ):
                raise ValueError(
                    "You cannot change your own Admin role."
                )

            # ---------------------------------------------
            # PREVENT DUPLICATE USERNAME
            # ---------------------------------------------

            duplicate = con.execute(
                """
                SELECT id
                FROM users
                WHERE LOWER(username) = LOWER(?)
                AND id != ?
                """,
                (username, user_id)
            ).fetchone()

            if duplicate:

                raise ValueError(
                    "That username is already being used."
                )

            # ---------------------------------------------
            # PREVENT DEMOTING THE LAST ACTIVE ADMIN
            # ---------------------------------------------

            if (
                old_role == "admin"
                and role != "admin"
                and target["active"]
            ):

                active_admins = con.execute(
                    """
                    SELECT COUNT(*) AS total
                    FROM users
                    WHERE LOWER(role) = 'admin'
                    AND active = 1
                    """
                ).fetchone()

                if int(active_admins["total"]) <= 1:

                    raise ValueError(
                        "Cannot change the role of the last "
                        "active Admin account."
                    )

            # ---------------------------------------------
            # VALIDATE PROJECT ASSIGNMENTS
            # ---------------------------------------------

            final_project_ids = set()

            if role in ("biller", "qc"):

                # Submitted assignments must be active.
                if clean_project_ids:

                    placeholders = ",".join(
                        ["?"] * len(clean_project_ids)
                    )

                    valid_projects = con.execute(
                        f"""
                        SELECT id
                        FROM projects
                        WHERE active = 1
                        AND id IN ({placeholders})
                        """,
                        tuple(sorted(clean_project_ids))
                    ).fetchall()

                    valid_ids = {
                        int(row["id"])
                        for row in valid_projects
                    }

                    if valid_ids != clean_project_ids:

                        raise ValueError(
                            "One or more selected projects "
                            "are invalid or inactive."
                        )

                # Preserve previously assigned inactive projects.
                existing_assignments = con.execute(
                    """
                    SELECT up.project_id
                    FROM user_projects up
                    INNER JOIN projects p
                        ON p.id = up.project_id
                    WHERE up.user_id = ?
                    AND p.active = 0
                    """,
                    (user_id,)
                ).fetchall()

                preserved_inactive_ids = {
                    int(row["project_id"])
                    for row in existing_assignments
                }

                final_project_ids = (
                    clean_project_ids |
                    preserved_inactive_ids
                )

                if not final_project_ids:

                    raise ValueError(
                        "Select at least one project "
                        "for QC and Biller accounts."
                    )

            # Admin accounts use global active-project access.
            # Their individual project assignments are cleared.
            # ---------------------------------------------
            # UPDATE ACCOUNT DETAILS
            # ---------------------------------------------

            if password:

                password_hash = generate_password_hash(
                    password
                )

                con.execute(
                    """
                    UPDATE users
                    SET
                        display_name = ?,
                        username = ?,
                        role = ?,
                        password_hash = ?
                    WHERE id = ?
                    """,
                    (
                        display_name,
                        username,
                        role,
                        password_hash,
                        user_id
                    )
                )

            else:

                con.execute(
                    """
                    UPDATE users
                    SET
                        display_name = ?,
                        username = ?,
                        role = ?
                    WHERE id = ?
                    """,
                    (
                        display_name,
                        username,
                        role,
                        user_id
                    )
                )

            # ---------------------------------------------
            # UPDATE PROJECT ASSIGNMENTS
            # ---------------------------------------------

            con.execute(
                """
                DELETE FROM user_projects
                WHERE user_id = ?
                """,
                (user_id,)
            )

            if role in ("biller", "qc"):

                con.executemany(
                    """
                    INSERT INTO user_projects (
                        user_id,
                        project_id
                    )
                    VALUES (?, ?)
                    """,
                    [
                        (user_id, project_id)
                        for project_id in sorted(final_project_ids)
                    ]
                )

            # ---------------------------------------------
            # ACTIVITY LOG
            # ---------------------------------------------

            con.execute(
                """
                INSERT INTO activity_logs (
                    username,
                    action,
                    details,
                    created_at
                )
                VALUES (?, ?, ?, ?)
                """,
                (
                    session.get("username", ""),
                    "EDIT ACCOUNT",
                    f"Updated account: {username} "
                    f"({old_role.upper()} to {role.upper()})",
                    datetime.now().isoformat(timespec="seconds")
                )
            )

        flash(
            f"Account '{username}' updated successfully.",
            "success"
        )

    except ValueError as exc:

        flash(str(exc), "danger")

    except Exception:

        app.logger.exception(
            "Account editing failed"
        )

        flash(
            "Unable to update account. "
            "Please check the account details.",
            "danger"
        )

    return redirect(url_for("admin_accounts"))


# =========================================================
# TOGGLE ACCOUNT STATUS
# ACTIVATE / DEACTIVATE ACCOUNT
# =========================================================

@app.route(
    "/admin/account/toggle-status",
    methods=["POST"]
)
@login_required
@admin_required
def toggle_account_status():

    validate_csrf()

    user_id = request.form.get(
        "user_id", ""
    ).strip()

    if not user_id.isdigit():

        flash(
            "Invalid account selected.",
            "danger"
        )
        return redirect(url_for("admin_accounts"))

    user_id = int(user_id)

    try:

        with db() as con:

            # ---------------------------------------------
            # PREVENT CONCURRENT ADMIN STATUS CHANGES
            # ---------------------------------------------

            if is_postgres():

                con.execute(
                    """
                    SELECT pg_advisory_xact_lock(746281)
                    """
                )

            # ---------------------------------------------
            # GET TARGET ACCOUNT
            # ---------------------------------------------

            target = con.execute(
                """
                SELECT
                    id,
                    username,
                    role,
                    active
                FROM users
                WHERE id = ?
                """,
                (user_id,)
            ).fetchone()

            if not target:

                raise ValueError(
                    "The selected account does not exist."
                )

            username = target["username"]

            role = normalize_text(
                target["role"]
            ).lower()

            currently_active = bool(
                target["active"]
            )

            # ---------------------------------------------
            # DETERMINE NEW STATUS
            # ---------------------------------------------

            new_status = 0 if currently_active else 1

            # ---------------------------------------------
            # PREVENT SELF-DEACTIVATION
            # ---------------------------------------------

            current_user_id = session.get("user_id")

            if (
                currently_active
                and str(current_user_id) == str(user_id)
            ):

                raise ValueError(
                    "You cannot deactivate your own account."
                )

            # ---------------------------------------------
            # PREVENT DEACTIVATING THE LAST ACTIVE ADMIN
            # ---------------------------------------------

            if currently_active and role == "admin":

                active_admins = con.execute(
                    """
                    SELECT COUNT(*) AS total
                    FROM users
                    WHERE LOWER(role) = 'admin'
                    AND active = 1
                    """
                ).fetchone()

                if int(active_admins["total"]) <= 1:

                    raise ValueError(
                        "Cannot deactivate the last active "
                        "Admin account."
                    )

            # ---------------------------------------------
            # UPDATE ACCOUNT STATUS
            # ---------------------------------------------

            con.execute(
                """
                UPDATE users
                SET active = ?
                WHERE id = ?
                """,
                (
                    new_status,
                    user_id
                )
            )

            # ---------------------------------------------
            # ACTIVITY LOG
            # ---------------------------------------------

            action = (
                "ACTIVATE ACCOUNT"
                if new_status
                else "DEACTIVATE ACCOUNT"
            )

            con.execute(
                """
                INSERT INTO activity_logs (
                    username,
                    action,
                    details,
                    created_at
                )
                VALUES (?, ?, ?, ?)
                """,
                (
                    session.get("username", ""),
                    action,
                    f"{action.title()}: {username} "
                    f"(Account ID: {user_id})",
                    datetime.now().isoformat(timespec="seconds")
                )
            )

        # ---------------------------------------------
        # SUCCESS MESSAGE
        # ---------------------------------------------

        if new_status:

            flash(
                f"Account '{username}' activated successfully.",
                "success"
            )

        else:

            flash(
                f"Account '{username}' deactivated successfully.",
                "success"
            )

    except ValueError as exc:

        flash(
            str(exc),
            "danger"
        )

    except Exception:

        app.logger.exception(
            "Account status update failed"
        )

        flash(
            "Unable to update account status.",
            "danger"
        )

    return redirect(url_for("admin_accounts"))

# =========================================================
# UPDATE ACCOUNT PROJECT ASSIGNMENTS
# =========================================================

@app.route(
    "/admin/account/<int:user_id>/projects",
    methods=["POST"]
)
@login_required
@admin_required
def update_account_projects(user_id):

    validate_csrf()

    project_ids = (
        request.form.getlist("project_ids")
        or request.form.getlist("projects")
        or request.form.getlist("project_id")
    )

    clean_project_ids = set()

    for project_id in project_ids:

        try:
            clean_project_ids.add(
                int(project_id)
            )
        except (TypeError, ValueError):
            continue

    try:

        with db() as con:

            user = con.execute(
                """
                SELECT id, username, role
                FROM users
                WHERE id = ?
                """,
                (user_id,)
            ).fetchone()

            if not user:
                abort(404)

            role = normalize_text(
                user["role"]
            ).lower()

            if role == "admin":

                flash(
                    "Admin accounts automatically have access to all projects.",
                    "info"
                )

                return redirect(
                    url_for("admin_accounts") + f"#account-{user_id}"
                )

            if role not in ("biller", "qc"):
                abort(400)

            if not clean_project_ids:

                raise ValueError(
                    "Select at least one project."
                )

            placeholders = ",".join(
                ["?"] * len(clean_project_ids)
            )

            valid_projects = con.execute(
                f"""
                SELECT id
                FROM projects
                WHERE active = 1
                AND id IN ({placeholders})
                """,
                tuple(sorted(clean_project_ids))
            ).fetchall()

            valid_ids = {
                int(row["id"])
                for row in valid_projects
            }

            if valid_ids != clean_project_ids:

                raise ValueError(
                    "One or more selected projects are invalid or inactive."
                )

            # Replace only the account's assignments.
            con.execute(
                """
                DELETE FROM user_projects
                WHERE user_id = ?
                """,
                (user_id,)
            )

            con.executemany(
                """
                INSERT INTO user_projects (
                    user_id,
                    project_id
                )
                VALUES (?, ?)
                """,
                [
                    (user_id, project_id)
                    for project_id in sorted(clean_project_ids)
                ]
            )

            con.execute(
                """
                INSERT INTO activity_logs (
                    username,
                    action,
                    details,
                    created_at
                )
                VALUES (?, ?, ?, ?)
                """,
                (
                    session.get("username", ""),
                    "UPDATE ACCOUNT PROJECTS",
                    (
                        f"Updated projects for "
                        f"{user['username']}"
                    ),
                    datetime.now().isoformat(timespec="seconds")
                )
            )

        flash(
            "Account project assignments updated successfully.",
            "success"
        )

    except ValueError as exc:

        flash(
            str(exc),
            "danger"
        )

    except Exception:

        app.logger.exception(
            "Account project update failed"
        )

        flash(
            "Unable to update account projects.",
            "danger"
        )

    return redirect(
        url_for("admin_accounts") + f"#account-{user_id}"
    )


# =========================================================
# ACTIVATE / DEACTIVATE ACCOUNT
# =========================================================

@app.route(
    "/admin/account/<int:user_id>/toggle",
    methods=["POST"]
)
@login_required
@admin_required
def toggle_account(user_id):

    validate_csrf()

    current_user_id = int(
        session.get("user_id")
    )

    try:

        with db() as con:

            user = con.execute(
                """
                SELECT
                    id,
                    username,
                    role,
                    active
                FROM users
                WHERE id = ?
                """,
                (user_id,)
            ).fetchone()

            if not user:
                abort(404)

            is_active = bool(user["active"])

            role = normalize_text(
                user["role"]
            ).lower()

            # Prevent self-deactivation.
            if user_id == current_user_id and is_active:

                flash(
                    "You cannot deactivate your own account.",
                    "danger"
                )

                return redirect(url_for("admin_accounts"))

            # Prevent deactivating the last active Admin.
            if role == "admin" and is_active:

                active_admins = con.execute(
                    """
                    SELECT COUNT(*) AS total
                    FROM users
                    WHERE LOWER(role) = 'admin'
                    AND active = 1
                    """
                ).fetchone()

                if int(active_admins["total"] or 0) <= 1:

                    flash(
                        "The last active Admin account cannot be deactivated.",
                        "danger"
                    )

                    return redirect(url_for("admin_accounts"))

            new_status = 0 if is_active else 1

            con.execute(
                """
                UPDATE users
                SET active = ?
                WHERE id = ?
                """,
                (
                    new_status,
                    user_id
                )
            )

            con.execute(
                """
                INSERT INTO activity_logs (
                    username,
                    action,
                    details,
                    created_at
                )
                VALUES (?, ?, ?, ?)
                """,
                (
                    session.get("username", ""),
                    (
                        "DEACTIVATE ACCOUNT"
                        if new_status == 0
                        else "ACTIVATE ACCOUNT"
                    ),
                    f"{user['username']} account status changed",
                    datetime.now().isoformat(timespec="seconds")
                )
            )

        flash(
            (
                "Account deactivated successfully."
                if new_status == 0
                else "Account activated successfully."
            ),
            "success"
        )

    except Exception:

        app.logger.exception(
            "Account status update failed"
        )

        flash(
            "Unable to update account status.",
            "danger"
        )

    return redirect(url_for("admin_accounts"))


# =========================================================
# ADMIN PROJECT MANAGEMENT
# =========================================================

@app.route("/admin/projects")
@login_required
@admin_required
def admin_projects():

    project_list = []

    with db() as con:

        projects = con.execute(
            """
            SELECT
                id,
                project_name,
                name,
                role,
                active,
                created_at
            FROM projects
            ORDER BY project_name
            """
        ).fetchall()

        assignments = con.execute(
            """
            SELECT
                up.project_id,
                u.id AS user_id,
                u.display_name,
                u.username,
                u.role,
                u.active
            FROM user_projects up
            INNER JOIN users u
                ON u.id = up.user_id
            ORDER BY u.display_name, u.username
            """
        ).fetchall()

    assignment_map = {}

    for assignment in assignments:

        project_id = int(
            assignment["project_id"]
        )

        assignment_map.setdefault(
            project_id,
            []
        ).append({
            "user_id": int(assignment["user_id"]),
            "display_name": (
                assignment["display_name"]
                or assignment["username"]
            ),
            "username": assignment["username"],
            "role": normalize_text(
                assignment["role"]
            ).upper(),
            "active": bool(assignment["active"])
        })

    for project in projects:

        project_id = int(project["id"])

        assigned_accounts = assignment_map.get(
            project_id,
            []
        )

        account_names = []

        for account in assigned_accounts:

            status = (
                ""
                if account["active"]
                else " (INACTIVE)"
            )

            account_names.append(
                f"{account['display_name']} "
                f"({account['role']}){status}"
            )

        project_list.append({
            "id": project_id,
            "project_name": project["project_name"],
            "name": project["name"] or "",
            "role": normalize_text(
                project["role"]
            ).upper(),
            "active": bool(project["active"]),
            "created_at": project["created_at"],
            "assigned_accounts": (
                ", ".join(account_names)
                if account_names
                else "No accounts assigned"
            )
        })

    return render_template(
        "projects.html",
        projects=project_list
    )


# =========================================================
# CREATE PROJECT
# =========================================================

@app.route(
    "/admin/projects",
    methods=["POST"]
)
@app.route(
    "/admin/project/create",
    methods=["POST"]
)
@login_required
@admin_required
def create_project():

    validate_csrf()

    project_name = normalize_text(
        request.form.get("project_name")
    )

    if not project_name:

        flash(
            "Project Name is required.",
            "warning"
        )

        return redirect(url_for("admin_projects"))

    try:

        with db() as con:

            existing = con.execute(
                """
                SELECT id
                FROM projects
                WHERE LOWER(TRIM(project_name))
                    = LOWER(TRIM(?))
                """,
                (project_name,)
            ).fetchone()

            if existing:

                flash(
                    "A project with that name already exists.",
                    "danger"
                )

                return redirect(url_for("admin_projects"))

            # Keep legacy columns for database compatibility, but
            # project creation no longer requires or assigns a role/name.
            con.execute(
                """
                INSERT INTO projects (
                    project_name,
                    name,
                    role,
                    active,
                    created_at
                )
                VALUES (?, NULL, '', 1, ?)
                """,
                (
                    project_name,
                    datetime.now().isoformat(timespec="seconds")
                )
            )

            con.execute(
                """
                INSERT INTO activity_logs (
                    username,
                    action,
                    details,
                    created_at
                )
                VALUES (?, ?, ?, ?)
                """,
                (
                    session.get("username", ""),
                    "CREATE PROJECT",
                    f"Created project: {project_name}",
                    datetime.now().isoformat(timespec="seconds")
                )
            )

        flash(
            "Project created successfully.",
            "success"
        )

    except Exception:

        app.logger.exception(
            "Project creation failed"
        )

        flash(
            "Unable to create project. Please check the project details.",
            "danger"
        )

    return redirect(url_for("admin_projects"))


# =========================================================
# EDIT PROJECT
# =========================================================

@app.route(
    "/admin/project/<int:project_id>/edit",
    methods=["POST"]
)
@login_required
@admin_required
def edit_project(project_id):

    validate_csrf()

    project_name = normalize_text(
        request.form.get("project_name")
    )

    if not project_name:

        flash(
            "Project Name is required.",
            "warning"
        )

        return redirect(url_for("admin_projects"))

    try:

        with db() as con:

            project = con.execute(
                """
                SELECT id, project_name
                FROM projects
                WHERE id = ?
                """,
                (project_id,)
            ).fetchone()

            if not project:
                abort(404)

            duplicate = con.execute(
                """
                SELECT id
                FROM projects
                WHERE LOWER(TRIM(project_name))
                    = LOWER(TRIM(?))
                AND id <> ?
                """,
                (project_name, project_id)
            ).fetchone()

            if duplicate:

                flash(
                    "Another project already uses that name.",
                    "danger"
                )

                return redirect(url_for("admin_projects"))

            old_name = project["project_name"]

            # Only update the project name. Legacy name/role columns
            # remain untouched for database compatibility.
            con.execute(
                """
                UPDATE projects
                SET project_name = ?
                WHERE id = ?
                """,
                (project_name, project_id)
            )

            con.execute(
                """
                INSERT INTO activity_logs (
                    username,
                    action,
                    details,
                    created_at
                )
                VALUES (?, ?, ?, ?)
                """,
                (
                    session.get("username", ""),
                    "EDIT PROJECT",
                    f"Updated project {old_name} to {project_name}",
                    datetime.now().isoformat(timespec="seconds")
                )
            )

        flash(
            "Project updated successfully.",
            "success"
        )

    except Exception:

        app.logger.exception(
            "Project update failed"
        )

        flash(
            "Unable to update project.",
            "danger"
        )

    return redirect(url_for("admin_projects"))


# =========================================================
# ACTIVATE / DEACTIVATE PROJECT
# =========================================================

@app.route(
    "/admin/project/<int:project_id>/toggle",
    methods=["POST"]
)
@login_required
@admin_required
def toggle_project(project_id):

    validate_csrf()

    try:

        with db() as con:

            project = con.execute(
                """
                SELECT id, project_name, active
                FROM projects
                WHERE id = ?
                """,
                (project_id,)
            ).fetchone()

            if not project:
                abort(404)

            new_status = (
                0
                if bool(project["active"])
                else 1
            )

            con.execute(
                """
                UPDATE projects
                SET active = ?
                WHERE id = ?
                """,
                (
                    new_status,
                    project_id
                )
            )

            con.execute(
                """
                INSERT INTO activity_logs (
                    username,
                    action,
                    details,
                    created_at
                )
                VALUES (?, ?, ?, ?)
                """,
                (
                    session.get("username", ""),
                    (
                        "DEACTIVATE PROJECT"
                        if new_status == 0
                        else "ACTIVATE PROJECT"
                    ),
                    (
                        f"{project['project_name']} "
                        f"status changed"
                    ),
                    datetime.now().isoformat(timespec="seconds")
                )
            )

        flash(
            (
                "Project deactivated. Historical records were preserved."
                if new_status == 0
                else "Project activated successfully."
            ),
            "success"
        )

    except Exception:

        app.logger.exception(
            "Project status update failed"
        )

        flash(
            "Unable to update project status.",
            "danger"
        )

    return redirect(url_for("admin_projects"))


# =========================================================
# ADMIN ERROR RECORD LIST
# Shows all records, including unassigned records.
# =========================================================

@app.route("/admin/errors")
@login_required
@admin_required
def admin_errors():

    with db() as con:

        records = con.execute(
            """
            SELECT
                record_key,
                billed_date,
                week,
                biller_name,
                username,
                pro,
                source,
                auditor,
                audit_type,
                project_id
            FROM error_records
            ORDER BY
                CASE
                    WHEN billed_date IS NULL
                        OR billed_date = ''
                    THEN 1
                    ELSE 0
                END,
                billed_date DESC,
                record_key DESC
            LIMIT 250
            """
        ).fetchall()

    return render_template(
        "admin_errors.html",
        records=records
    )

# =========================================================
# ADMIN ERROR DETAILS
# Read-only view of all error records.
# =========================================================

@app.route("/admin/error-details")
@login_required
@admin_required
def admin_error_details():

    selected_date = normalize_text(
        request.args.get("date")
    )

    with db() as con:

        date_rows = con.execute(
            """
            SELECT DISTINCT billed_date
            FROM error_records
            WHERE billed_date IS NOT NULL
            AND TRIM(billed_date) <> ''
            ORDER BY billed_date DESC
            """
        ).fetchall()

        dates = [
            row["billed_date"]
            for row in date_rows
        ]

        if selected_date and selected_date not in dates:

            flash(
                "The selected date is invalid.",
                "warning"
            )

            selected_date = ""

        # Automatically select the latest date
        # when Admin first opens the page.
        if not selected_date:
            selected_date = dates[0] if dates else ""

        if selected_date:

            rows = con.execute(
                """
                SELECT
                    e.record_key,
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
                    e.verified_error,
                    e.biller_comment,
                    e.qc_explanation,
                    e.qc_name,
                    e.updated_at,
                    e.project_id,
                    p.project_name
                FROM error_records e
                LEFT JOIN projects p
                    ON p.id = e.project_id
                WHERE e.billed_date = ?
                ORDER BY e.source_row DESC
                """,
                (selected_date,)
            ).fetchall()

        else:

            rows = []

    errors = [
        dict(row)
        for row in rows
    ]

    projects = []
    with db() as con:
        placeholders = ",".join(["?"] * len(project_ids))
        project_rows = con.execute(
            f"SELECT id, project_name FROM projects WHERE active = 1 AND id IN ({placeholders}) ORDER BY project_name",
            tuple(project_ids)
        ).fetchall()
        projects = [dict(row) for row in project_rows]

    return render_template(
        "errors.html",
        errors=errors,
        dates=dates,
        selected_date=selected_date,
        projects=projects
    )

# =========================================================
# ADMIN EDIT ORIGINAL ERROR RECORD DETAILS
# Does not modify Biller or QC responses.
# =========================================================

@app.route(
    "/admin/error/<record_key>",
    methods=["GET", "POST"]
)
@login_required
@admin_required
def edit_error(record_key):

    with db() as con:

        record = con.execute(
            """
            SELECT *
            FROM error_records
            WHERE record_key = ?
            """,
            (record_key,)
        ).fetchone()

        if not record:
            abort(404)

        if request.method == "POST":

            validate_csrf()

            billed_date_input = normalize_text(
                request.form.get("billed_date")
            )

            parsed_date = (
                parse_date(billed_date_input)
                if billed_date_input
                else None
            )

            if billed_date_input and not parsed_date:

                flash(
                    "Please enter a valid billed date.",
                    "danger"
                )

                return redirect(
                    url_for(
                        "edit_error",
                        record_key=record_key
                    )
                )

            billed_date = (
                parsed_date.isoformat()
                if parsed_date
                else ""
            )

            week = normalize_text(
                request.form.get("week")
            )

            if not week and billed_date:
                week = week_for_billed_date(billed_date)

            source = normalize_text(
                request.form.get("source")
            )

            audit_type = normalize_text(
                request.form.get("audit_type")
            )

            biller_name = normalize_text(
                request.form.get("biller_name")
            )

            username = normalize_text(
                request.form.get("username")
            )

            id_number = normalize_text(
                request.form.get("id_number")
            )

            auditor = normalize_text(
                request.form.get("auditor")
            )

            pro = normalize_text(
                request.form.get("pro")
            )

            zb_field = normalize_text(
                request.form.get("zb_field")
            )

            error_details = normalize_text(
                request.form.get("error_details")
            )

            if not username or not pro:

                flash(
                    "Username and PRO are required.",
                    "warning"
                )

                return redirect(
                    url_for(
                        "edit_error",
                        record_key=record_key
                    )
                )

            # Only original audit fields are updated.
            # Do not update:
            # verified_error
            # biller_comment
            # qc_explanation
            # qc_name
            # project_id
            con.execute(
                """
                UPDATE error_records
                SET
                    billed_date = ?,
                    week = ?,
                    source = ?,
                    audit_type = ?,
                    biller_name = ?,
                    username = ?,
                    id_number = ?,
                    auditor = ?,
                    pro = ?,
                    zb_field = ?,
                    error_details = ?
                WHERE record_key = ?
                """,
                (
                    billed_date,
                    week,
                    source,
                    audit_type,
                    biller_name,
                    username,
                    id_number,
                    auditor,
                    pro,
                    zb_field,
                    error_details,
                    record_key
                )
            )

            con.execute(
                """
                INSERT INTO activity_logs (
                    username,
                    action,
                    details,
                    created_at
                )
                VALUES (?, ?, ?, ?)
                """,
                (
                    session.get("username", ""),
                    "EDIT ERROR RECORD",
                    f"Updated original audit fields for {record_key}",
                    datetime.now().isoformat(timespec="seconds")
                )
            )

            flash(
                "Error record updated. Biller and QC responses were preserved.",
                "success"
            )

            return redirect(
                url_for("admin_errors")
            )

    return render_template(
        "edit_error.html",
        record=record
    )


# =========================================================
# ACCURACY HELPERS
# =========================================================

ACCURACY_TYPES = (
    "INTERNAL",
    "EXTERNAL_WC",
    "EXTERNAL_WOC"
)



def normalize_accuracy_type(value):

    value = normalize_text(value).upper()

    value = value.replace("-", "_")
    value = value.replace(" ", "_")

    # INTERNAL
    if value in (
        "INTERNAL",
        "INTERNAL_ACCURACY",
        "ACCURACY"
    ):
        return "INTERNAL"

    # EXTERNAL WITH CLAIMS
    if value in (
        "EXTERNAL_WC",
        "EXTERNAL_WC_ACCURACY",
        "WC",
        "PER_AUDITED_ACCURACY"
    ):
        return "EXTERNAL_WC"

    # EXTERNAL WITHOUT CLAIMS
    if value in (
        "EXTERNAL_WOC",
        "EXTERNAL_WOC_ACCURACY",
        "WOC",
        "PER_AUDITED_ACCURACY___WOC"
    ):
        return "EXTERNAL_WOC"

    return None

def parse_accuracy_date(value, month_value=None):
    """Parse full dates and short dates using the row's month and year."""

    parsed = parse_date(value)

    if parsed is not None:
        return parsed

    raw_date = normalize_text(value)
    raw_month = normalize_text(month_value)

    if not raw_date or not raw_month:
        return None

    # Support both 2-digit and 4-digit years.
    year_formats = (
        "%b - %y",
        "%B - %y",
        "%b %y",
        "%B %y",
        "%b - %Y",
        "%B - %Y",
        "%b-%Y",
        "%B-%Y",
        "%b %Y",
        "%B %Y",
        "%Y-%m"
    )

    year = None

    for fmt in year_formats:
        try:
            year = datetime.strptime(
                raw_month,
                fmt
            ).year
            break
        except (TypeError, ValueError):
            continue

    if year is None:
        return None

    # Support short date values such as Nov - 02.
    date_formats = (
        "%b - %d",
        "%B - %d",
        "%b-%d",
        "%B-%d",
        "%b %d",
        "%B %d"
    )

    for fmt in date_formats:
        try:
            month_day = datetime.strptime(
                raw_date,
                fmt
            )

            return date(
                year,
                month_day.month,
                month_day.day
            )

        except (TypeError, ValueError):
            continue

    return None

def accuracy_period_key(value, period_type, month_value=None):
    """Return a stable period key, resolving short dates using the row month."""
    parsed = parse_accuracy_date(value, month_value)

    if parsed is None:
        raw_month = normalize_text(month_value)
        # If record_date is absent but MONTH is present, monthly summaries
        # can still include the row without inventing a specific day.
        if period_type == "monthly" and raw_month:
            for fmt in ("%Y-%m", "%b - %y", "%B - %y", "%b %y", "%B %y"):
                try:
                    month_date = datetime.strptime(raw_month, fmt).date()
                    return month_date.strftime("%Y-%m")
                except (TypeError, ValueError):
                    continue

    if not parsed:
        return None

    if period_type == "daily":
        return parsed.isoformat()
    if period_type == "weekly":
        monday = parsed - timedelta(days=parsed.weekday())
        return monday.isoformat()
    if period_type == "monthly":
        return parsed.strftime("%Y-%m")
    return None


def accuracy_period_dates(period_key, period_type):

    if not period_key:
        return None, None

    try:

        if period_type == "daily":

            start = date.fromisoformat(period_key)
            return start, start

        if period_type == "weekly":

            start = date.fromisoformat(period_key)

            return (
                start,
                start + timedelta(days=6)
            )

        if period_type == "monthly":

            year, month = map(
                int,
                period_key.split("-")
            )

            start = date(year, month, 1)

            if month == 12:

                end = date(year + 1, 1, 1) - timedelta(days=1)

            else:

                end = date(year, month + 1, 1) - timedelta(days=1)

            return start, end

    except (ValueError, TypeError):

        return None, None

    return None, None


def previous_accuracy_period(period_key, period_type):

    start, end = accuracy_period_dates(
        period_key,
        period_type
    )

    if not start:
        return None

    if period_type == "daily":

        previous = start - timedelta(days=1)

        return previous.isoformat()

    if period_type == "weekly":

        previous = start - timedelta(days=7)

        return previous.isoformat()

    if period_type == "monthly":

        if start.month == 1:

            previous = date(
                start.year - 1,
                12,
                1
            )

        else:

            previous = date(
                start.year,
                start.month - 1,
                1
            )

        return previous.strftime("%Y-%m")

    return None


def format_accuracy(value):

    if value is None:
        return "—"

    try:

        return f"{float(value):.2f}%"

    except (ValueError, TypeError):

        return "—"


def calculate_weighted_accuracy(audited, errors):

    audited = number_value(audited) or 0
    errors = number_value(errors) or 0

    if audited <= 0:
        return None

    return (
        (audited - errors)
        / audited
    ) * 100


# =========================================================
# GET ACCURACY RECORDS
# =========================================================

def get_accuracy_records(project_ids=None):

    with db() as con:

        base_query = """
            SELECT
                id,
                sheet,
                username,
                biller_name,
                month,
                weekly,
                record_date,
                accuracy_type,
                accuracy_value,
                bills_processed,
                audited,
                errors,
                source_row,
                project_id,
                imported_at
            FROM accuracy_records
            WHERE record_date IS NOT NULL
              AND TRIM(record_date) != ''
        """

        params = ()

        if project_ids is not None:

            project_ids = [
                int(project_id)
                for project_id in project_ids
            ]

            if not project_ids:
                return []

            placeholders = ", ".join(
                ["?"] * len(project_ids)
            )

            base_query += f"""
                AND project_id IN ({placeholders})
            """

            params = tuple(project_ids)

        base_query += " ORDER BY record_date"

        rows = con.execute(
            base_query,
            params
        ).fetchall()

    return [
        dict(row)
        for row in rows
    ]

# =========================================================
# BUILD ACCURACY SUMMARY
# =========================================================


def build_accuracy_context(
    allowed_project_ids=None,
    selected_project="all",
    period_type="monthly",
    selected_period=None,
    below_from_date=None,
    below_to_date=None,
    below_category="all"
):
    if period_type not in ("daily", "weekly", "monthly"):
        period_type = "monthly"

    
    # -----------------------------------------------------
    # INDEPENDENT BELOW 90% DATE RANGE
    # -----------------------------------------------------

    try:
        below_from = (
            date.fromisoformat(below_from_date)
            if below_from_date else None
        )
    except (TypeError, ValueError):
        below_from = None

    try:
        below_to = (
            date.fromisoformat(below_to_date)
            if below_to_date else None
        )
    except (TypeError, ValueError):
        below_to = None

    if below_from and below_to and below_from > below_to:
        below_from, below_to = below_to, below_from
    
    below_category = normalize_text(
        below_category
    ).lower() or "all"

    if below_category not in ("all", "internal", "external"):
        below_category = "all"

    allowed_set = None
    if allowed_project_ids is not None:
        allowed_set = {int(value) for value in allowed_project_ids}

    records = get_accuracy_records(
        sorted(allowed_set) if allowed_set is not None else None
    )

    with db() as con:
        project_rows = con.execute(
            "SELECT id, project_name, active FROM projects ORDER BY project_name"
        ).fetchall()

    project_names = {int(row["id"]): row["project_name"] for row in project_rows}

    # Validate the requested project against the viewer's permitted set.
    if selected_project != "all":
        try:
            selected_id = int(selected_project)
        except (TypeError, ValueError):
            selected_id = None
        if selected_id is None or (allowed_set is not None and selected_id not in allowed_set):
            selected_project = "all"
        else:
            records = [
                row for row in records
                if row.get("project_id") is not None
                and int(row["project_id"]) == selected_id
            ]


    
    # -----------------------------------------------------
    # GET ALL ERROR DATES
    # Use BILLED DATE from ERROR sheet
    # -----------------------------------------------------

    error_conditions = [
        "billed_date IS NOT NULL",
        "TRIM(billed_date) != ''"
    ]

    error_params = []

    if allowed_set is not None:

        if not allowed_set:
            error_rows = []

        else:
            placeholders = ", ".join(
                ["?"] * len(allowed_set)
            )

            error_conditions.append(
                f"project_id IN ({placeholders})"
            )

            error_params.extend(sorted(allowed_set))

            if selected_project != "all":
                error_conditions.append("project_id = ?")
                error_params.append(int(selected_project))

            with db() as con:

                error_rows = con.execute(
                    f"""
                    SELECT billed_date, project_id, imported_at
                    FROM error_records
                    WHERE {" AND ".join(error_conditions)}
                    """,
                    tuple(error_params)
                ).fetchall()

    else:

        if selected_project != "all":
            error_conditions.append("project_id = ?")
            error_params.append(int(selected_project))

        with db() as con:

            error_rows = con.execute(
                f"""
                SELECT billed_date, project_id, imported_at
                FROM error_records
                WHERE {" AND ".join(error_conditions)}
                """,
                tuple(error_params)
            ).fetchall()

    # -----------------------------------------------------
    # BUILD AVAILABLE PERIODS
    # Show periods from the latest import per project/sheet.
    # Keep all historical records for period comparisons.
    # -----------------------------------------------------

    period_keys = set()
    row_periods = {}

    # Map period keys for ALL historical accuracy records.
    for row in records:
        key = accuracy_period_key(
            row.get("record_date"),
            period_type,
            row.get("month")
        )

        if key:
            row_periods[id(row)] = key

    # Find latest accuracy import per project and sheet.
    latest_accuracy_imports = {}

    for row in records:
        imported_at = normalize_text(row.get("imported_at"))
        if not imported_at:
            continue

        import_key = (
            row.get("project_id"),
            normalize_text(row.get("sheet")).upper()
        )

        if imported_at > latest_accuracy_imports.get(import_key, ""):
            latest_accuracy_imports[import_key] = imported_at

    # Add periods only from the latest accuracy imports.
    for row in records:
        import_key = (
            row.get("project_id"),
            normalize_text(row.get("sheet")).upper()
        )

        if (
            normalize_text(row.get("imported_at"))
            != latest_accuracy_imports.get(import_key)
        ):
            continue

        key = accuracy_period_key(
            row.get("record_date"),
            period_type,
            row.get("month")
        )

        if key:
            period_keys.add(key)

    # Find latest ERROR import per project.
    latest_error_imports = {}

    for row in error_rows:
        imported_at = normalize_text(row["imported_at"])
        if not imported_at:
            continue

        project_id = row["project_id"]

        if imported_at > latest_error_imports.get(project_id, ""):
            latest_error_imports[project_id] = imported_at

    # Add periods only from the latest ERROR imports.
    for row in error_rows:
        project_id = row["project_id"]

        if (
            normalize_text(row["imported_at"])
            != latest_error_imports.get(project_id)
        ):
            continue

        key = accuracy_period_key(
            row["billed_date"],
            period_type
        )

        if key:
            period_keys.add(key)

    periods = sorted(period_keys)

    if selected_period not in period_keys:
        selected_period = periods[-1] if periods else ""

    previous_period = previous_accuracy_period(
        selected_period,
        period_type
    )

    def aggregate_record(target_map, group_key, row):
        username = normalize_text(row.get("username"))
        project_id = row.get("project_id")
        project_id = int(project_id) if project_id is not None else None
        accuracy_type = normalize_accuracy_type(row.get("accuracy_type"))
        item = target_map.setdefault(group_key, {
            "username": username,
            "biller_name": normalize_text(row.get("biller_name")) or username,
            "project_id": project_id,
            "project_name": project_names.get(project_id, "Unassigned") if project_id is not None else "Unassigned",
            "accuracy_type": accuracy_type,
            "audited": 0.0,
            "errors": 0.0,
            "bills_processed": 0.0
        })
        item["audited"] += number_value(row.get("audited")) or 0
        item["errors"] += number_value(row.get("errors")) or 0
        item["bills_processed"] += number_value(row.get("bills_processed")) or 0

    current_aggregates = {}
    previous_aggregates = {}
    for row in records:
        period_key = row_periods.get(id(row))
        accuracy_type = normalize_accuracy_type(row.get("accuracy_type"))
        if not period_key or accuracy_type not in ACCURACY_TYPES:
            continue
        project_id = row.get("project_id")
        project_id = int(project_id) if project_id is not None else None
        group_key = (normalize_text(row.get("username")), project_id, accuracy_type)
        if period_key == selected_period:
            aggregate_record(current_aggregates, group_key, row)
        elif period_key == previous_period:
            aggregate_record(previous_aggregates, group_key, row)

    summary = {}
    below_90 = {accuracy_type: [] for accuracy_type in ACCURACY_TYPES}
    for accuracy_type in ACCURACY_TYPES:
        matching = [item for item in current_aggregates.values() if item["accuracy_type"] == accuracy_type]
        audited = sum(item["audited"] for item in matching)
        errors = sum(item["errors"] for item in matching)
        bills = sum(item["bills_processed"] for item in matching)
        summary[accuracy_type] = {
            "audited": audited,
            "errors": errors,
            "bills_processed": bills,
            "accuracy": calculate_weighted_accuracy(audited, errors)
        }

    improvements = []
    for key, current in current_aggregates.items():
        current_accuracy = calculate_weighted_accuracy(current["audited"], current["errors"])
        current["accuracy"] = current_accuracy

        previous = previous_aggregates.get(key)
        previous_accuracy = (
            calculate_weighted_accuracy(previous["audited"], previous["errors"])
            if previous else None
        )
        if previous_accuracy is None:
            change = None
            status = "New Record"
            note = "No previous period record is available for comparison."
        else:
            change = current_accuracy - previous_accuracy if current_accuracy is not None else None
            if change is None:
                status, note = "No Data", "No audited records for this period."
            elif change > 0:
                status = "Improved"
                note = f"Improved by {change:.2f} percentage points" + (", but still below 90%." if current_accuracy < 90 else ".")
            elif change < 0:
                status = "Declined"
                note = f"Declined by {abs(change):.2f} percentage points."
            else:
                status, note = "No Change", "Accuracy remained the same."

        improvements.append({
            "biller_name": current["biller_name"],
            "username": current["username"],
            "project_id": current["project_id"],
            "project_name": current["project_name"],
            "accuracy_type": current["accuracy_type"],
            "previous_accuracy": previous_accuracy,
            "current_accuracy": current_accuracy,
            "change": change,
            "status": status,
            "note": note
        })

    # -----------------------------------------------------
    # INDEPENDENT BELOW 90% DATE RANGE FILTER
    # -----------------------------------------------------

    below_date_aggregates = {}

    for row in records:

        accuracy_type = normalize_accuracy_type(
            row.get("accuracy_type")
        )

        if accuracy_type not in ACCURACY_TYPES:
            continue

        record_day = parse_accuracy_date(
            row.get("record_date"),
            row.get("month")
        )

        if not record_day:
            continue

        if below_from and record_day < below_from:
            continue

        if below_to and record_day > below_to:
            continue

        project_id = row.get("project_id")
        project_id = int(project_id) if project_id is not None else None

        group_key = (
            normalize_text(row.get("username")),
            project_id,
            accuracy_type
        )

        date_key = record_day.isoformat()

        aggregate_key = (group_key, date_key)

        aggregate_record(
            below_date_aggregates,
            aggregate_key,
            row
        )

    # Group daily aggregates by biller, project, and accuracy type.
    below_groups = {}

    for (group_key, date_key), item in below_date_aggregates.items():

        below_groups.setdefault(group_key, []).append({
            **item,
            "record_date": date_key,
            "accuracy": calculate_weighted_accuracy(
                item["audited"],
                item["errors"]
            )
        })

    # Compare first available date against latest available date.
    for group_key, daily_items in below_groups.items():

        daily_items.sort(
            key=lambda item: item["record_date"]
        )

        first_record = daily_items[0]
        latest_record = daily_items[-1]

        first_accuracy = first_record["accuracy"]
        latest_accuracy = latest_record["accuracy"]

        if latest_accuracy is None or latest_accuracy >= 90:
            continue

        if first_accuracy is None:
            change = None
            status = "New Record"

        else:
            change = latest_accuracy - first_accuracy

            if change > 0:
                status = "Improved"
            elif change < 0:
                status = "Declined"
            else:
                status = "No Change"

        below_90[latest_record["accuracy_type"]].append({
            **latest_record,
            "initial_date": first_record["record_date"],
            "latest_date": latest_record["record_date"],
            "initial_accuracy": first_accuracy,
            "latest_accuracy": latest_accuracy,
            "change": change,
            "status": status,
            "accuracy_display": format_accuracy(latest_accuracy)
        })

    # -----------------------------------------------------
    # FILTER BELOW 90% BY CATEGORY
    # -----------------------------------------------------

    if below_category == "internal":

        for accuracy_type in below_90:
            if accuracy_type != "INTERNAL":
                below_90[accuracy_type] = []

    elif below_category == "external":

        for accuracy_type in below_90:
            if not accuracy_type.startswith("EXTERNAL"):
                below_90[accuracy_type] = []

    for accuracy_type in ACCURACY_TYPES:
        below_90[accuracy_type].sort(
            key=lambda item: item["accuracy"]
        )

    improvements.sort(key=lambda item: (
        item["project_name"], item["biller_name"], item["accuracy_type"]
    ))

    if allowed_set is None:
        visible_projects = [
            {"id": int(row["id"]), "project_name": row["project_name"]}
            for row in project_rows if bool(con_value(row, "active", 1))
        ]
    else:
        visible_projects = [
            {"id": int(row["id"]), "project_name": row["project_name"]}
            for row in project_rows
            if int(row["id"]) in allowed_set and bool(con_value(row, "active", 1))
        ]

    return {
        "below_from_date": (
            below_from.isoformat()
            if below_from else ""
        ),
        "below_to_date": (
            below_to.isoformat()
            if below_to else ""
        ),
        "below_category": below_category,
        "below_90": below_90,
        "format_accuracy": format_accuracy,
        "improvements": improvements,
        "period_type": period_type,
        "periods": periods,
        "projects": visible_projects,
        "selected_period": selected_period,
        "selected_project": str(selected_project),
        "summary": summary
    }


# =========================================================
# SAFE ROW VALUE HELPER
# =========================================================

def con_value(row, key, default=None):

    try:

        return row[key]

    except (KeyError, IndexError, TypeError):

        return default


# =========================================================
# ADMIN ACCURACY PAGE
# =========================================================

@app.route("/admin/accuracy")
@login_required
@admin_required
def admin_accuracy():

    period_type = normalize_text(
        request.args.get("period")
    ).lower() or "monthly"

    selected_project = normalize_text(
        request.args.get("project_id")
    ) or "all"

    selected_period = normalize_text(
        request.args.get("date")
    )

    below_from_date = normalize_text(
        request.args.get("below_from")
    )

    below_to_date = normalize_text(
        request.args.get("below_to")
    )

    context = build_accuracy_context(
        allowed_project_ids=None,
        selected_project=selected_project,
        period_type=period_type,
        selected_period=selected_period,
        below_from_date=below_from_date,
        below_to_date=below_to_date
    )

    context["accuracy_endpoint"] = "admin_accuracy"

    return render_template(
        "accuracy.html",
        **context
    )


# =========================================================
# QC ACCURACY PAGE
# QC can only view accuracy for assigned projects.
# =========================================================

@app.route("/qc/accuracy")
@login_required
@qc_required
def qc_accuracy():

    allowed_ids = get_user_project_ids()

    period_type = normalize_text(
        request.args.get("period")
    ).lower() or "monthly"

    selected_project = normalize_text(
        request.args.get("project_id")
    ) or "all"

    selected_period = normalize_text(
        request.args.get("date")
    )

    below_from_date = normalize_text(
        request.args.get("below_from")
    )

    below_to_date = normalize_text(
        request.args.get("below_to")
    )

    context = build_accuracy_context(
        allowed_project_ids=allowed_ids,
        selected_project=selected_project,
        period_type=period_type,
        selected_period=selected_period,
        below_from_date=below_from_date,
        below_to_date=below_to_date
    )

    context["accuracy_endpoint"] = "qc_accuracy"
    
    return render_template(
        "accuracy.html",
        **context
    )


# =========================================================
# BILLER: GET OWN ASSIGNED PROJECT IDS
# =========================================================

def get_biller_project_ids():

    user_id = session.get("user_id")

    if not user_id:
        return []

    return get_user_project_ids(user_id)


# =========================================================
# BILLER: ACCURACY SUMMARY FOR SELECTED MONTH
# =========================================================

def get_biller_accuracy_summary(username, project_ids, selected_month=None):

    if not project_ids:
        return {
            "months": [],
            "selected_month": "",
            "internal_summary": {
                "accuracy": None,
                "audited": 0,
                "errors": 0
            },
            "external_wc_summary": {
                "accuracy": None,
                "audited": 0,
                "errors": 0
            },
            "external_woc_summary": {
                "accuracy": None,
                "audited": 0,
                "errors": 0
            }
        }

    placeholders = ", ".join(
        ["?"] * len(project_ids)
    )

    with db() as con:

        rows = con.execute(
            f"""
            SELECT
                record_date,
                month,
                accuracy_type,
                audited,
                errors,
                project_id
            FROM accuracy_records
            WHERE LOWER(TRIM(username)) = LOWER(TRIM(?))
            AND project_id IN ({placeholders})
            ORDER BY record_date
            """,
            (username, *project_ids)
        ).fetchall()

    months = set()

    for row in rows:

        record_month = month_from_date(
            row["record_date"]
        )

        if not record_month:
            record_month = normalize_text(
                row["month"]
            )

        if record_month:
            months.add(record_month)

    months = sorted(months)

    if selected_month not in months:

        selected_month = (
            months[-1]
            if months
            else ""
        )

    totals = {
        "INTERNAL": {
            "audited": 0.0,
            "errors": 0.0
        },
        "EXTERNAL_WC": {
            "audited": 0.0,
            "errors": 0.0
        },
        "EXTERNAL_WOC": {
            "audited": 0.0,
            "errors": 0.0
        }
    }

    for row in rows:

        record_month = month_from_date(
            row["record_date"]
        )

        if not record_month:
            record_month = normalize_text(
                row["month"]
            )

        if record_month != selected_month:
            continue

        accuracy_type = normalize_accuracy_type(
            row["accuracy_type"]
        )

        if accuracy_type not in totals:
            continue

        totals[accuracy_type]["audited"] += (
            number_value(row["audited"]) or 0
        )

        totals[accuracy_type]["errors"] += (
            number_value(row["errors"]) or 0
        )

    summaries = {}

    for accuracy_type, values in totals.items():

        summaries[accuracy_type] = {
            "audited": values["audited"],
            "errors": values["errors"],
            "accuracy": calculate_weighted_accuracy(
                values["audited"],
                values["errors"]
            )
        }

    return {
        "months": months,
        "selected_month": selected_month,
        "internal_summary": summaries["INTERNAL"],
        "external_wc_summary": summaries["EXTERNAL_WC"],
        "external_woc_summary": summaries["EXTERNAL_WOC"]
    }


# =========================================================
# BILLER DASHBOARD
# =========================================================

@app.route("/biller")
@login_required
@biller_required
def biller_dashboard():

    username = session.get("username", "")

    project_ids = get_biller_project_ids()

    if not project_ids:

        flash(
            "Your account has no assigned active project. Please contact Admin.",
            "warning"
        )

        response_counts = {
            "total": 0,
            "verified": 0,
            "disputed": 0,
            "pending": 0
        }

        accuracy_context = get_biller_accuracy_summary(
            username,
            [],
            request.args.get("month")
        )

        return render_template(
            "biller.html",
            response_counts=response_counts,
            **accuracy_context,
            format_accuracy=format_accuracy
        )

    placeholders = ", ".join(
        ["?"] * len(project_ids)
    )

    with db() as con:

        rows = con.execute(
            f"""
            SELECT
                COUNT(*) AS total,
                SUM(
                    CASE
                        WHEN UPPER(TRIM(COALESCE(verified_error, '')))
                            = 'Y'
                        THEN 1 ELSE 0
                    END
                ) AS verified,
                SUM(
                    CASE
                        WHEN UPPER(TRIM(COALESCE(verified_error, '')))
                            = 'N'
                        THEN 1 ELSE 0
                    END
                ) AS disputed,
                SUM(
                    CASE
                        WHEN verified_error IS NULL
                            OR TRIM(verified_error) = ''
                        THEN 1 ELSE 0
                    END
                ) AS pending
            FROM error_records
            WHERE LOWER(TRIM(username)) = LOWER(TRIM(?))
            AND project_id IN ({placeholders})
            """,
            (username, *project_ids)
        ).fetchone()

    response_counts = {
        "total": int(rows["total"] or 0),
        "verified": int(rows["verified"] or 0),
        "disputed": int(rows["disputed"] or 0),
        "pending": int(rows["pending"] or 0)
    }

    accuracy_context = get_biller_accuracy_summary(
        username,
        project_ids,
        request.args.get("month")
    )

    return render_template(
        "biller.html",
        response_counts=response_counts,
        **accuracy_context,
        format_accuracy=format_accuracy
    )


# =========================================================
# BILLER INTERNAL ACCURACY
# =========================================================

@app.route("/biller/accuracy")
@login_required
@biller_required
def biller_accuracy():
    return redirect(url_for("biller_internal_accuracy"))


@app.route("/biller/accuracy/internal")
@login_required
@biller_required
def biller_internal_accuracy():

    return redirect(
        url_for(
            "biller_dashboard",
            accuracy_type="INTERNAL",
            month=request.args.get("month", "")
        )
    )


# =========================================================
# BILLER EXTERNAL ACCURACY
# =========================================================

@app.route("/biller/accuracy/external")
@login_required
@biller_required
def biller_external_accuracy():

    return redirect(
        url_for(
            "biller_dashboard",
            accuracy_type="EXTERNAL",
            month=request.args.get("month", "")
        )
    )

# =========================================================
# BILLER ERROR RECORDS
# Own records and assigned projects only.
# =========================================================

@app.route("/biller/errors")
@login_required
@biller_required
def biller_errors():

    username = session.get("username", "")

    project_ids = get_biller_project_ids()

    if not project_ids:

        errors = []
        dates = []
        selected_date = ""

    else:

        placeholders = ", ".join(
            ["?"] * len(project_ids)
        )

        with db() as con:

            date_rows = con.execute(
                f"""
                SELECT DISTINCT billed_date
                FROM error_records
                WHERE LOWER(TRIM(username)) = LOWER(TRIM(?))
                AND project_id IN ({placeholders})
                AND billed_date IS NOT NULL
                AND TRIM(billed_date) <> ''
                ORDER BY billed_date DESC
                """,
                (username, *project_ids)
            ).fetchall()

            dates = [
                row["billed_date"]
                for row in date_rows
            ]

            selected_date = normalize_text(
                request.args.get("date")
            )

            if selected_date and selected_date not in dates:
                selected_date = ""

            # Automatically select the latest available date.
            if not selected_date:
                selected_date = dates[0] if dates else ""

            if selected_date:

                rows = con.execute(
                    f"""
                    SELECT
                        e.record_key,
                        e.billed_date,
                        e.pro,
                        e.project_id,
                        p.project_name,
                        e.week,
                        e.source,
                        e.username,
                        e.auditor,
                        e.audit_type,
                        e.biller_name,
                        e.zb_field,
                        e.error_details,
                        e.verified_error,
                        e.biller_comment,
                        e.qc_explanation,
                        e.qc_name,
                        e.updated_at
                    FROM error_records e
                    LEFT JOIN projects p
                        ON p.id = e.project_id
                    WHERE LOWER(TRIM(e.username)) = LOWER(TRIM(?))
                    AND e.project_id IN ({placeholders})
                    AND e.billed_date = ?
                    ORDER BY e.source_row DESC
                    """,
                    (
                        username,
                        *project_ids,
                        selected_date
                    )
                ).fetchall()

            else:

                rows = []

        errors = [
            dict(row)
            for row in rows
        ]

    return render_template(
    "biller_errors.html",
    errors=errors,
    dates=dates,
    selected_date=selected_date
    )

# =========================================================
# QC DASHBOARD
# Redirect to QC Error Records.
# =========================================================

@app.route("/qc")
@login_required
@qc_required
def qc_dashboard():

    return redirect(
        url_for("qc_errors")
    )


# =========================================================
# QC ERROR RECORDS
# Only VERIFIED ERROR = N and assigned projects.
# =========================================================

@app.route("/qc/errors")
@login_required
@qc_required
def qc_errors():

    project_ids = get_user_project_ids()

    if not project_ids:

        flash(
            "Your QC account has no assigned active project. Please contact Admin.",
            "warning"
        )

        return render_template(
            "errors.html",
            errors=[],
            dates=[],
            selected_date="",
            projects=[]
        )

    placeholders = ", ".join(
        ["?"] * len(project_ids)
    )

    with db() as con:

        date_rows = con.execute(
            f"""
            SELECT DISTINCT billed_date
            FROM error_records
            WHERE project_id IN ({placeholders})
            AND UPPER(TRIM(COALESCE(verified_error, '')))
                = 'N'
            AND billed_date IS NOT NULL
            AND TRIM(billed_date) <> ''
            ORDER BY billed_date DESC
            """,
            tuple(project_ids)
        ).fetchall()

        dates = [
            row["billed_date"]
            for row in date_rows
        ]

        selected_date = normalize_text(
            request.args.get("date")
        )

        if selected_date and selected_date not in dates:
            selected_date = ""

        # Automatically select the latest available date
        # when QC first opens the page.
        if not selected_date:
            selected_date = dates[0] if dates else ""

        base_sql = f"""
            SELECT
                e.record_key,
                e.billed_date,
                e.pro,
                e.project_id,
                p.project_name,
                e.week,
                e.source,
                e.username,
                e.auditor,
                e.audit_type,
                e.biller_name,
                e.zb_field,
                e.error_details,
                e.verified_error,
                e.biller_comment,
                e.qc_explanation,
                e.qc_name,
                e.updated_at
            FROM error_records e
            LEFT JOIN projects p
                ON p.id = e.project_id
            WHERE e.project_id IN ({placeholders})
            AND UPPER(TRIM(COALESCE(e.verified_error, '')))
                = 'N'
        """

        if selected_date:

            rows = con.execute(
                base_sql + """
                AND e.billed_date = ?
                ORDER BY e.source_row DESC
                """,
                (
                    *project_ids,
                    selected_date
                )
            ).fetchall()

        else:

            rows = []

        # Supply only the QC account's assigned active projects to the
        # Workbook Import selector. This is template context only; the
        # existing server-side upload authorization remains unchanged.
        project_rows = con.execute(
            f"""
            SELECT id, project_name
            FROM projects
            WHERE active = 1
              AND id IN ({placeholders})
            ORDER BY project_name
            """,
            tuple(project_ids)
        ).fetchall()

    projects = [dict(row) for row in project_rows]

    errors = [
        dict(row)
        for row in rows
    ]

    return render_template(
        "errors.html",
        errors=errors,
        dates=dates,
        selected_date=selected_date,
        projects=projects
    )

# =========================================================
# SAVE BILLER / QC RESPONSE
# =========================================================

@app.route(
    "/response/<record_key>",
    methods=["POST"]
)
@login_required
def save_response(record_key):

    validate_csrf()

    role = normalize_text(
        session.get("role")
    ).lower()

    username = normalize_text(
        session.get("username")
    )

    user_id = session.get("user_id")

    if role not in ("biller", "qc"):

        abort(403)

    try:

        with db() as con:

            record = con.execute(
                """
                SELECT
                    record_key,
                    username,
                    project_id,
                    verified_error,
                    biller_comment,
                    qc_explanation,
                    qc_name,
                    updated_at,
                    updated_by
                FROM error_records
                WHERE record_key = ?
                """,
                (record_key,)
            ).fetchone()

            if not record:
                abort(404)

            project_id = record["project_id"]

            if project_id is None:
                abort(403)

            project_id = int(project_id)
            response_updated_at = datetime.now().isoformat(
                timespec="seconds"
            )

            response_updated_by = username

            # -------------------------------------------------
            # BILLER PERMISSIONS
            # -------------------------------------------------

            if role == "biller":

                allowed_ids = get_user_project_ids(
                    user_id
                )

                if project_id not in allowed_ids:
                    abort(403)

                if normalize_text(
                    record["username"]
                ).lower() != username.lower():

                    abort(403)

                verified_error = normalize_text(
                    request.form.get("verified_error")
                ).upper()

                biller_comment = (
                    request.form.get("biller_comment")
                    or ""
                ).strip()

                if verified_error not in ("Y", "N"):

                    flash(
                        "Please select Y or N for Verified Error.",
                        "warning"
                    )

                    return redirect(
                        url_for(
                            "biller_errors"
                        )
                    )

                cursor = con.execute(
                    """
                    UPDATE error_records
                    SET
                        verified_error = ?,
                        biller_comment = ?,
                        updated_at = ?,
                        updated_by = ?
                    WHERE record_key = ?
                    AND project_id = ?
                    AND LOWER(TRIM(COALESCE(username, ''))) = ?
                    """,
                    (
                        verified_error,
                        biller_comment,
                        response_updated_at,
                        response_updated_by,
                        record_key,
                        project_id,
                        username.lower()
                    )
                )

                if cursor.rowcount != 1:

                    flash(
                        "Unable to update this record. Please refresh the page.",
                        "warning"
                    )

                    return redirect(
                        url_for("biller_errors")
                    )

                latest_record = con.execute(
                    """
                    SELECT
                        qc_explanation,
                        qc_name
                    FROM error_records
                    WHERE record_key = ?
                    AND project_id = ?
                    """,
                    (
                        record_key,
                        project_id
                    )
                ).fetchone()

                if not latest_record:
                    abort(404)

                con.execute(
                    """
                    INSERT INTO biller_responses (
                        record_key,
                        verified_error,
                        biller_comment,
                        qc_explanation,
                        qc_name,
                        updated_by,
                        updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT (record_key)
                    DO UPDATE SET
                        verified_error = excluded.verified_error,
                        biller_comment = excluded.biller_comment,
                        updated_by = excluded.updated_by,
                        updated_at = excluded.updated_at
                    """,
                    (
                        record_key,
                        verified_error,
                        biller_comment,
                        latest_record["qc_explanation"],
                        latest_record["qc_name"],
                        response_updated_by,
                        response_updated_at
                    )
                )

                action = "BILLER RESPONSE"

                details = (
                    f"{username} updated record "
                    f"{record_key}; Verified Error={verified_error}"
                )

            # -------------------------------------------------
            # QC PERMISSIONS
            # -------------------------------------------------

            elif role == "qc":

                allowed_ids = get_user_project_ids(
                    user_id
                )

                if project_id not in allowed_ids:
                    abort(403)

                if normalize_text(
                    record["verified_error"]
                ).upper() != "N":

                    abort(403)

                qc_explanation = (
                    request.form.get("qc_explanation")
                    or ""
                ).strip()

                if not qc_explanation:

                    flash(
                        "QC Explanation is required.",
                        "warning"
                    )

                    return redirect(
                        url_for(
                            "qc_errors"
                        )
                    )

                # QC name is always taken from the
                # authenticated account, never from POST.
                qc_name = normalize_text(
                    session.get("display_name")
                ) or username

                cursor = con.execute(
                    """
                    UPDATE error_records
                    SET
                        qc_explanation = ?,
                        qc_name = ?,
                        updated_at = ?,
                        updated_by = ?
                    WHERE record_key = ?
                    AND project_id = ?
                    AND UPPER(TRIM(COALESCE(verified_error, '')))
                        = 'N'
                    """,
                    (
                        qc_explanation,
                        qc_name,
                        response_updated_at,
                        response_updated_by,
                        record_key,
                        project_id
                    )
                )

                if cursor.rowcount != 1:

                    flash(
                        "This record has already been updated. Please refresh the page.",
                        "warning"
                    )

                    return redirect(
                        url_for("qc_errors")
                    )

                latest_record = con.execute(
                    """
                    SELECT
                        verified_error,
                        biller_comment
                    FROM error_records
                    WHERE record_key = ?
                    AND project_id = ?
                    """,
                    (
                        record_key,
                        project_id
                    )
                ).fetchone()

                if not latest_record:
                    abort(404)


                con.execute(
                    """
                    INSERT INTO biller_responses (
                        record_key,
                        verified_error,
                        biller_comment,
                        qc_explanation,
                        qc_name,
                        updated_by,
                        updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT (record_key)
                    DO UPDATE SET
                        qc_explanation = excluded.qc_explanation,
                        qc_name = excluded.qc_name,
                        updated_by = excluded.updated_by,
                        updated_at = excluded.updated_at
                    """,
                    (
                        record_key,
                        latest_record["verified_error"],
                        latest_record["biller_comment"],
                        qc_explanation,
                        qc_name,
                        response_updated_by,
                        response_updated_at
                    )
                )
                
                action = "QC RESPONSE"

                details = (
                    f"{qc_name} submitted QC explanation "
                    f"for record {record_key}"
                )

            # -------------------------------------------------
            # ACTIVITY LOG
            # -------------------------------------------------

            con.execute(
                """
                INSERT INTO activity_logs (
                    username,
                    action,
                    details,
                    created_at
                )
                VALUES (?, ?, ?, ?)
                """,
                (
                    username,
                    action,
                    details,
                    datetime.now().isoformat(
                        timespec="seconds"
                    )
                )
            )

        flash(
            "Response saved successfully.",
            "success"
        )

    except HTTPException:
        raise

    except Exception:

        app.logger.exception(
            "Unable to save response"
        )

        flash(
            "Unable to save your response. Please try again.",
            "danger"
        )

    if role == "qc":

        return redirect(
            url_for("qc_errors")
        )

    return redirect(
        url_for("biller_errors")
    )


# =========================================================
# EXPORT BILLER / QC RESPONSES
# ADMIN ONLY
# =========================================================

@app.route("/admin/export-responses")
@login_required
@admin_required
def export_responses():

    output = io.StringIO()

    writer = csv.writer(output)

    writer.writerow([
        "PROJECT",
        "BILLED DATE",
        "WEEK",
        "BILLER NAME",
        "USERNAME",
        "PRO",
        "VERIFIED ERROR",
        "BILLER COMMENT",
        "QC NAME",
        "QC EXPLANATION",
        "UPDATED AT"
    ])

    with db() as con:

        rows = con.execute(
            """
            SELECT
                p.project_name,
                e.billed_date,
                e.week,
                e.biller_name,
                e.username,
                e.pro,
                e.verified_error,
                e.biller_comment,
                e.qc_name,
                e.qc_explanation,
                e.updated_at
            FROM error_records e
            LEFT JOIN projects p
                ON p.id = e.project_id
            ORDER BY
                e.billed_date DESC,
                e.source_row DESC
            """
        ).fetchall()

    for row in rows:

        writer.writerow([
            row["project_name"] or "Unassigned",
            row["billed_date"] or "",
            row["week"] or "",
            row["biller_name"] or "",
            row["username"] or "",
            row["pro"] or "",
            row["verified_error"] or "",
            row["biller_comment"] or "",
            row["qc_name"] or "",
            row["qc_explanation"] or "",
            row["updated_at"] or ""
        ])

    response = Response(
        output.getvalue(),
        mimetype="text/csv"
    )

    response.headers[
        "Content-Disposition"
    ] = (
        'attachment; filename="biller_qc_responses.csv"'
    )

    return response


# =========================================================
# ADMIN ACTIVITY LOGS
# =========================================================

@app.route("/admin/import-history")
@login_required
@admin_required
def admin_import_history():
    with db() as con:
        rows = con.execute(
            """SELECT id, filename, uploaded_by, records_imported, imported_at
               FROM import_history ORDER BY id DESC LIMIT 500"""
        ).fetchall()
    return render_template("import_history.html", history=[dict(row) for row in rows], scope="All system imports")


@app.route("/qc/import-history")
@login_required
@qc_required
def qc_import_history():
    username = normalize_text(session.get("username"))
    with db() as con:
        rows = con.execute(
            """SELECT id, filename, uploaded_by, records_imported, imported_at
               FROM import_history
               WHERE LOWER(TRIM(COALESCE(uploaded_by, ''))) = LOWER(?)
               ORDER BY id DESC LIMIT 500""",
            (username,)
        ).fetchall()
    return render_template("import_history.html", history=[dict(row) for row in rows], scope="Your QC imports")


@app.route("/admin/logs")
@login_required
@admin_required
def admin_logs():

    with db() as con:

        rows = con.execute(
            """
            SELECT
                id,
                username,
                action,
                details,
                created_at
            FROM activity_logs
            ORDER BY id DESC
            LIMIT 500
            """
        ).fetchall()

    return render_template(
        "logs.html",
        rows=rows
    )


# =========================================================
# ERROR HANDLERS
# =========================================================

@app.errorhandler(400)
def bad_request(error):

    return render_template(
        "base.html",
        error_message="Bad request."
    ), 400


@app.errorhandler(403)
def forbidden(error):

    return render_template(
        "base.html",
        error_message="You do not have permission to access this page."
    ), 403


@app.errorhandler(404)
def not_found(error):

    return render_template(
        "base.html",
        error_message="The requested page or record was not found."
    ), 404


@app.errorhandler(500)
def internal_server_error(error):

    app.logger.exception(
        "Internal server error"
    )

    return render_template(
        "base.html",
        error_message="An unexpected server error occurred."
    ), 500


# =========================================================
# MANUAL DATABASE INITIALIZATION COMMAND
# =========================================================

@app.cli.command("init-db")
def initialize_database_command():

    init_db()

    print(
        "Database initialization completed successfully."
    )

# =========================================================
# LOCAL DEVELOPMENT
# Render uses the Procfile / Gunicorn command.
# =========================================================

if __name__ == "__main__":

    app.run(
        host="127.0.0.1",
        port=int(
            os.environ.get("PORT", 5001)
        ),
        debug=False
    )

