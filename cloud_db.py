# -*- coding: utf-8 -*-
"""
Cloud Database & Storage Connector for Control de Facturacion
Supports:
1. PostgreSQL (Supabase / Neon) via DATABASE_URL
2. SQLite (local facturacion.db) as seamless fallback
3. Supabase Storage for PDF/Excel uploads via SUPABASE_URL & SUPABASE_KEY
4. Local file archive fallback (/uploads & uploaded_files_archive)
"""
import os
import sqlite3
import datetime
import re
from typing import List, Dict, Any, Optional, Tuple, Union

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
SUPABASE_URL = os.environ.get("SUPABASE_URL", "").strip()
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "").strip()
SUPABASE_BUCKET = os.environ.get("SUPABASE_BUCKET", "facturas-docs").strip()

# Check for psycopg2 availability
HAS_PSYCOPG2 = False
try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    HAS_PSYCOPG2 = True
except ImportError:
    pass

# Determine if cloud PostgreSQL is active
USE_POSTGRES = bool(DATABASE_URL and HAS_PSYCOPG2 and (
    DATABASE_URL.startswith("postgresql://") or DATABASE_URL.startswith("postgres://")
))

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "facturacion.db")

class CloudRow:
    """Provides dual access (by key string AND by numeric index) identical to sqlite3.Row."""
    def __init__(self, d: dict, t: tuple):
        self._d = d
        self._t = t

    def __getitem__(self, key: Union[int, str]):
        if isinstance(key, int):
            return self._t[key]
        return self._d.get(key)

    def get(self, key: str, default: Any = None):
        return self._d.get(key, default)

    def keys(self):
        return self._d.keys()

    def values(self):
        return self._d.values()

    def items(self):
        return self._d.items()

    def __iter__(self):
        # Yields tuple elements for positional unpacking: a, b, c = row
        return iter(self._t)

    def __contains__(self, key: str):
        return key in self._d

    def __len__(self):
        return len(self._t)

    def __repr__(self):
        return f"<CloudRow {self._d}>"


def translate_sql(sql: str, is_pg: bool, has_params: bool = False) -> str:
    """Translates SQLite query syntax to PostgreSQL if running on cloud PG."""
    if not is_pg:
        return sql

    translated = sql

    # If query has parameters, escape any literal % (e.g. LIKE '%TEST%') so psycopg2 doesn't try to interpolate it
    if has_params:
        # Replace % with %% only where it's not part of an existing %s
        translated = re.sub(r'%(?!s)', '%%', translated)

    # Replace SQLite parameter placeholder ? with PostgreSQL %s
    translated = translated.replace("?", "%s")

    # SQLite AUTOINCREMENT to SERIAL
    translated = re.sub(r'INTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT', 'SERIAL PRIMARY KEY', translated, flags=re.IGNORECASE)

    # SQLite BLOB to BYTEA
    translated = re.sub(r'\bBLOB\b', 'BYTEA', translated)

    # SQLite BOOLEAN DEFAULT 0 -> BOOLEAN DEFAULT FALSE
    translated = re.sub(r'BOOLEAN\s+DEFAULT\s+0', 'BOOLEAN DEFAULT FALSE', translated, flags=re.IGNORECASE)
    translated = re.sub(r'BOOLEAN\s+DEFAULT\s+1', 'BOOLEAN DEFAULT TRUE', translated, flags=re.IGNORECASE)

    # Replace SQLite INSERT OR IGNORE
    if "INSERT OR IGNORE INTO cargos" in translated:
        translated = translated.replace("INSERT OR IGNORE INTO cargos (title) VALUES (%s)",
                                        "INSERT INTO cargos (title) VALUES (%s) ON CONFLICT (title) DO NOTHING")
    elif "INSERT OR IGNORE INTO matrix_cells" in translated:
        translated = translated.replace("INSERT OR IGNORE INTO matrix_cells", "INSERT INTO matrix_cells")
        if "ON CONFLICT" not in translated:
            translated += " ON CONFLICT (client_id, period_col) DO NOTHING"
    elif "INSERT OR IGNORE" in translated:
        translated = translated.replace("INSERT OR IGNORE INTO", "INSERT INTO")

    # Replace SQLite INSERT OR REPLACE INTO
    if "INSERT OR REPLACE INTO uploaded_files_archive" in translated:
        translated = translated.replace(
            "INSERT OR REPLACE INTO uploaded_files_archive",
            "INSERT INTO uploaded_files_archive"
        )
        if "ON CONFLICT" not in translated:
            translated += " ON CONFLICT (file_url) DO UPDATE SET filename = EXCLUDED.filename, content_type = EXCLUDED.content_type, file_bytes = EXCLUDED.file_bytes, record_id = EXCLUDED.record_id, doc_type = EXCLUDED.doc_type, created_at = EXCLUDED.created_at"
    elif "INSERT OR REPLACE INTO billing_records" in translated:
        translated = translated.replace(
            "INSERT OR REPLACE INTO billing_records",
            "INSERT INTO billing_records"
        )
        if "ON CONFLICT" not in translated:
            translated += " ON CONFLICT (record_id) DO NOTHING"

    return translated


class CloudCursor:
    def __init__(self, raw_cursor, is_pg: bool):
        self._c = raw_cursor
        self.is_pg = is_pg

    def execute(self, sql: str, params: tuple = ()):
        has_params = bool(params)
        final_sql = translate_sql(sql, self.is_pg, has_params=has_params)

        if self.is_pg:
            if has_params:
                # PostgreSQL requires params as tuple or list
                if not isinstance(params, (tuple, list)):
                    params = (params,)
                # Adapt any binary bytes for psycopg2
                clean_params = []
                for p in params:
                    if isinstance(p, (bytes, bytearray)):
                        clean_params.append(psycopg2.Binary(p))
                    else:
                        clean_params.append(p)
                return self._c.execute(final_sql, tuple(clean_params))
            else:
                return self._c.execute(final_sql)
        else:
            return self._c.execute(final_sql, params)

    def executemany(self, sql: str, seq_of_parameters):
        final_sql = translate_sql(sql, self.is_pg, has_params=True)
        if self.is_pg:
            clean_seq = []
            for params in seq_of_parameters:
                clean_params = [psycopg2.Binary(p) if isinstance(p, (bytes, bytearray)) else p for p in params]
                clean_seq.append(tuple(clean_params))
            return self._c.executemany(final_sql, clean_seq)
        else:
            return self._c.executemany(final_sql, seq_of_parameters)

    def _convert_val(self, val):
        if isinstance(val, memoryview):
            return bytes(val)
        return val

    def fetchone(self):
        row = self._c.fetchone()
        if row is None:
            return None
        if self.is_pg:
            col_names = [d[0] for d in self._c.description]
            clean_row = tuple(self._convert_val(v) for v in row)
            d = dict(zip(col_names, clean_row))
            return CloudRow(d, clean_row)
        else:
            return row

    def fetchall(self):
        rows = self._c.fetchall()
        if not rows:
            return []
        if self.is_pg:
            col_names = [d[0] for d in self._c.description]
            res = []
            for r in rows:
                clean_row = tuple(self._convert_val(v) for v in r)
                d = dict(zip(col_names, clean_row))
                res.append(CloudRow(d, clean_row))
            return res
        else:
            return rows

    @property
    def description(self):
        return self._c.description

    def close(self):
        return self._c.close()


class CloudConnection:
    def __init__(self, raw_conn, is_pg: bool):
        self._conn = raw_conn
        self.is_pg = is_pg
        self._row_factory = None

    @property
    def row_factory(self):
        return self._row_factory

    @row_factory.setter
    def row_factory(self, value):
        self._row_factory = value
        if not self.is_pg:
            self._conn.row_factory = value

    def cursor(self):
        return CloudCursor(self._conn.cursor(), self.is_pg)

    def commit(self):
        return self._conn.commit()

    def rollback(self):
        return self._conn.rollback()

    def close(self):
        return self._conn.close()

    def execute(self, sql: str, params: tuple = ()):
        c = self.cursor()
        c.execute(sql, params)
        return c

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is not None:
            self.rollback()
        else:
            self.commit()


def get_db() -> CloudConnection:
    """Returns a unified CloudConnection connected to PostgreSQL or SQLite."""
    if USE_POSTGRES:
        pg_url = DATABASE_URL
        if pg_url.startswith("postgres://"):
            pg_url = "postgresql://" + pg_url[11:]
        if "sslmode" not in pg_url and "?" not in pg_url:
            pg_url += "?sslmode=require"
        conn = psycopg2.connect(pg_url)
        return CloudConnection(conn, True)
    else:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        return CloudConnection(conn, False)


def get_cloud_status() -> Dict[str, Any]:
    """Provides operational cloud health diagnostics."""
    status = {
        "use_postgres": USE_POSTGRES,
        "database_engine": "PostgreSQL (Cloud Supabase/Neon)" if USE_POSTGRES else "SQLite (Local facturacion.db)",
        "cloud_storage_active": bool(SUPABASE_URL and SUPABASE_KEY),
        "cloud_storage_bucket": SUPABASE_BUCKET if (SUPABASE_URL and SUPABASE_KEY) else "Local (/uploads & DB archive)",
        "connected": False,
        "records_count": 0,
        "clients_count": 0
    }
    try:
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM clients")
        status["clients_count"] = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM billing_records")
        status["records_count"] = c.fetchone()[0]
        conn.close()
        status["connected"] = True
    except Exception as e:
        status["error"] = str(e)
    return status


async def upload_file_to_cloud(file_bytes: bytes, filename: str, content_type: str) -> Optional[str]:
    """
    Uploads a file to Supabase Storage bucket if credentials are configured.
    Returns the public URL, or None if cloud storage is not configured.
    """
    if not (SUPABASE_URL and SUPABASE_KEY):
        return None

    try:
        import httpx
        url = f"{SUPABASE_URL.rstrip('/')}/storage/v1/object/{SUPABASE_BUCKET}/{filename}"
        headers = {
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "apikey": SUPABASE_KEY,
            "Content-Type": content_type or "application/octet-stream",
            "x-upsert": "true"
        }
        async with httpx.AsyncClient(timeout=30.0) as client:
            resp = await client.post(url, content=file_bytes, headers=headers)
            if resp.status_code in (200, 201):
                public_url = f"{SUPABASE_URL.rstrip('/')}/storage/v1/object/public/{SUPABASE_BUCKET}/{filename}"
                return public_url
            else:
                print(f"Supabase Storage respondió status {resp.status_code}: {resp.text}")
    except Exception as e:
        print(f"Error subiendo archivo a Supabase Storage: {e}")
    return None
