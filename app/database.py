from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Iterable, Optional


DEFAULT_DB_PATH = "data/local.db"


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS clients (
    id TEXT PRIMARY KEY,
    company TEXT,
    branch TEXT,
    alias TEXT,
    name TEXT NOT NULL,
    platform TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS yampi_orders (
    client_id TEXT NOT NULL,
    order_id TEXT NOT NULL,
    status TEXT,
    status_name TEXT,
    total TEXT,
    created_at TEXT,
    created_date TEXT,
    updated_at TEXT,
    value_products REAL DEFAULT 0,
    value_shipment REAL DEFAULT 0,
    value_discount REAL DEFAULT 0,
    value_tax REAL DEFAULT 0,
    customer_name TEXT,
    customer_email TEXT,
    raw_json TEXT NOT NULL,
    extracted_at TEXT NOT NULL,
    PRIMARY KEY (client_id, order_id)
);

CREATE TABLE IF NOT EXISTS sync_state (
    client_id TEXT NOT NULL,
    source TEXT NOT NULL,
    cursor TEXT,
    updated_at TEXT NOT NULL,
    PRIMARY KEY (client_id, source)
);
"""


def connect(db_path: str = DEFAULT_DB_PATH) -> sqlite3.Connection:
    path = Path(db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(SCHEMA_SQL)
    _ensure_column(conn, "clients", "company", "TEXT")
    _ensure_column(conn, "clients", "branch", "TEXT")
    _ensure_column(conn, "clients", "alias", "TEXT")
    _ensure_column(conn, "yampi_orders", "status_name", "TEXT")
    _ensure_column(conn, "yampi_orders", "created_date", "TEXT")
    _ensure_column(conn, "yampi_orders", "value_products", "REAL DEFAULT 0")
    _ensure_column(conn, "yampi_orders", "value_shipment", "REAL DEFAULT 0")
    _ensure_column(conn, "yampi_orders", "value_discount", "REAL DEFAULT 0")
    _ensure_column(conn, "yampi_orders", "value_tax", "REAL DEFAULT 0")
    conn.commit()


def _ensure_column(conn: sqlite3.Connection, table: str, column: str, col_type: str) -> None:
    columns = {row["name"] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}
    if column not in columns:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {col_type}")


def upsert_client(
    conn: sqlite3.Connection,
    client_id: str,
    company: str,
    branch: str,
    alias: str,
    name: str,
    platform: str,
) -> None:
    conn.execute(
        """
        INSERT INTO clients (id, company, branch, alias, name, platform)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            company = excluded.company,
            branch = excluded.branch,
            alias = excluded.alias,
            name = excluded.name,
            platform = excluded.platform;
        """,
        (client_id, company, branch, alias, name, platform),
    )


def upsert_orders(conn: sqlite3.Connection, rows: Iterable[tuple]) -> None:
    conn.executemany(
        """
        INSERT INTO yampi_orders (
            client_id, order_id, status, status_name, total, created_at, created_date, updated_at,
            value_products, value_shipment, value_discount, value_tax,
            customer_name, customer_email, raw_json, extracted_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(client_id, order_id) DO UPDATE SET
            status = excluded.status,
            status_name = excluded.status_name,
            total = excluded.total,
            created_at = excluded.created_at,
            created_date = excluded.created_date,
            updated_at = excluded.updated_at,
            value_products = excluded.value_products,
            value_shipment = excluded.value_shipment,
            value_discount = excluded.value_discount,
            value_tax = excluded.value_tax,
            customer_name = excluded.customer_name,
            customer_email = excluded.customer_email,
            raw_json = excluded.raw_json,
            extracted_at = excluded.extracted_at;
        """,
        list(rows),
    )


def get_cursor(conn: sqlite3.Connection, client_id: str, source: str) -> Optional[str]:
    row = conn.execute(
        "SELECT cursor FROM sync_state WHERE client_id = ? AND source = ?",
        (client_id, source),
    ).fetchone()
    return row["cursor"] if row else None


def set_cursor(conn: sqlite3.Connection, client_id: str, source: str, cursor: Optional[str], updated_at: str) -> None:
    conn.execute(
        """
        INSERT INTO sync_state (client_id, source, cursor, updated_at)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(client_id, source) DO UPDATE SET
            cursor = excluded.cursor,
            updated_at = excluded.updated_at;
        """,
        (client_id, source, cursor, updated_at),
    )


def fetch_orders_for_export(conn: sqlite3.Connection, client_id: str) -> list[sqlite3.Row]:
    return conn.execute(
        """
        SELECT
            y.client_id, c.company, c.branch, c.alias, order_id, status_name, total, created_at, updated_at,
            customer_name, customer_email, extracted_at
        FROM yampi_orders y
        LEFT JOIN clients c ON c.id = y.client_id
        WHERE y.client_id = ?
        ORDER BY updated_at DESC, order_id DESC
        """,
        (client_id,),
    ).fetchall()


def fetch_monthly_for_export(
    conn: sqlite3.Connection, client_id: str, start_date: str = "", end_date: str = ""
) -> list[sqlite3.Row]:
    where_parts = ["y.client_id = ?"]
    params: list[str] = [client_id]

    if start_date:
        where_parts.append("y.created_date >= ?")
        params.append(start_date)
    if end_date:
        where_parts.append("y.created_date <= ?")
        params.append(end_date)

    where_sql = " AND ".join(where_parts)
    return conn.execute(
        f"""
        SELECT
            substr(y.created_date, 1, 7) || '-01' AS data,
            c.company AS nome_empresa,
            c.alias AS alias,
            ROUND(SUM(COALESCE(y.value_products, 0) + COALESCE(y.value_shipment, 0)), 2) AS vendas_de_produto,
            ROUND(SUM(COALESCE(y.value_discount, 0)), 2) AS descontos_concedidos,
            ROUND(SUM(COALESCE(y.value_tax, 0)), 2) AS juros_de_venda
        FROM yampi_orders y
        LEFT JOIN clients c ON c.id = y.client_id
        WHERE {where_sql}
        GROUP BY substr(y.created_date, 1, 7), c.company, c.alias
        ORDER BY data DESC
        """,
        params,
    ).fetchall()


def delete_orders_by_period(conn: sqlite3.Connection, client_id: str, start_date: str, end_date: str) -> int:
    cur = conn.execute(
        """
        DELETE FROM yampi_orders
        WHERE client_id = ?
          AND created_date >= ?
          AND created_date <= ?
        """,
        (client_id, start_date, end_date),
    )
    return int(cur.rowcount or 0)
