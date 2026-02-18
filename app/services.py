from __future__ import annotations

import csv
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, Optional

from app.connectors.yampi import YampiClient
from app.database import (
    delete_orders_by_period,
    fetch_monthly_for_export,
    fetch_orders_for_export,
    get_cursor,
    set_cursor,
    upsert_orders,
)


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _pick(dct: Dict[str, Any], *keys: str) -> Optional[Any]:
    for key in keys:
        if key in dct and dct[key] is not None:
            return dct[key]
    return None


def _extract_date(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value[:10]
    if isinstance(value, dict):
        date_value = value.get("date")
        if isinstance(date_value, str):
            return date_value[:10]
    return str(value)[:10]


def _extract_status_name(status: Any) -> str:
    if status is None:
        return ""
    if isinstance(status, str):
        return status
    if isinstance(status, dict):
        status_data = status.get("data")
        if isinstance(status_data, dict):
            name = status_data.get("name")
            if name:
                return str(name)
    return str(status)


def _to_float(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _extract_spreadsheet_field(order: Dict[str, Any], field: str) -> str:
    spreadsheet = order.get("spreadsheet")
    if not isinstance(spreadsheet, dict):
        return ""
    data = spreadsheet.get("data")
    if not isinstance(data, list):
        return ""
    for item in data:
        if not isinstance(item, dict):
            continue
        value = item.get(field)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


def _extract_transaction_field(order: Dict[str, Any], field: str) -> str:
    transactions = order.get("transactions")
    if not isinstance(transactions, dict):
        return ""
    data = transactions.get("data")
    if not isinstance(data, list):
        return ""
    for item in data:
        if not isinstance(item, dict):
            continue
        value = item.get(field)
        if value is None:
            continue
        if isinstance(value, dict):
            text = _extract_date(value).strip()
        else:
            text = str(value).strip()
        if text:
            return text
    return ""


def _extract_payment_date(order: Dict[str, Any]) -> str:
    value = _extract_spreadsheet_field(order, "payment_date")
    if value:
        return value
    return _extract_transaction_field(order, "captured_at")


def _extract_cancelled_date(order: Dict[str, Any]) -> str:
    value = _extract_spreadsheet_field(order, "cancelled_date")
    if value:
        return value
    return _extract_transaction_field(order, "cancelled_at")


def _iter_month_ranges(start_date: str, end_date: str) -> list[tuple[str, str]]:
    start = datetime.strptime(start_date, "%Y-%m-%d").date()
    end = datetime.strptime(end_date, "%Y-%m-%d").date()
    ranges: list[tuple[str, str]] = []

    current = start
    while current <= end:
        first = current.replace(day=1)
        if first.month == 12:
            next_month = first.replace(year=first.year + 1, month=1, day=1)
        else:
            next_month = first.replace(month=first.month + 1, day=1)
        month_end = min(end, next_month - timedelta(days=1))
        ranges.append((current.isoformat(), month_end.isoformat()))
        current = month_end + timedelta(days=1)

    return ranges


def _normalize_orders(client_id: str, raw_orders: Iterable[Dict[str, Any]]) -> tuple[list[tuple], Optional[str]]:
    rows: list[tuple] = []
    max_updated: Optional[str] = None
    extracted_at = _utc_now_iso()

    for order in raw_orders:
        customer = order.get("customer", {}) if isinstance(order.get("customer"), dict) else {}

        order_id = str(_pick(order, "id", "order_id", "number"))
        updated_at = _pick(order, "updated_at", "date_updated")
        created_at = _pick(order, "created_at", "date_created")
        created_date = _extract_date(created_at)
        status_value = order.get("status")
        status_name = _extract_status_name(status_value)
        value_products = _to_float(order.get("value_products"))
        value_shipment = _to_float(order.get("value_shipment"))
        value_discount = _to_float(order.get("value_discount"))
        value_tax = _to_float(order.get("value_tax"))
        payment_date = _extract_payment_date(order)
        cancelled_date = _extract_cancelled_date(order)

        if updated_at and (max_updated is None or str(updated_at) > max_updated):
            max_updated = str(updated_at)

        rows.append(
            (
                client_id,
                order_id,
                str(status_value or ""),
                status_name,
                str(_pick(order, "total", "total_value", "value_total") or ""),
                str(created_at or ""),
                created_date,
                str(updated_at or ""),
                value_products,
                value_shipment,
                value_discount,
                value_tax,
                payment_date,
                cancelled_date,
                str(_pick(customer, "name", "full_name") or ""),
                str(_pick(customer, "email") or ""),
                json.dumps(order, ensure_ascii=False),
                extracted_at,
            )
        )

    return rows, max_updated


def sync_yampi_orders(
    conn,
    client_id: str,
    base_url: str,
    alias: str,
    token: str = "",
    user_token: str = "",
    user_secret_key: str = "",
    page_size: int = 100,
    max_pages: int = 1000,
    start_date: str = "",
    end_date: str = "",
) -> int:
    source = "yampi_orders"
    cursor = get_cursor(conn, client_id, source)
    if start_date or end_date:
        cursor = None
    yampi = YampiClient(
        base_url=base_url,
        token=token,
        user_token=user_token,
        user_secret_key=user_secret_key,
    )

    def _sync_window(window_start: str, window_end: str, window_cursor: Optional[str]) -> tuple[int, Optional[str]]:
        rows_total = 0
        max_seen = window_cursor
        scroll_id: Optional[str] = None
        total_pages = 1
        page = 1

        while page <= min(total_pages, max_pages):
            raw_orders, next_scroll_id, current_total_pages = yampi.fetch_orders(
                alias=alias,
                page=page,
                page_size=page_size,
                scroll_id=scroll_id,
                updated_since=window_cursor,
                start_date=window_start or None,
                end_date=window_end or None,
            )
            total_pages = max(total_pages, current_total_pages)
            if not raw_orders:
                break

            rows, page_cursor = _normalize_orders(client_id, raw_orders)
            if rows:
                upsert_orders(conn, rows)
                rows_total += len(rows)
                conn.commit()

            if page_cursor and (max_seen is None or page_cursor > max_seen):
                max_seen = page_cursor

            if next_scroll_id:
                scroll_id = str(next_scroll_id)
            page += 1

        return rows_total, max_seen

    total_rows = 0
    max_seen_cursor = cursor

    if start_date and end_date:
        # Evita erro de limite 10k da API processando em blocos mensais.
        for month_start, month_end in _iter_month_ranges(start_date, end_date):
            rows, _ = _sync_window(month_start, month_end, None)
            total_rows += rows
    else:
        rows, max_seen_cursor = _sync_window("", "", cursor)
        total_rows += rows

    if not start_date and not end_date:
        set_cursor(conn, client_id, source, max_seen_cursor, _utc_now_iso())
        conn.commit()
    return total_rows


def export_orders_csv(conn, client_id: str, output_path: str) -> int:
    rows = fetch_orders_for_export(conn, client_id)
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    headers = [
        "client_id",
        "company",
        "branch",
        "alias",
        "order_id",
        "status_name",
        "total",
        "created_at",
        "updated_at",
        "customer_name",
        "customer_email",
        "extracted_at",
    ]

    with path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL)
        writer.writerow(headers)
        for row in rows:
            writer.writerow([row[h] for h in headers])

    return len(rows)


def export_monthly_sheet_csv(
    conn, client_id: str, output_path: str, start_date: str = "", end_date: str = ""
) -> int:
    rows = fetch_monthly_for_export(conn, client_id, start_date=start_date, end_date=end_date)
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    headers = [
        "Data",
        "Nome Empresa",
        "Alias",
        "Vendas de Produto",
        "Descontos concedidos",
        "Juros de Venda",
    ]

    with path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL)
        writer.writerow(headers)
        for row in rows:
            writer.writerow(
                [
                    _format_date_br(row["data"]),
                    row["nome_empresa"] or "",
                    row["alias"] or "",
                    _format_money_br(row["vendas_de_produto"] or 0),
                    _format_money_br(row["descontos_concedidos"] or 0),
                    _format_money_br(row["juros_de_venda"] or 0),
                ]
            )

    return len(rows)


def reprocess_orders_for_period(
    conn,
    client_id: str,
    base_url: str,
    alias: str,
    start_date: str,
    end_date: str,
    token: str = "",
    user_token: str = "",
    user_secret_key: str = "",
    page_size: int = 100,
) -> tuple[int, int]:
    deleted = delete_orders_by_period(conn, client_id, start_date, end_date)
    conn.commit()
    synced = sync_yampi_orders(
        conn=conn,
        client_id=client_id,
        base_url=base_url,
        alias=alias,
        token=token,
        user_token=user_token,
        user_secret_key=user_secret_key,
        page_size=page_size,
        start_date=start_date,
        end_date=end_date,
    )
    return deleted, synced


def _format_date_br(yyyy_mm_dd: str) -> str:
    if not yyyy_mm_dd:
        return ""
    try:
        return datetime.strptime(yyyy_mm_dd[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return yyyy_mm_dd


def _format_money_br(value: float) -> str:
    s = f"{float(value):,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"
