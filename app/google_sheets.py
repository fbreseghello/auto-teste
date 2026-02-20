from __future__ import annotations

import csv
import os
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path

SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets"


@dataclass(frozen=True)
class MonthlySheetsSettings:
    spreadsheet_id: str
    credentials_json_path: str
    worksheet_prefix: str = ""
    clear_before_upload: bool = True
    create_worksheet_if_missing: bool = False


def _env_flag(value: str, default: bool = False) -> bool:
    text = value.strip().lower()
    if not text:
        return default
    return text in {"1", "true", "t", "yes", "y", "on"}


def _runtime_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def _resolve_credentials_path(path_value: str) -> Path:
    raw = Path(path_value).expanduser()
    if raw.is_absolute():
        return raw
    return (_runtime_root() / raw).resolve()


def _safe_worksheet_name(name: str) -> str:
    cleaned = name.strip().replace("/", "_").replace("\\", "_")
    if not cleaned:
        return "mensal"
    if len(cleaned) > 100:
        return cleaned[:100]
    return cleaned


def _normalize_worksheet_key(value: str) -> str:
    raw = unicodedata.normalize("NFKD", value or "")
    ascii_text = "".join(ch for ch in raw if not unicodedata.combining(ch))
    return "".join(ch.lower() for ch in ascii_text if ch.isalnum())


def worksheet_name_for_client(alias: str, client_id: str, prefix: str = "") -> str:
    base = alias.strip() or client_id.strip() or "mensal"
    return _safe_worksheet_name(f"{prefix}{base}")


def resolve_monthly_sheets_settings() -> MonthlySheetsSettings:
    spreadsheet_id = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", "").strip()
    credentials_json_path = os.getenv("GOOGLE_SHEETS_CREDENTIALS_JSON", "").strip()
    worksheet_prefix = os.getenv("GOOGLE_SHEETS_WORKSHEET_PREFIX", "").strip()
    clear_before_upload = _env_flag(os.getenv("GOOGLE_SHEETS_CLEAR_BEFORE_UPLOAD", ""), default=True)
    create_worksheet_if_missing = _env_flag(
        os.getenv("GOOGLE_SHEETS_CREATE_WORKSHEET_IF_MISSING", ""),
        default=False,
    )

    if not spreadsheet_id:
        raise ValueError("Defina GOOGLE_SHEETS_SPREADSHEET_ID no .env.")
    if not credentials_json_path:
        raise ValueError("Defina GOOGLE_SHEETS_CREDENTIALS_JSON no .env.")
    resolved_credentials = _resolve_credentials_path(credentials_json_path)
    if not resolved_credentials.exists():
        raise ValueError(
            f"Arquivo de credencial Google nao encontrado em: {resolved_credentials}. "
            "Informe o caminho completo do JSON da Service Account."
        )

    return MonthlySheetsSettings(
        spreadsheet_id=spreadsheet_id,
        credentials_json_path=str(resolved_credentials),
        worksheet_prefix=worksheet_prefix,
        clear_before_upload=clear_before_upload,
        create_worksheet_if_missing=create_worksheet_if_missing,
    )


def _load_csv_values(csv_path: str) -> list[list[str]]:
    file_path = Path(csv_path).expanduser()
    if not file_path.exists():
        raise FileNotFoundError(f"CSV nao encontrado: {csv_path}")

    rows: list[list[str]] = []
    with file_path.open("r", newline="", encoding="utf-8-sig") as csv_file:
        reader = csv.reader(csv_file)
        for row in reader:
            rows.append([str(cell) for cell in row])
    return rows


def _build_sheets_service(credentials_json_path: str):
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "Dependencias do Google Sheets ausentes. Execute: pip install -r requirements.txt"
        ) from exc

    creds = service_account.Credentials.from_service_account_file(
        credentials_json_path,
        scopes=[SHEETS_SCOPE],
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _find_existing_worksheet_name(service, spreadsheet_id: str, worksheet_name: str) -> str:
    metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = metadata.get("sheets", [])
    target_key = _normalize_worksheet_key(worksheet_name)

    for sheet in sheets:
        properties = sheet.get("properties", {})
        title = str(properties.get("title") or "")
        if title == worksheet_name:
            return title
        if target_key and _normalize_worksheet_key(title) == target_key:
            return title
    return ""


def _create_worksheet(service, spreadsheet_id: str, worksheet_name: str) -> None:

    body = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": worksheet_name,
                    }
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def upload_csv_to_google_sheet(
    csv_path: str,
    spreadsheet_id: str,
    worksheet_name: str,
    credentials_json_path: str,
    clear_before_upload: bool = True,
    create_worksheet_if_missing: bool = False,
) -> int:
    values = _load_csv_values(csv_path)
    service = _build_sheets_service(credentials_json_path)
    safe_name = _safe_worksheet_name(worksheet_name)
    existing_name = _find_existing_worksheet_name(service, spreadsheet_id, safe_name)
    if not existing_name:
        if not create_worksheet_if_missing:
            raise ValueError(
                f"Aba '{safe_name}' nao encontrada na planilha. "
                "Crie a aba antes de enviar ou configure sheets_worksheet no clients.json."
            )
        _create_worksheet(service, spreadsheet_id, safe_name)
        existing_name = safe_name

    sheet_range = f"'{existing_name}'!A1"
    clear_range = f"'{existing_name}'!A:Z"
    if clear_before_upload:
        service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=clear_range,
            body={},
        ).execute()

    if values:
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=sheet_range,
            valueInputOption="USER_ENTERED",
            body={"values": values},
        ).execute()

    return max(0, len(values) - 1)
