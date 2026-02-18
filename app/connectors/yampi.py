from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

import requests


ACTIVE_STATUS_IDS = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13"]


class YampiClient:
    def __init__(
        self,
        base_url: str,
        token: str = "",
        user_token: str = "",
        user_secret_key: str = "",
        timeout: int = 30,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        headers = {"Accept": "application/json"}
        if token:
            headers["Authorization"] = f"Bearer {token}"
        if user_token and user_secret_key:
            headers["User-Token"] = user_token
            headers["User-Secret-Key"] = user_secret_key
        self.session.headers.update(headers)

    def test_connection(self, alias: str) -> Tuple[bool, str]:
        response = self.session.get(
            f"{self.base_url}/{alias}/orders/filters",
            timeout=self.timeout,
        )
        if response.status_code == 200:
            return True, "Conexao com API OK"
        snippet = response.text[:300].replace("\n", " ").strip()
        return False, f"Erro {response.status_code}: {snippet}"

    def fetch_orders(
        self,
        alias: str,
        page: int = 1,
        page_size: int = 100,
        scroll_id: Optional[str] = None,
        updated_since: Optional[str] = None,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
    ) -> Tuple[List[Dict[str, Any]], Optional[str], int]:
        params: Dict[str, Any] = {"page": page, "limit": page_size}
        if scroll_id:
            params["scroll_id"] = scroll_id
        params["include"] = "items"
        if updated_since and not start_date and not end_date:
            # Parametro comum para cargas incrementais; ajuste conforme contrato da API.
            params["updated_at_min"] = updated_since
        if start_date:
            params["date"] = f"created_at:{start_date}|{end_date or start_date}"
        if end_date:
            params["date"] = f"created_at:{start_date or end_date}|{end_date}"
        for index, status_id in enumerate(ACTIVE_STATUS_IDS):
            params[f"status_id[{index}]"] = status_id

        response = self.session.get(
            f"{self.base_url}/{alias}/orders",
            params=params,
            timeout=self.timeout,
        )
        if response.status_code >= 400:
            snippet = response.text[:400].replace("\n", " ").strip()
            raise requests.HTTPError(
                f"{response.status_code} Client Error: {snippet} | url={response.url}",
                response=response,
            )
        payload = response.json()

        next_scroll_id: Optional[str] = None
        total_pages = 1
        if isinstance(payload, list):
            return payload, next_scroll_id, total_pages
        if isinstance(payload, dict):
            meta = payload.get("meta", {}) if isinstance(payload.get("meta"), dict) else {}
            pagination = meta.get("pagination", {}) if isinstance(meta, dict) else {}
            if isinstance(pagination, dict):
                total_pages = int(pagination.get("total_pages") or 1)
            next_scroll_id = (
                meta.get("next_scroll_id")
                or payload.get("next_scroll_id")
                or payload.get("scroll_id")
            )
        if "data" in payload and isinstance(payload["data"], list):
            return payload["data"], next_scroll_id, total_pages
        if "orders" in payload and isinstance(payload["orders"], list):
            return payload["orders"], next_scroll_id, total_pages
        if "results" in payload and isinstance(payload["results"], list):
            return payload["results"], next_scroll_id, total_pages
        return [], next_scroll_id, total_pages
