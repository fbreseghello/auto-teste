from __future__ import annotations

import time
from typing import Any, Dict, List, Optional, Tuple

import requests


ACTIVE_STATUS_IDS = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13"]
RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}


class YampiClient:
    def __init__(
        self,
        base_url: str,
        token: str = "",
        user_token: str = "",
        user_secret_key: str = "",
        timeout: int = 30,
        max_retries: int = 4,
        retry_backoff_seconds: float = 1.5,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.max_retries = max_retries
        self.retry_backoff_seconds = retry_backoff_seconds
        self.session = requests.Session()
        self.session.headers.update({"Accept": "application/json"})
        self._primary_auth_headers: Dict[str, str] = {}
        self._fallback_auth_headers: Dict[str, str] = {}

        has_user_auth = bool(user_token and user_secret_key)
        has_bearer = bool(token)

        if has_user_auth:
            self._primary_auth_headers = {
                "User-Token": user_token,
                "User-Secret-Key": user_secret_key,
            }
            if has_bearer:
                # Se user-token falhar com 401, tenta bearer automaticamente.
                self._fallback_auth_headers = {"Authorization": f"Bearer {token}"}
        elif has_bearer:
            self._primary_auth_headers = {"Authorization": f"Bearer {token}"}

    def test_connection(self, alias: str) -> Tuple[bool, str]:
        try:
            response = self._request("GET", f"{self.base_url}/{alias}/orders/filters")
        except requests.RequestException as exc:
            return False, str(exc)
        if response.status_code == 200:
            return True, "Conexao com API OK"
        snippet = response.text[:300].replace("\n", " ").strip()
        return False, f"Erro {response.status_code}: {snippet}"

    def _request(self, method: str, url: str, params: Optional[Dict[str, Any]] = None) -> requests.Response:
        attempt = 0
        while True:
            try:
                headers = dict(self.session.headers)
                headers.update(self._primary_auth_headers)
                response = self.session.request(
                    method=method,
                    url=url,
                    params=params,
                    headers=headers,
                    timeout=self.timeout,
                )
            except requests.RequestException as exc:
                if attempt >= self.max_retries:
                    raise
                wait_seconds = self.retry_backoff_seconds * (2**attempt)
                time.sleep(wait_seconds)
                attempt += 1
                continue

            if response.status_code in RETRYABLE_STATUS_CODES and attempt < self.max_retries:
                wait_seconds = self.retry_backoff_seconds * (2**attempt)
                time.sleep(wait_seconds)
                attempt += 1
                continue

            if response.status_code == 401 and self._fallback_auth_headers:
                fallback_headers = dict(self.session.headers)
                fallback_headers.update(self._fallback_auth_headers)
                fallback_response = self.session.request(
                    method=method,
                    url=url,
                    params=params,
                    headers=fallback_headers,
                    timeout=self.timeout,
                )
                if fallback_response.status_code < 400:
                    return fallback_response
                response = fallback_response

            if response.status_code >= 400:
                snippet = response.text[:250].replace("\n", " ").strip()
                raise requests.HTTPError(
                    f"{response.status_code} Client Error: {snippet} | url={response.url}",
                    response=response,
                )

            return response

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

        response = self._request("GET", f"{self.base_url}/{alias}/orders", params=params)
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
