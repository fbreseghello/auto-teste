from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict

from dotenv import load_dotenv


@dataclass(frozen=True)
class ClientConfig:
    id: str
    company: str
    branch: str
    alias: str
    name: str
    platform: str
    base_url: str
    token: str
    token_env: str
    user_token: str
    user_token_env: str
    user_secret_key: str
    user_secret_key_env: str
    page_size: int = 100


def _slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_")


def load_clients_config(config_path: str = "config/clients.json") -> Dict[str, ClientConfig]:
    load_dotenv()
    file_path = Path(config_path)
    if not file_path.exists():
        raise FileNotFoundError(
            f"Arquivo de configuracao nao encontrado: {config_path}. "
            "Copie config/clients.example.json para config/clients.json."
        )

    raw_clients = json.loads(file_path.read_text(encoding="utf-8"))
    clients: Dict[str, ClientConfig] = {}

    for item in raw_clients:
        company = item["company"].strip()
        branch = item["branch"].strip()
        client_id = item.get("id", f"{_slug(company)}_{_slug(branch)}").strip()
        token_env = item.get("token_env", "").strip()
        user_token_env = item.get("user_token_env", "").strip()
        user_secret_key_env = item.get("user_secret_key_env", "").strip()
        token = os.getenv(token_env, "").strip() if token_env else ""
        user_token = os.getenv(user_token_env, "").strip() if user_token_env else ""
        user_secret_key = os.getenv(user_secret_key_env, "").strip() if user_secret_key_env else ""

        client = ClientConfig(
            id=client_id,
            company=company,
            branch=branch,
            alias=item.get("alias", branch).strip(),
            name=item.get("name", f"{company} - {branch}"),
            platform=item["platform"],
            base_url=item["base_url"].rstrip("/"),
            token=token,
            token_env=token_env,
            user_token=user_token,
            user_token_env=user_token_env,
            user_secret_key=user_secret_key,
            user_secret_key_env=user_secret_key_env,
            page_size=int(item.get("page_size", 100)),
        )
        if client.id in clients:
            raise ValueError(f"ID duplicado no config: '{client.id}'.")
        clients[client.id] = client

    return clients
