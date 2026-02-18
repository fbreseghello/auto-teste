from __future__ import annotations

import json
import os
import re
import shutil
import sys
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


DEFAULT_CONFIG_PATH = "config/clients.json"
DEFAULT_CONFIG_TEMPLATE_PATH = "config/clients.example.json"
DEFAULT_ENV_PATH = ".env"
DEFAULT_BUNDLE_DIR = "clientes"
PLACEHOLDER_PREFIXES = ("COLE_AQUI", "SEU_", "YOUR_")


def _slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_")


def _runtime_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def _resolve_path(path: str) -> Path:
    file_path = Path(path)
    if file_path.is_absolute():
        return file_path
    return (_runtime_root() / file_path).resolve()


def _clean_secret(value: str) -> str:
    text = value.strip()
    upper = text.upper()
    for prefix in PLACEHOLDER_PREFIXES:
        if upper.startswith(prefix):
            return ""
    return text


def resolve_runtime_paths(
    config_path: str = DEFAULT_CONFIG_PATH,
    env_path: str = DEFAULT_ENV_PATH,
    bundle_dir: str = DEFAULT_BUNDLE_DIR,
) -> tuple[Path, Path]:
    config_file = _resolve_path(config_path)
    env_file = _resolve_path(env_path)
    bundle_root = _resolve_path(bundle_dir)
    bundle_config = bundle_root / "clients.json"
    bundle_env = bundle_root / ".env"
    if bundle_config.exists():
        config_file = bundle_config
    if bundle_env.exists():
        env_file = bundle_env
    return config_file, env_file


def ensure_runtime_files(
    config_path: str = DEFAULT_CONFIG_PATH,
    template_path: str = DEFAULT_CONFIG_TEMPLATE_PATH,
    env_path: str = DEFAULT_ENV_PATH,
) -> tuple[Path, Path]:
    config_file, env_file = resolve_runtime_paths(config_path=config_path, env_path=env_path)
    config_file.parent.mkdir(parents=True, exist_ok=True)
    if not config_file.exists():
        example_file = _resolve_path(template_path)
        if not example_file.exists():
            raise FileNotFoundError(
                f"Arquivo de configuracao nao encontrado: {config_path}. "
                f"Template ausente: {template_path}."
            )
        shutil.copy2(example_file, config_file)

    env_file.parent.mkdir(parents=True, exist_ok=True)
    if not env_file.exists():
        env_file.write_text("AUTO_TESTE_GITHUB_REPO=fbreseghello/auto-teste\n", encoding="utf-8")
    return config_file, env_file


def set_env_values(values: dict[str, str], env_path: str = DEFAULT_ENV_PATH) -> None:
    file_path = _resolve_path(env_path)
    file_path.parent.mkdir(parents=True, exist_ok=True)
    if not file_path.exists():
        file_path.write_text("", encoding="utf-8")

    lines = file_path.read_text(encoding="utf-8").splitlines()
    key_to_index: dict[str, int] = {}
    for index, line in enumerate(lines):
        text = line.strip()
        if not text or text.startswith("#") or "=" not in text:
            continue
        key = text.split("=", 1)[0].strip()
        if key:
            key_to_index[key] = index

    for key, value in values.items():
        line = f"{key}={value}"
        if key in key_to_index:
            lines[key_to_index[key]] = line
        else:
            lines.append(line)
        os.environ[key] = value

    content = "\n".join(lines).rstrip() + "\n"
    file_path.write_text(content, encoding="utf-8")


def save_client_credentials(
    client: ClientConfig,
    token: str = "",
    user_token: str = "",
    user_secret_key: str = "",
    env_path: str = DEFAULT_ENV_PATH,
) -> None:
    updates: dict[str, str] = {}
    if client.token_env:
        updates[client.token_env] = token.strip()
    if client.user_token_env:
        updates[client.user_token_env] = user_token.strip()
    if client.user_secret_key_env:
        updates[client.user_secret_key_env] = user_secret_key.strip()
    if not updates:
        raise ValueError("Cliente sem variaveis de ambiente configuradas para credenciais.")
    _, runtime_env_path = resolve_runtime_paths(env_path=env_path)
    set_env_values(updates, env_path=str(runtime_env_path))


def load_clients_config(config_path: str = DEFAULT_CONFIG_PATH) -> Dict[str, ClientConfig]:
    file_path, env_file = ensure_runtime_files(config_path=config_path)
    load_dotenv(dotenv_path=env_file, override=True)

    raw_clients = json.loads(file_path.read_text(encoding="utf-8"))
    clients: Dict[str, ClientConfig] = {}

    for item in raw_clients:
        company = item["company"].strip()
        branch = item["branch"].strip()
        client_id = item.get("id", f"{_slug(company)}_{_slug(branch)}").strip()
        token_env = item.get("token_env", "").strip()
        user_token_env = item.get("user_token_env", "").strip()
        user_secret_key_env = item.get("user_secret_key_env", "").strip()
        token = _clean_secret(os.getenv(token_env, "")) if token_env else ""
        user_token = _clean_secret(os.getenv(user_token_env, "")) if user_token_env else ""
        user_secret_key = _clean_secret(os.getenv(user_secret_key_env, "")) if user_secret_key_env else ""

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
