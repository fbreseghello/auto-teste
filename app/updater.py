from __future__ import annotations

import os
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path

import requests


DEFAULT_GITHUB_REPO_ENV = "AUTO_TESTE_GITHUB_REPO"
GITHUB_API_BASE = "https://api.github.com"
DEFAULT_EXCLUDE_TOP_LEVEL = {
    ".git",
    ".github",
    ".venv",
    "data",
    "exports",
    "__pycache__",
}
DEFAULT_EXCLUDE_FILES = {
    ".env",
    "config/clients.json",
}


@dataclass(frozen=True)
class UpdateCheck:
    repo: str
    current_version: str
    latest_version: str
    has_update: bool
    release_name: str
    release_url: str
    zip_url: str


@dataclass(frozen=True)
class UpdateResult:
    updated: bool
    message: str
    latest_version: str = ""
    files_copied: int = 0


def resolve_repo(explicit_repo: str = "", env_var: str = DEFAULT_GITHUB_REPO_ENV) -> str:
    repo = explicit_repo.strip() or os.getenv(env_var, "").strip()
    if not repo:
        raise ValueError(
            "Repositorio GitHub nao informado. "
            f"Defina {env_var}=dono/repositorio no ambiente/.env ou use --repo."
        )
    if "/" not in repo:
        raise ValueError(f"Repositorio invalido '{repo}'. Use o formato dono/repositorio.")
    return repo


def _normalize_version(value: str) -> str:
    normalized = value.strip().lower()
    if normalized.startswith("v"):
        normalized = normalized[1:]
    return normalized


def _version_tuple(version: str) -> tuple[int, ...]:
    parts: list[int] = []
    for token in _normalize_version(version).replace("-", ".").split("."):
        if token.isdigit():
            parts.append(int(token))
        else:
            digits = "".join(ch for ch in token if ch.isdigit())
            parts.append(int(digits) if digits else 0)
    return tuple(parts or [0])


def _is_newer(candidate: str, current: str) -> bool:
    return _version_tuple(candidate) > _version_tuple(current)


def check_for_updates(
    current_version: str,
    repo: str = "",
    timeout: int = 20,
) -> UpdateCheck:
    resolved_repo = resolve_repo(repo)
    url = f"{GITHUB_API_BASE}/repos/{resolved_repo}/releases/latest"
    response = requests.get(url, timeout=timeout)
    if response.status_code == 404:
        raise ValueError(
            f"Nenhuma release encontrada em '{resolved_repo}'. "
            "Publique pelo menos uma release no GitHub."
        )
    response.raise_for_status()
    payload = response.json()

    latest_tag = str(payload.get("tag_name") or "").strip()
    if not latest_tag:
        raise ValueError("Release sem tag_name. Verifique a release no GitHub.")
    latest_version = _normalize_version(latest_tag)
    current_normalized = _normalize_version(current_version)
    has_update = _is_newer(latest_version, current_normalized)

    zip_url = str(payload.get("zipball_url") or "").strip()
    if not zip_url:
        raise ValueError("Release sem zipball_url. Verifique a API do GitHub.")

    return UpdateCheck(
        repo=resolved_repo,
        current_version=current_normalized,
        latest_version=latest_version,
        has_update=has_update,
        release_name=str(payload.get("name") or latest_tag),
        release_url=str(payload.get("html_url") or ""),
        zip_url=zip_url,
    )


def _should_skip(relative_path: str) -> bool:
    normalized = relative_path.replace("\\", "/").strip("/")
    if not normalized:
        return True
    top_level = normalized.split("/", 1)[0]
    if top_level in DEFAULT_EXCLUDE_TOP_LEVEL:
        return True
    if normalized in DEFAULT_EXCLUDE_FILES:
        return True
    return False


def _copy_tree(source_root: Path, target_root: Path) -> int:
    copied = 0
    for path in source_root.rglob("*"):
        if path.is_dir():
            continue
        relative = path.relative_to(source_root).as_posix()
        if _should_skip(relative):
            continue
        destination = target_root / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, destination)
        copied += 1
    return copied


def _detect_source_root(extract_dir: Path) -> Path:
    entries = [item for item in extract_dir.iterdir() if item.is_dir()]
    if len(entries) == 1:
        return entries[0]
    return extract_dir


def apply_update_from_github(
    current_version: str,
    project_dir: str = ".",
    repo: str = "",
    force: bool = False,
    timeout: int = 60,
) -> UpdateResult:
    check = check_for_updates(current_version=current_version, repo=repo)
    if not check.has_update and not force:
        return UpdateResult(
            updated=False,
            latest_version=check.latest_version,
            message=f"Sem atualizacao. Versao atual: {check.current_version}.",
            files_copied=0,
        )

    target = Path(project_dir).resolve()
    with tempfile.TemporaryDirectory(prefix="auto_teste_update_") as tmp:
        tmp_path = Path(tmp)
        zip_path = tmp_path / "update.zip"
        response = requests.get(check.zip_url, timeout=timeout)
        response.raise_for_status()
        zip_path.write_bytes(response.content)

        extract_dir = tmp_path / "extract"
        extract_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(extract_dir)

        source_root = _detect_source_root(extract_dir)
        copied = _copy_tree(source_root, target)

    return UpdateResult(
        updated=True,
        latest_version=check.latest_version,
        files_copied=copied,
        message=(
            f"Atualizacao aplicada para {check.latest_version}. "
            f"Arquivos atualizados: {copied}."
        ),
    )
