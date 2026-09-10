"""Пути проекта, .env, TOML-конфиги, кэш сырых ответов, суточные квоты.

Состояние прогона живёт в work/ (кэш ответов API, квоты, курсы) и целиком
удаляемо: снёс work/ — следующий прогон просто сходит в сеть заново.
Итоги — в final/, превью распознанных позиций — в корне или по --preview-path.
"""

from __future__ import annotations

import hashlib
import json
import os
import threading
import time
import tomllib
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
CONFIG = ROOT / "config"
WORK = ROOT / "work"
CACHE = WORK / "cache"
QUOTA = WORK / "quota.json"
FX_CACHE = WORK / "fx.json"
FINAL = ROOT / "final"
START = ROOT / "start"
PREVIEW = ROOT / "превью_позиции.xlsx"
DEFAULT_OUTPUT = FINAL / "BOM_Приложение_1.xlsx"

_write_lock = threading.Lock()


def utcnow() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def ensure_dirs() -> None:
    for d in (WORK, CACHE, FINAL):
        d.mkdir(parents=True, exist_ok=True)


# ── конфиги ──────────────────────────────────────────────────────────────────

_cfg_cache: dict[str, dict] = {}


def load_toml(name: str) -> dict:
    """Читает config/<name>.toml с кэшированием в процессе."""
    if name not in _cfg_cache:
        path = CONFIG / f"{name}.toml"
        if not path.exists():
            raise SystemExit(f"нет файла конфигурации: {path}")
        with path.open("rb") as fh:
            _cfg_cache[name] = tomllib.load(fh)
    return _cfg_cache[name]


def settings() -> dict:
    return load_toml("settings")


# ── .env ─────────────────────────────────────────────────────────────────────

_env: dict[str, str] | None = None


def load_env(path: Path | None = None) -> dict[str, str]:
    """Читает .env проекта. Отсутствие файла — не ошибка: источники без ключа пропустятся."""
    global _env
    if _env is not None and path is None:
        return _env
    target = path or (ROOT / ".env")
    data: dict[str, str] = {}
    if target.exists():
        for line in target.read_text(encoding="utf-8-sig", errors="replace").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, val = line.split("=", 1)
            data[key.strip()] = val.strip().strip('"').strip("'")
    if path is None:
        # переменные окружения имеют приоритет над файлом
        for key in list(data):
            if os.environ.get(key):
                data[key] = os.environ[key]
        _env = data
    return data


def env(key: str, default: str = "") -> str:
    return load_env().get(key) or os.environ.get(key) or default


def env_flag(key: str, default: bool = False) -> bool:
    """Булев флаг из .env: true/1/yes/on."""
    value = env(key, "")
    if value == "":
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


# ── кэш сырых ответов ────────────────────────────────────────────────────────

def cache_path(source_id: str, query: str) -> Path:
    digest = hashlib.sha1(query.encode("utf-8")).hexdigest()[:16]
    return CACHE / source_id / f"{digest}.json"


def cache_get(source_id: str, query: str, ttl_hours: float):
    path = cache_path(source_id, query)
    if not path.exists():
        return None
    try:
        blob = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    age_h = (time.time() - blob.get("at_ts", 0)) / 3600
    if age_h >= ttl_hours:
        return None
    return blob.get("payload")


def cache_put(source_id: str, query: str, payload) -> str:
    if not settings().get("cache", {}).get("keep_raw", True):
        return ""
    path = cache_path(source_id, query)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps({"at": utcnow(), "at_ts": time.time(), "query": query, "payload": payload},
                   ensure_ascii=False),
        encoding="utf-8",
    )
    return str(path.relative_to(WORK)).replace("\\", "/")


# ── суточные квоты ───────────────────────────────────────────────────────────

def quota_take(source_id: str, limit: int, n: int = 1) -> bool:
    """Списывает n единиц суточной квоты. False — лимит исчерпан, запрос делать нельзя."""
    if not limit:
        return True
    ensure_dirs()
    today = datetime.now(timezone.utc).date().isoformat()
    with _write_lock:
        data = {}
        if QUOTA.exists():
            try:
                data = json.loads(QUOTA.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                data = {}
        row = data.get(source_id, {})
        if row.get("date") != today:
            row = {"date": today, "used": 0}
        if row["used"] + n > limit:
            return False
        row["used"] += n
        data[source_id] = row
        QUOTA.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    return True
