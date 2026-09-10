"""Реестр площадок из config/sources.toml + проверка готовности."""

from __future__ import annotations

from dataclasses import dataclass, field

from . import store

TIERS = ("api", "http", "browser", "ext")


@dataclass
class Source:
    id: str
    title: str
    url: str
    tier: str
    adapter: str
    kind: str = "franchise"
    region: str = "global"
    gives: list[str] = field(default_factory=list)
    auth_env: list[str] = field(default_factory=list)
    rate: dict = field(default_factory=dict)
    variant_policy: str = "all"
    query_url: str = ""
    bom_tool: str = ""
    store_id: str = ""
    currency: str = ""
    accepts_cyrillic_mpn: bool = False
    insecure_tls: bool = False   # российские площадки с отечественным сертификатом
    reachable: bool = True
    blocked_by: str = ""
    enabled: bool = True
    notes: str = ""

    @property
    def rps(self) -> float:
        return float(self.rate.get("rps", 1.0))

    @property
    def daily_limit(self) -> int:
        return int(self.rate.get("daily", 0))

    def missing_keys(self) -> list[str]:
        return [name for name in self.auth_env if not store.env(name)]

    def deep_link(self, mpn: str) -> str:
        if not self.query_url:
            return self.url
        import urllib.parse
        return self.query_url.replace("{mpn}", urllib.parse.quote(mpn, safe=""))

    def gives_dc(self) -> bool:
        return "dc" in self.gives


def load_sources(raw: list[dict] | None = None) -> list[Source]:
    """Реестр из config/sources.toml. raw — готовый список записей (для тестов)."""
    if raw is None:
        raw = store.load_toml("sources").get("source", [])
    out = []
    for row in raw:
        row = dict(row)
        row["store_id"] = row.pop("store", "")
        known = Source.__dataclass_fields__.keys()
        unknown = set(row) - set(known)
        if unknown:
            raise SystemExit(
                f"source '{row.get('id')}': неизвестные поля {sorted(unknown)} — см. docs/sources.md")
        out.append(Source(**row))
    ids = [s.id for s in out]
    dupes = {i for i in ids if ids.count(i) > 1}
    if dupes:
        raise SystemExit(f"дублирующиеся id источников: {sorted(dupes)}")
    return out


def select(tiers: list[str] | None = None, only: list[str] | None = None) -> list[Source]:
    """Источники, готовые к опросу. --only перебивает enabled, но не отсутствие ключа."""
    out = []
    for src in load_sources():
        if only:
            if src.id not in only:
                continue
        else:
            if not src.enabled:
                continue
            if tiers and src.tier not in tiers:
                continue
        out.append(src)
    return out


def readiness(src: Source) -> tuple[str, str]:
    """(состояние, пояснение) — почему источник будет или не будет опрошен."""
    if not src.enabled:
        return "выключен", src.notes[:60]
    if src.missing_keys():
        return "нет ключа", ", ".join(src.missing_keys())
    if not src.reachable:
        return "недоступен", src.blocked_by or "с этой машины не открывается"
    if src.tier in ("browser", "ext"):
        return "браузер", src.blocked_by or "нужен браузер"
    return "готов", ""
