"""Реестр адаптеров: имя из config/sources.toml → функция поиска.

Контракт адаптера (общий с count pars):
    search(src: Source, query: str) -> tuple[list[dict], object]
    Возвращает (список заготовок предложений, сырой ответ для кэша).
    Ничего не найдено → ([], payload). Источник сломан → исключение.
Только официальные API: браузерных и HTTP-парсящих адаптеров здесь нет намеренно.
"""

from __future__ import annotations

from . import api_aggregator, api_franchise

REGISTRY = {
    "mouser": api_franchise.mouser,
    "digikey": api_franchise.digikey,
    "farnell": api_franchise.farnell,
    "tme": api_franchise.tme,
    "nexar": api_aggregator.nexar,
    "oemsecrets": api_aggregator.oemsecrets,
    "arrow": api_aggregator.arrow,
}


def get(name: str):
    if name not in REGISTRY:
        raise SystemExit(f"нет адаптера '{name}' — см. .agents/references/environment.md")
    return REGISTRY[name]
