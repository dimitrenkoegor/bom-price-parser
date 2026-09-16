"""Регламент отбора одного предложения на позицию. Единственный источник правды.

Правила — .agents/rules/price-selection.md и part-matching.md; параметры —
[selection] в config/settings.toml:
  1. нет артикула или производителя → RFQ (сюда такие позиции не доходят);
  2. артикул принимается при match_quality ∈ accept_quality (exact, mask,
     suffix, prefix_series) — аналоги и «частичные» совпадения отвергаются;
  3. производитель обязан совпасть (с учётом поглощений); площадка без поля
     производителя (oemsecrets) принимается только при exact;
  4. продавец — только авторизованный: франшизные API по типу площадки,
     агрегаторы — по списку authorized_sellers;
  5. количество закупки — запрошенное, но не меньше MOQ, кратности заказа и
     нижнего ценового брекета; цена — по брекету при этом количестве, не-USD
     переводится в USD;
  6. лучший: сначала наличие ≥ количества закупки, затем минимальная
     стоимость закупки (цена × количество закупки), при равной — меньшая
     цена за штуку, затем больший склад. Так при запросе 300 шт Cut Tape по $6
     (300 × 6 = 1800) побеждает катушку по $4 от 500 шт (500 × 4 = 2000), а
     катушка по $3 от 500 шт (1500) побеждает Cut Tape;
  7. MOQ выше запроса — не отказ, а пометка в сноске: в отчёт идёт реальный
     минимум заказа.
"""

from __future__ import annotations

import math

from . import fx as fxmod
from . import mpn as mpnmod
from . import offer as offermod
from . import registry, store

DEFAULT_ACCEPT = ("exact", "mask", "suffix", "prefix_series")


def config() -> dict:
    return store.settings().get("selection", {})


def accept_set(cfg: dict | None = None) -> set[str]:
    return set((cfg or config()).get("accept_quality") or DEFAULT_ACCEPT)


def seller_allowed(rec: dict, src: registry.Source | None, cfg: dict) -> bool:
    """Франшизный API — авторизован по определению; агрегатор — по имени продавца."""
    if src is not None and src.kind == "franchise":
        return True
    seller = (rec.get("seller") or "").lower()
    return any(name.lower() in seller for name in cfg.get("authorized_sellers") or [])


def acceptable(rec: dict, part: mpnmod.Part, src: registry.Source | None,
               cfg: dict | None = None) -> tuple[bool, str]:
    """→ (годится ли предложение, причина отказа)."""
    cfg = cfg or config()
    quality = rec.get("match_quality")
    if quality not in accept_set(cfg):
        return False, f"артикул: {quality}"
    if not rec.get("manufacturer"):
        if quality != "exact":
            return False, "производитель площадкой не указан, а артикул не точный"
    elif rec.get("manufacturer_match") != "yes":
        return False, f"производитель {rec.get('manufacturer')!r} ≠ {part.manufacturer!r}"
    if not seller_allowed(rec, src, cfg):
        return False, f"продавец {rec.get('seller')!r} не авторизован"
    if not rec.get("price_breaks"):
        return False, "нет цены"
    return True, ""


def enrich(rec: dict, part: mpnmod.Part, rates: fxmod.Rates) -> dict | None:
    """Количество закупки, цена при нём в USD, стоимость закупки, наличие, срок.
    None — цену не посчитать."""
    breaks = [b for b in rec.get("price_breaks") or []
              if b.get("price") is not None and offermod.to_int(b.get("qty"), 0) > 0]
    if not breaks:
        return None
    qty = part.qty or 1
    # Реальный минимум заказа: наибольшее из MOQ площадки, кратности заказа
    # (Mult у Mouser, multiples у TME, packSize у Farnell; у DigiKey кратности нет)
    # и нижнего ценового брекета. Меньше этого площадка не продаст.
    moq = int(rec.get("moq") or 1)
    multiple = int(rec.get("order_multiple") or 0)
    lowest = min(offermod.to_int(b["qty"]) for b in breaks)
    min_order = max(moq, multiple, lowest)
    order_qty = max(qty, min_order)
    if multiple > 1:
        order_qty = math.ceil(order_qty / multiple) * multiple
    price, break_qty, _below = offermod.price_at(breaks, order_qty)
    if price is None:
        return None
    usd = rates.to_usd(price, rec.get("currency") or "USD")
    if usd is None:
        return None
    stock = int(rec.get("stock_qty") or 0)
    days = rec.get("lead_time_days")
    weeks = None
    if days is not None and days > 0:
        weeks = max(1, math.ceil(int(days) / 7))
    rec["_price_native"] = price
    rec["_price_usd"] = round(usd, 6)
    rec["_break_qty"] = break_qty
    rec["_moq"] = min_order
    rec["_moq_exceeded"] = min_order > qty
    rec["_order_qty"] = order_qty
    rec["_total_usd"] = round(usd * order_qty, 4)
    rec["_stock"] = stock
    rec["_in_stock"] = stock >= order_qty > 0
    rec["_lead_weeks"] = weeks
    return rec


def _key(rec: dict):
    return (0 if rec["_in_stock"] else 1, rec["_total_usd"], rec["_price_usd"], -rec["_stock"])


def best(offers: list[dict], part: mpnmod.Part, rates: fxmod.Rates,
         sources_by_id: dict[str, registry.Source], cfg: dict | None = None) -> dict:
    """Одно предложение по регламенту либо RFQ с причинами отказов."""
    cfg = cfg or config()
    rejected: dict[str, int] = {}
    candidates: list[dict] = []
    for rec in offers:
        src = sources_by_id.get(rec.get("source_id") or "")
        ok, why = acceptable(rec, part, src, cfg)
        if not ok:
            rejected[why] = rejected.get(why, 0) + 1
            continue
        if enrich(rec, part, rates) is None:
            rejected["нет цены при количестве"] = rejected.get("нет цены при количестве", 0) + 1
            continue
        candidates.append(rec)
    if not candidates:
        reason = "; ".join(f"{why} ×{n}" if n > 1 else why for why, n in rejected.items())
        return rfq(part, reason or "ничего не найдено")

    chosen = min(candidates, key=_key)
    src = sources_by_id.get(chosen.get("source_id") or "")
    distributor = src.title if src is not None and src.kind == "franchise" else (chosen.get("seller") or "")
    return {
        "status": "FOUND",
        "pn": part.mpn, "qty": part.qty, "manufacturer": chosen.get("manufacturer") or part.manufacturer,
        "description": part.raw,
        "distributor": distributor, "distr_pn": chosen.get("sku") or "",
        "resolved_mpn": chosen.get("offer_pn") or "",
        "quality": chosen.get("match_quality"),
        "moq": chosen["_moq"], "stock": chosen["_stock"],
        "in_stock": chosen["_in_stock"],
        "lead": None if chosen["_in_stock"] else chosen["_lead_weeks"],
        "price_usd": chosen["_price_usd"], "price_native": chosen["_price_native"],
        "currency": chosen.get("currency") or "USD",
        "min_not_met": chosen["_moq_exceeded"], "break_qty": chosen["_break_qty"],
        "order_qty": chosen["_order_qty"], "total_usd": chosen["_total_usd"],
        "source_id": chosen.get("source_id"), "source_url": chosen.get("source_url") or "",
        "candidates": len(candidates),
        "llm_assisted": False, "llm_reason": "",
    }


def rfq(part: mpnmod.Part, reason: str = "") -> dict:
    return {
        "status": "RFQ", "pn": part.mpn, "qty": part.qty, "manufacturer": part.manufacturer,
        "description": part.raw, "reason": reason, "llm_assisted": False,
    }
