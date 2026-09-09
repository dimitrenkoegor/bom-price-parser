"""Опрос API по одной позиции: источники × варианты написания → предложения в памяти.

Журналов нет — котировщику нужен ответ здесь и сейчас. Зато есть кэш сырых
ответов (work/cache, TTL из settings) и суточные квоты: повторный прогон того же
запроса в сеть не ходит, а oemsecrets не тратит дневной лимит впустую.

Второй проход: если предложения нашлись по написанию, отличному от запрошенного
(суффикс, маска, префикс серии), точный найденный артикул задаётся всем
источникам ещё раз — точечный поиск у Mouser/TME/Farnell даёт больше, чем keyword.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from . import adapters, mpn as mpnmod, offer as offermod, registry, store

_QUALITY_RANK = {q: i for i, q in enumerate(mpnmod.QUALITIES)}


@dataclass
class Stats:
    queries: int = 0
    cached: int = 0
    errors: int = 0
    by_source: dict = field(default_factory=dict)

    def bump(self, source_id: str, key: str, n: int = 1, text: str = "") -> None:
        row = self.by_source.setdefault(
            source_id, {"queries": 0, "cached": 0, "offers": 0, "errors": 0, "error_text": ""})
        row[key] = row.get(key, 0) + n
        if text:
            row["error_text"] = text[:160]


def usable_sources(kinds: tuple[str, ...] = ("franchise",)) -> list[registry.Source]:
    """Площадки тира api, включённые, с ключами и достижимые. kinds — franchise | aggregator."""
    out = []
    for src in registry.load_sources():
        if src.tier != "api" or not src.enabled or src.kind not in kinds:
            continue
        if src.missing_keys() or not src.reachable:
            continue
        limit = store.env("OEMSECRETS_DAILY_LIMIT") if src.id == "oemsecrets" else ""
        if limit and limit.isdigit():
            src.rate = dict(src.rate, daily=int(limit))
        out.append(src)
    return out


def _variants_for(part: mpnmod.Part, src: registry.Source) -> list[dict]:
    variants = part.variants or mpnmod.build_variants(part.mpn)
    if part.mpn_script == "cyrillic" and not src.accepts_cyrillic_mpn:
        variants = [v for v in variants if not mpnmod.has_cyrillic(v["mpn"])]
    return variants


def _best_symbol(requested: str, raw: dict) -> tuple[str, str, str]:
    """У площадки может быть несколько написаний артикула (TME): берём лучшее совпадение."""
    candidates = [raw.get("offer_mpn") or raw.get("offer_pn") or ""] + list(raw.get("alt_mpns") or [])
    best = ("", "mismatch", "")
    for cand in candidates:
        cand = (cand or "").strip()
        if not cand:
            continue
        quality, delta = mpnmod.match_quality(requested, cand, strict=True)
        if _QUALITY_RANK[quality] < _QUALITY_RANK[best[1]] or not best[0]:
            best = (cand, quality, delta)
    return best


def _finalize(raw: dict, part: mpnmod.Part, src: registry.Source, variant_id: str,
              raw_ref: str) -> dict | None:
    offer_pn, quality, delta = _best_symbol(part.mpn, raw)
    if quality == "mismatch":
        return None
    currency = raw.get("currency") or src.currency or "USD"
    lead_days = raw.get("lead_time_days")
    if lead_days is None:
        lead_days = offermod.lead_days(raw.get("lead_time_raw"))
    return offermod.new(
        run_id="", fetched_at=store.utcnow(),
        part_id=part.part_id, req_pn=part.mpn, req_qty=part.qty,
        req_manufacturer=part.manufacturer,
        source_id=src.id, source_title=src.title,
        source_url=raw.get("source_url") or src.deep_link(part.mpn),
        seller=raw.get("seller") or src.title,
        seller_kind=raw.get("seller_kind") or src.kind,
        offer_pn=offer_pn, offer_mpn=offer_pn, sku=raw.get("sku") or "",
        manufacturer=raw.get("manufacturer") or "",
        manufacturer_match=mpnmod.manufacturer_match(part.manufacturer, raw.get("manufacturer") or ""),
        match_quality=quality, match_delta=delta, variant=variant_id,
        stock_qty=raw.get("stock_qty"), stock_note=raw.get("stock_note") or "",
        moq=raw.get("moq"), order_multiple=raw.get("order_multiple"),
        packaging=raw.get("packaging") or "",
        lead_time_days=lead_days, lead_time_raw=str(raw.get("lead_time_raw") or ""),
        date_code=raw.get("date_code") or "",
        date_code_source=offermod.dc_source(src, raw.get("date_code") or ""),
        price_breaks=offermod.clean_breaks(raw.get("price_breaks"), currency),
        currency=currency, lifecycle=raw.get("lifecycle") or "", rohs=raw.get("rohs") or "",
        provenance="api", raw_ref=raw_ref,
    )


def query(src: registry.Source, text: str, part: mpnmod.Part, variant_id: str,
          stats: Stats, use_cache: bool = True) -> list[dict]:
    """Один запрос к площадке. Исключения наружу — их считает вызывающий."""
    ttl = float(store.settings().get("cache", {}).get("price_ttl_hours", 24))
    raws = store.cache_get(src.id, text, ttl) if use_cache else None
    if raws is not None:
        stats.cached += 1
        stats.bump(src.id, "cached")
    else:
        if src.daily_limit and not store.quota_take(src.id, src.daily_limit):
            raise RuntimeError(f"суточная квота исчерпана ({src.daily_limit}/сут) — остаток завтра")
        raws, _payload = adapters.get(src.adapter)(src, text)
        store.cache_put(src.id, text, raws)
        stats.queries += 1
        stats.bump(src.id, "queries")
    out = []
    for raw in raws or []:
        rec = _finalize(raw, part, src, variant_id, store.cache_path(src.id, text).name)
        if rec:
            out.append(rec)
    stats.bump(src.id, "offers", len(out))
    return out


def collect_part(part: mpnmod.Part, sources: list[registry.Source], stats: Stats,
                 use_cache: bool = True, log=print) -> list[dict]:
    """Все предложения площадок по позиции. По каждой площадке варианты написания
    перебираются до первого, давшего хоть что-то."""
    offers: list[dict] = []
    for src in sources:
        for variant in _variants_for(part, src):
            try:
                got = query(src, variant["mpn"], part, variant["id"], stats, use_cache)
            except Exception as exc:                  # noqa: BLE001 — площадка не роняет прогон
                stats.errors += 1
                stats.bump(src.id, "errors", text=f"{type(exc).__name__}: {exc}")
                log(f"    {src.id}: {type(exc).__name__}: {str(exc)[:120]}")
                break
            offers.extend(got)
            if got:
                break
    return offers


def resolved_mpns(offers: list[dict], part: mpnmod.Part, accept: set[str],
                  limit: int = 3) -> list[str]:
    """Точные артикулы из принятых предложений, отличные от запрошенного написания."""
    seen: set[str] = set()
    out: list[str] = []
    requested = mpnmod.canonical(part.mpn)
    for rec in offers:
        if rec.get("match_quality") not in accept:
            continue
        pn = (rec.get("offer_pn") or "").strip()
        key = mpnmod.canonical(pn)
        if not pn or key == requested or key in seen:
            continue
        seen.add(key)
        out.append(pn)
        if len(out) >= limit:
            break
    return out


def collect_with_resolution(part: mpnmod.Part, sources: list[registry.Source], stats: Stats,
                            accept: set[str], use_cache: bool = True, log=print) -> list[dict]:
    """Основной проход + повторный опрос найденными точными артикулами."""
    offers = collect_part(part, sources, stats, use_cache, log)
    limit = int(store.settings().get("selection", {}).get("resolved_limit", 3))
    extra: list[dict] = []
    for rmpn in resolved_mpns(offers, part, accept, limit):
        for src in sources:
            try:
                extra.extend(query(src, rmpn, part, "resolved", stats, use_cache))
            except Exception as exc:                  # noqa: BLE001
                stats.errors += 1
                stats.bump(src.id, "errors", text=f"{type(exc).__name__}: {exc}")
    seen: set[tuple] = set()
    merged: list[dict] = []
    for rec in offers + extra:
        key = (rec.get("source_id"), rec.get("seller"), rec.get("sku") or "",
               mpnmod.canonical(rec.get("offer_pn") or ""))
        if key in seen:
            continue
        seen.add(key)
        merged.append(rec)
    return merged


def summary(stats: Stats) -> str:
    lines = [f"Запросов: {stats.queries}, из кэша: {stats.cached}, ошибок: {stats.errors}"]
    for source_id, row in sorted(stats.by_source.items()):
        note = f"  {row['error_text']}" if row.get("error_text") else ""
        lines.append(f"  {source_id:<12} запросов {row['queries']:>4}  из кэша {row['cached']:>4}  "
                     f"предложений {row['offers']:>5}  ошибок {row['errors']:>3}{note}")
    return "\n".join(lines)
