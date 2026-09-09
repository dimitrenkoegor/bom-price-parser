"""Адаптеры официальных API дистрибьюторов.

Формы запросов сверены с рабочим кодом соседнего проекта — там эти ключи
уже годами ходят в бой. Отличие принципиальное: здесь мы НЕ отбрасываем
предложения по несовпадению артикула, а возвращаем всё и метим качеством
совпадения. Отбор — работа отчёта, не адаптера.
"""

from __future__ import annotations

import base64
import time
import urllib.parse

from .. import net, store

_tokens: dict[str, dict] = {}


def _cached_token(key: str, factory, ttl_default: int = 300) -> str | None:
    now = time.time()
    cur = _tokens.get(key)
    if cur and cur["expires_at"] > now + 30:
        return cur["token"]
    got = factory()
    if not got:
        return None
    token, expires = got
    _tokens[key] = {"token": token, "expires_at": now + int(expires or ttl_default)}
    return token


# ── Mouser ───────────────────────────────────────────────────────────────────

def mouser(src, query: str):
    key = store.env("MOUSER_API_KEY")
    if not key:
        return [], None
    attempts = (
        ("partnumber", {"SearchByPartRequest": {"mouserPartNumber": query,
                                                "partSearchOptions": ""}}),
        ("keyword", {"SearchByKeywordRequest": {"keyword": query, "records": 20,
                                                "startingRecord": 0}}),
    )
    payload = None
    parts = []
    for endpoint, body in attempts:
        url = (f"https://api.mouser.com/api/v1.0/search/{endpoint}"
               f"?apiKey={urllib.parse.quote(key)}")
        payload = net.post_json(url, body, rps=src.rps)
        if payload.get("Errors"):
            continue
        parts = (payload.get("SearchResults") or {}).get("Parts") or []
        if parts:
            break

    offers = []
    for p in parts:
        breaks = [{"qty": pb.get("Quantity"), "price": pb.get("Price"),
                   "currency": pb.get("Currency") or "USD"}
                  for pb in p.get("PriceBreaks") or []]
        offers.append({
            "offer_mpn": (p.get("ManufacturerPartNumber") or "").strip(),
            "offer_pn": (p.get("ManufacturerPartNumber") or "").strip(),
            "sku": p.get("MouserPartNumber") or "",
            "manufacturer": p.get("Manufacturer") or "",
            "stock_qty": p.get("Availability"),
            "stock_note": p.get("FactoryStock") or "",
            "moq": p.get("Min"), "order_multiple": p.get("Mult"),
            "lead_time_raw": p.get("LeadTime") or "",
            "price_breaks": breaks,
            "currency": (breaks[0]["currency"] if breaks else "USD"),
            "lifecycle": p.get("LifecycleStatus") or "",
            "rohs": p.get("ROHSStatus") or "",
            "source_url": p.get("ProductDetailUrl") or src.deep_link(query),
            "packaging": p.get("Packaging") or "",
        })
    return offers, payload


# ── DigiKey ──────────────────────────────────────────────────────────────────

def _digikey_token():
    cid, secret = store.env("DIGIKEY_CLIENT_ID"), store.env("DIGIKEY_CLIENT_SECRET")
    if not cid or not secret:
        return None

    def make():
        body = net.post_form("https://api.digikey.com/v1/oauth2/token", {
            "client_id": cid, "client_secret": secret,
            "grant_type": "client_credentials"}, rps=2.0)
        import json
        data = json.loads(body.decode())
        return data["access_token"], data.get("expires_in", 600)

    return _cached_token("digikey", make, 600)


def digikey(src, query: str):
    cid = store.env("DIGIKEY_CLIENT_ID")
    token = _digikey_token()
    if not (cid and token):
        return [], None

    payload = net.post_json(
        "https://api.digikey.com/products/v4/search/keyword",
        {"Keywords": query, "Limit": 10},
        headers={"Authorization": "Bearer " + token, "X-DIGIKEY-Client-Id": cid,
                 "X-DIGIKEY-Locale-Site": "US", "X-DIGIKEY-Locale-Language": "en",
                 "X-DIGIKEY-Locale-Currency": "USD"},
        rps=src.rps)

    offers = []
    for p in payload.get("Products") or []:
        # v4: цены, сток и MOQ живут в ProductVariations (Cut Tape / Tape & Reel /
        # Digi-Reel), на уровне товара их нет. Одна вариация = одно предложение.
        lead_weeks = p.get("ManufacturerLeadWeeks")
        common = {
            "offer_mpn": (p.get("ManufacturerProductNumber") or "").strip(),
            "offer_pn": (p.get("ManufacturerProductNumber") or "").strip(),
            "manufacturer": (p.get("Manufacturer") or {}).get("Name") or "",
            "lead_time_raw": f"{lead_weeks} weeks" if lead_weeks else "",
            "currency": "USD",
            "lifecycle": (p.get("ProductStatus") or {}).get("Status") or "",
            "rohs": p.get("Classifications", {}).get("RohsStatus", "") if isinstance(
                p.get("Classifications"), dict) else "",
            "source_url": p.get("ProductUrl") or src.deep_link(query),
        }
        variations = p.get("ProductVariations") or []
        if not variations:
            breaks = [{"qty": sp.get("BreakQuantity"), "price": sp.get("UnitPrice"),
                       "currency": "USD"} for sp in p.get("StandardPricing") or []]
            if not breaks and p.get("UnitPrice"):
                breaks = [{"qty": 1, "price": p["UnitPrice"], "currency": "USD"}]
            offers.append(dict(common, sku=p.get("DigiKeyPartNumber") or "",
                               stock_qty=p.get("QuantityAvailable"),
                               moq=p.get("MinimumOrderQuantity"), price_breaks=breaks))
            continue
        for v in variations:
            breaks = [{"qty": sp.get("BreakQuantity"), "price": sp.get("UnitPrice"),
                       "currency": "USD"} for sp in v.get("StandardPricing") or []]
            if not breaks:
                continue                     # вариация без прайса — не предложение
            offers.append(dict(
                common, sku=v.get("DigiKeyProductNumber") or "",
                stock_qty=v.get("QuantityAvailableforPackageType"),
                moq=v.get("MinimumOrderQuantity") or (breaks[0].get("qty") if breaks else None),
                order_multiple=v.get("StandardPackage"),
                packaging=(v.get("PackageType") or {}).get("Name") or "",
                price_breaks=breaks,
            ))
    return offers, payload


# ── Farnell / Newark (element14) ─────────────────────────────────────────────

def farnell(src, query: str):
    key = store.env("FARNELL_API_KEY")
    if not key:
        return [], None
    store_id = src.store_id or "uk.farnell.com"

    def ask(term_field: str, limit: int):
        params = (f"term={term_field}%3A{urllib.parse.quote(query)}"
                  f"&storeInfo.id={store_id}&resultsSettings.offset=0"
                  f"&resultsSettings.numberOfResults={limit}"
                  f"&resultsSettings.responseGroup=large"
                  f"&callInfo.responseDataFormat=JSON&callinfo.apiKey={key}")
        data = net.get_json("https://api.element14.com/catalog/products?" + params,
                            rps=src.rps)
        for value in (data or {}).values():
            if isinstance(value, dict) and "products" in value:
                return value.get("products") or [], data
        return [], data

    products, payload = ask("manuPartNum", 10)
    if not products:
        products, payload = ask("any", 15)

    offers = []
    for p in products:
        stock = p.get("stock") or {}
        breaks = [{"qty": pb.get("from"), "price": pb.get("cost"),
                   "currency": src.currency or "GBP"} for pb in p.get("prices") or []]
        offers.append({
            "offer_mpn": (p.get("translatedManufacturerPartNumber") or "").strip(),
            "offer_pn": (p.get("translatedManufacturerPartNumber") or "").strip(),
            "sku": p.get("sku") or "",
            "manufacturer": p.get("brandName") or "",
            "stock_qty": stock.get("level"),
            "lead_time_raw": (f"{stock.get('leastLeadTime')} days"
                              if stock.get("leastLeadTime") else ""),
            "moq": p.get("translatedMinimumOrderQuality"),
            "order_multiple": p.get("packSize"),
            "price_breaks": breaks, "currency": src.currency or "GBP",
            "rohs": p.get("rohsStatusCode") or "",
            "source_url": (f"https://{store_id}/w/search?st="
                           f"{urllib.parse.quote(query)}"),
        })
    return offers, payload


# ── TME ──────────────────────────────────────────────────────────────────────

def _tme_token():
    token, secret = store.env("TME_TOKEN"), store.env("TME_APP_SECRET")
    if not (token and secret):
        return None

    def make():
        basic = base64.b64encode(f"{token}:{secret}".encode()).decode()
        body = net.post_form("https://api.tme.eu/auth/token",
                             {"grant_type": "client_credentials"},
                             headers={"Authorization": "Basic " + basic}, rps=2.0)
        import json
        data = json.loads(body.decode())
        return data["access_token"], data.get("expires_in", 300)

    return _cached_token("tme", make, 300)


def tme(src, query: str):
    access = _tme_token()
    if not access:
        return [], None
    auth = {"Authorization": "Bearer " + access}

    data = net.get_json("https://api.tme.eu/products?" + urllib.parse.urlencode(
        [("country", "DE"), ("mpns[]", query)]), headers=auth, rps=src.rps)
    products = (data.get("data") or {}).get("elements") or []

    if not products:
        phrase = query.strip()[:40]
        if len(phrase) >= 2:
            data = net.get_json("https://api.tme.eu/products/search?" + urllib.parse.urlencode(
                [("country", "DE"), ("phrase", phrase), ("scope[]", "products"),
                 ("limit", "20")]), headers=auth, rps=src.rps)
            products = ((data.get("data") or {}).get("products") or {}).get("elements") or []
    if not products:
        return [], data

    by_symbol = {p["symbol"]: p for p in products if p.get("symbol")}
    qs = [("country", "DE"), ("currency", "USD"), ("scope[]", "prices"), ("scope[]", "stock")]
    qs += [("symbols[]", s) for s in list(by_symbol)[:50]]
    priced = net.get_json("https://api.tme.eu/products/data?" + urllib.parse.urlencode(qs),
                          headers=auth, rps=src.rps)

    offers = []
    for el in (priced.get("data") or {}).get("elements") or []:
        prod = by_symbol.get(el.get("symbol"))
        if not prod:
            continue
        breaks = [{"qty": t.get("amount"), "price": t.get("price"), "currency": "USD"}
                  for t in (el.get("prices") or {}).get("elements") or []]
        symbols = [s for s in (prod.get("manufacturer_symbols") or []) if s] or [prod.get("symbol")]
        offers.append({
            "offer_mpn": (symbols[0] or "").strip(),
            "offer_pn": (symbols[0] or "").strip(),
            # у TME одна карточка может нести несколько написаний артикула
            # производителя; котировщик выбирает из них лучшее совпадение
            "alt_mpns": [s.strip() for s in symbols[1:]],
            "sku": prod.get("symbol") or "",
            "manufacturer": (prod.get("manufacturer") or {}).get("name") or "",
            "stock_qty": el.get("stock_quantity"),
            "moq": prod.get("minimal_amount"),
            "order_multiple": prod.get("multiples"),
            "price_breaks": breaks, "currency": "USD",
            "source_url": prod.get("product_information_page") or src.deep_link(query),
        })
    return offers, priced


# ── LCSC (без ключа) ─────────────────────────────────────────────────────────

def lcsc(src, query: str):
    """Публичный поисковый JSON витрины. Ключ не нужен, но эндпоинт неофициальный:
    при смене формата источник честно покажет ошибку в листе «Источники»."""
    url = ("https://wmsc.lcsc.com/ftps/wm/search/global?keyword="
           + urllib.parse.quote(query))
    payload = net.get_json(url, rps=src.rps,
                           headers={"Referer": "https://www.lcsc.com/"})

    result = payload.get("result") or {}
    products = (result.get("productSearchResultVO") or {}).get("productList") or []
    if not products and result.get("tipProductDetailUrlVO"):
        products = [result["tipProductDetailUrlVO"]]

    offers = []
    for p in products:
        breaks = [{"qty": t.get("ladder"), "price": t.get("usdPrice") or t.get("productPrice"),
                   "currency": "USD"} for t in p.get("productPriceList") or []]
        code = p.get("productCode") or ""
        offers.append({
            "offer_mpn": (p.get("productModel") or "").strip(),
            "offer_pn": (p.get("productModel") or "").strip(),
            "sku": code,
            "manufacturer": p.get("brandNameEn") or "",
            "stock_qty": p.get("stockNumber") or p.get("stockJs"),
            "moq": p.get("minPacketNumber"), "order_multiple": p.get("minBuyNumber"),
            "packaging": p.get("encapStandard") or "",
            "price_breaks": breaks, "currency": "USD",
            "source_url": (f"https://www.lcsc.com/product-detail/{code}.html"
                           if code else src.deep_link(query)),
        })
    return offers, payload
