"""Агрегаторы по ключу: Nexar (Octopart), OEMsecrets, Arrow.

Все три выключены в config/sources.toml до появления ключа. Код готов:
заполнить .env и переставить enabled = true — больше ничего не требуется.

Ценность Nexar: один запрос отдаёт продавцов Verical, Rochester, Avnet, TTI,
Arrow, Future и сам Octopart — то есть четыре обязательных сайта из листа 2,
включая недостижимые с этой сети.
"""

from __future__ import annotations

import json
import time
import urllib.parse

from .. import net, store

_token_cache: dict[str, dict] = {}

_NEXAR_QUERY = """
query multi($queries: [SupPartMatchQuery!]!) {
  supMultiMatch(queries: $queries, currency: "USD") {
    parts {
      mpn
      manufacturer { name }
      octopartUrl
      sellers {
        company { name isDistributorApi }
        isAuthorized
        offers {
          sku inventoryLevel moq packaging clickUrl updated factoryLeadDays
          prices { quantity price currency }
        }
      }
    }
  }
}
"""


def _nexar_token() -> str | None:
    cid, secret = store.env("NEXAR_CLIENT_ID"), store.env("NEXAR_CLIENT_SECRET")
    if not (cid and secret):
        return None
    cur = _token_cache.get("nexar")
    if cur and cur["expires_at"] > time.time() + 30:
        return cur["token"]
    body = net.post_form("https://identity.nexar.com/connect/token", {
        "grant_type": "client_credentials", "client_id": cid,
        "client_secret": secret, "scope": "supply.domain"}, rps=1.0)
    data = json.loads(body.decode())
    _token_cache["nexar"] = {"token": data["access_token"],
                             "expires_at": time.time() + int(data.get("expires_in", 3000))}
    return data["access_token"]


def nexar(src, query: str):
    token = _nexar_token()
    if not token:
        return [], None
    payload = net.post_json(
        "https://api.nexar.com/graphql",
        {"query": _NEXAR_QUERY, "variables": {"queries": [{"mpn": query, "limit": 5}]}},
        headers={"Authorization": "Bearer " + token}, rps=src.rps)

    offers = []
    for group in (payload.get("data") or {}).get("supMultiMatch") or []:
        for part in group.get("parts") or []:
            mpn = part.get("mpn") or ""
            manu = (part.get("manufacturer") or {}).get("name") or ""
            for seller in part.get("sellers") or []:
                company = (seller.get("company") or {}).get("name") or ""
                kind = "franchise" if seller.get("isAuthorized") else "broker"
                for off in seller.get("offers") or []:
                    breaks = [{"qty": pr.get("quantity"), "price": pr.get("price"),
                               "currency": pr.get("currency") or "USD"}
                              for pr in off.get("prices") or []]
                    lead = off.get("factoryLeadDays")
                    offers.append({
                        "offer_mpn": mpn, "offer_pn": mpn,
                        "sku": off.get("sku") or "", "manufacturer": manu,
                        "seller": company, "seller_kind": kind,
                        "stock_qty": off.get("inventoryLevel"),
                        "moq": off.get("moq"), "packaging": off.get("packaging") or "",
                        "lead_time_days": lead,
                        "lead_time_raw": f"{lead} days" if lead else "",
                        "price_breaks": breaks,
                        "currency": (breaks[0]["currency"] if breaks else "USD"),
                        "source_url": off.get("clickUrl") or part.get("octopartUrl")
                        or src.deep_link(query),
                    })
    return offers, payload


def oemsecrets(src, query: str):
    key = store.env("OEMSECRETS_API_KEY")
    if not key:
        return [], None
    # суточную квоту (rate.daily) списывает оркестратор до вызова адаптера —
    # здесь второй раз не списываем, иначе каждый запрос стоил бы две единицы

    url = ("https://oemsecretsapi.com/partsearch?"
           + urllib.parse.urlencode({"searchTerm": query, "apiKey": key,
                                     "currency": "USD", "countryCode": "US"}))
    payload = net.get_json(url, rps=src.rps)

    offers = []
    for row in payload.get("stock") or []:
        dist = row.get("distributor") or {}
        prices = row.get("prices") or {}
        currency = "USD" if "USD" in prices else (next(iter(prices), "USD"))
        breaks = [{"qty": t.get("unit_break"), "price": t.get("unit_price"),
                   "currency": currency} for t in prices.get(currency) or []]
        offers.append({
            "offer_mpn": row.get("part_number") or "",
            "offer_pn": row.get("part_number") or "",
            "sku": row.get("sku") or "",
            "manufacturer": row.get("manufacturer") or "",
            "seller": dist.get("distributor_name") or "",
            "seller_kind": "broker",
            "stock_qty": row.get("quantity_in_stock"),
            "moq": row.get("moq"), "packaging": row.get("packaging") or "",
            "lead_time_raw": row.get("lead_time") or "",
            "date_code": str(row.get("date_code") or ""),
            "price_breaks": breaks, "currency": currency,
            "source_url": row.get("buy_now_url") or src.deep_link(query),
        })
    return offers, payload


def arrow(src, query: str):
    key, login = store.env("ARROW_API_KEY"), store.env("ARROW_LOGIN")
    if not (key and login):
        return [], None
    url = ("https://api.arrow.com/itemservice/v4/en/search/token?"
           + urllib.parse.urlencode({"api_key": key, "login": login,
                                     "search_token": query, "resources": "pricing"}))
    payload = net.get_json(url, rps=src.rps)

    offers = []
    for item_group in payload.get("itemserviceresult", {}).get("data") or []:
        for part in item_group.get("PartList") or []:
            for source in part.get("Sources") or []:
                for item in source.get("sourceParts") or []:
                    breaks = [{"qty": p.get("minQty"), "price": p.get("resaleprice"),
                               "currency": p.get("currency") or "USD"}
                              for p in item.get("Prices", {}).get("resale") or []]
                    avail = (item.get("Availability") or [{}])[0]
                    offers.append({
                        "offer_mpn": part.get("partNum") or "",
                        "offer_pn": part.get("partNum") or "",
                        "manufacturer": part.get("manufacturer") or "",
                        "seller": "Arrow", "seller_kind": "franchise",
                        "stock_qty": avail.get("fohQty"),
                        "lead_time_raw": avail.get("leadTime") or "",
                        "moq": item.get("minOrderQuantity"),
                        "price_breaks": breaks, "currency": "USD",
                        "source_url": item.get("buyUrl") or src.deep_link(query),
                    })
    return offers, payload
