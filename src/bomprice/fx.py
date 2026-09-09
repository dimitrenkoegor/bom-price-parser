"""Курсы валют для отчёта. Один набор курсов на прогон, кэш на сутки.

Захват предложений валюту не трогает — конвертация живёт только здесь.
Курс и дата попадают в подвал каждого листа, чтобы цифры можно было
воспроизвести через месяц.
"""

from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass, field

from . import net, store

CBR_URL = "https://www.cbr.ru/scripts/XML_daily.asp"
ERAPI_URL = "https://open.er-api.com/v6/latest/USD"


@dataclass
class Rates:
    usd_rub: float = 0.0          # рублей за доллар, уже с наценкой
    usd_rub_cbr: float = 0.0      # чистый курс ЦБ, для подписи
    markup: float = 0.0
    per_usd: dict = field(default_factory=dict)   # валюта → единиц за 1 USD
    source: str = ""
    at: str = ""

    def to_usd(self, amount: float | None, currency: str) -> float | None:
        if amount is None:
            return None
        code = (currency or "USD").upper()
        if code == "USD":
            return round(amount, 6)
        if code == "RUB" and self.usd_rub_cbr:
            return round(amount / self.usd_rub_cbr, 6)
        rate = self.per_usd.get(code)
        if not rate:
            return None
        return round(amount / rate, 6)

    def to_rub(self, amount: float | None, currency: str) -> float | None:
        if amount is None:
            return None
        if (currency or "USD").upper() == "RUB":
            return round(amount, 2)
        usd = self.to_usd(amount, currency)
        if usd is None or not self.usd_rub:
            return None
        return round(usd * self.usd_rub, 2)

    def footer(self) -> str:
        parts = []
        if self.usd_rub:
            parts.append(f"USD/RUB {self.usd_rub:.4f} "
                         f"(ЦБ {self.usd_rub_cbr:.4f} + наценка {self.markup:g})")
        for code in ("EUR", "GBP", "CNY"):
            if code in self.per_usd:
                parts.append(f"{code}/USD {1 / self.per_usd[code]:.4f}")
        return f"Курсы на {self.at or '—'}: " + "; ".join(parts) if parts else \
            "Курсы недоступны — колонки USD/RUB пустые"


def _cbr_usd_rub() -> float:
    # ЦБ отдаёт XML в windows-1251 и вставляет <NumCode> перед <CharCode>
    xml = net.fetch(CBR_URL, rps=1.0).decode("windows-1251", "replace")
    match = re.search(
        r"<CharCode>USD</CharCode>.*?<Nominal>(\d+)</Nominal>"
        r".*?<Value>([\d,\.]+)</Value>", xml, re.S)
    if not match:
        raise RuntimeError("в ответе ЦБ нет USD")
    nominal = float(match.group(1) or 1)
    value = float(match.group(2).replace(",", "."))
    return value / nominal


def load(rub_rate: float = 0.0, markup: float | None = None,
         refresh: bool = False) -> Rates:
    cfg = store.settings().get("currency", {})
    markup = cfg.get("rub_markup", 4.0) if markup is None else markup
    ttl = float(cfg.get("fx_ttl_hours", 24)) * 3600

    cached = None
    if store.FX_CACHE.exists() and not refresh:
        try:
            blob = json.loads(store.FX_CACHE.read_text(encoding="utf-8"))
            if time.time() - blob.get("at_ts", 0) < ttl:
                cached = blob
        except (json.JSONDecodeError, OSError):
            cached = None

    if cached:
        per_usd, cbr, at, source = (cached["per_usd"], cached["usd_rub_cbr"],
                                    cached["at"], cached["source"] + " (кэш)")
    else:
        per_usd, cbr, source = {}, 0.0, ""
        try:
            data = net.get_json(ERAPI_URL, rps=1.0)
            per_usd = {k: float(v) for k, v in (data.get("rates") or {}).items()
                       if isinstance(v, (int, float))}
            source = "open.er-api.com"
        except Exception as exc:                       # noqa: BLE001
            print(f"  курсы: er-api недоступен ({type(exc).__name__}) — "
                  f"колонки USD по неUSD-источникам останутся пустыми")
        try:
            cbr = _cbr_usd_rub()
            source = (source + " + ЦБ РФ").strip(" +")
        except Exception as exc:                       # noqa: BLE001
            print(f"  курсы: ЦБ недоступен ({type(exc).__name__}) — "
                  f"рублёвые колонки останутся пустыми")
            if per_usd.get("RUB"):
                cbr = per_usd["RUB"]
                source += " (RUB по er-api)"
        at = store.utcnow()[:10]
        if per_usd or cbr:
            store.ensure_dirs()
            store.FX_CACHE.write_text(json.dumps(
                {"at": at, "at_ts": time.time(), "per_usd": per_usd,
                 "usd_rub_cbr": cbr, "source": source}, ensure_ascii=False),
                encoding="utf-8")

    manual = float(rub_rate or cfg.get("rub_rate", 0) or 0)
    effective = manual if manual > 0 else (cbr + markup if cbr else 0.0)
    return Rates(usd_rub=effective, usd_rub_cbr=cbr or manual, markup=markup if not manual else 0.0,
                 per_usd=per_usd, source=source or "ручной курс", at=at)
