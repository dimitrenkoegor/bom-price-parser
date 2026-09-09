"""Единая запись предложения + разбор «сырых» значений источников.

Одна строка offers.jsonl = одно предложение = одна строка листа «Предложения».
Валюта здесь НЕ конвертируется: пересчёт живёт только в отчёте, иначе повторная
сборка завтра даст другие цифры без объяснения.
"""

from __future__ import annotations

import re

FIELDS = (
    "run_id", "fetched_at", "part_id", "req_pn", "req_qty", "req_manufacturer",
    "source_id", "source_title", "source_url", "seller", "seller_kind",
    "offer_pn", "offer_mpn", "sku", "manufacturer", "manufacturer_match",
    "match_quality", "match_delta", "variant",
    "stock_qty", "stock_note", "moq", "order_multiple", "packaging",
    "lead_time_days", "lead_time_raw", "ship_from",
    "date_code", "date_code_norm", "date_code_source",
    "price_breaks", "currency", "lifecycle", "rohs",
    "provenance", "raw_ref", "warnings",
)


def new(**kw) -> dict:
    """Создаёт запись предложения с полным набором полей."""
    rec = {name: kw.get(name) for name in FIELDS}
    rec["price_breaks"] = kw.get("price_breaks") or []
    rec["warnings"] = kw.get("warnings") or []
    rec["stock_qty"] = to_int(kw.get("stock_qty"), 0)
    rec["moq"] = to_int(kw.get("moq"), 1) or 1
    rec["order_multiple"] = to_int(kw.get("order_multiple"), 0)
    rec["currency"] = (kw.get("currency") or "USD").upper()
    dc = (kw.get("date_code") or "").strip()
    if not looks_like_date_code(dc):
        dc = ""
    rec["date_code"] = dc
    rec["date_code_norm"] = norm_date_code(dc)
    if not rec["date_code_source"]:
        rec["date_code_source"] = "source" if dc else "нет данных"
    return rec


def dc_source(src, date_code: str) -> str:
    """Откуда взялся DC. Пустой DC у франшизы, которая его не публикует
    (поле gives в реестре без "dc"), — свойство канала, а не сбой сбора."""
    if looks_like_date_code(date_code or ""):
        return "source"
    if src is not None and src.kind == "franchise" and not src.gives_dc():
        return "n/a-франшиза"
    return "нет данных"


# ── разбор значений ──────────────────────────────────────────────────────────

def to_int(value, default: int = 0) -> int:
    if value is None or value == "":
        return default
    if isinstance(value, bool):
        return default
    if isinstance(value, (int, float)):
        return int(value)
    digits = re.sub(r"[^\d]", "", str(value))
    return int(digits) if digits else default


def to_float(value, default: float | None = None) -> float | None:
    if value is None or value == "":
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    text = re.sub(r"[^\d.,\-]", "", text)
    if not text:
        return default
    # «1 234,56» и «1,234.56» — определяем разделитель по последнему знаку
    if "," in text and "." in text:
        text = text.replace(",", "") if text.rindex(".") > text.rindex(",") else \
            text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".") if len(text.split(",")[-1]) <= 2 else text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return default


_LEAD_PATTERNS = (
    (re.compile(r"(\d+)\s*(?:-\s*\d+\s*)?(?:раб\.?\s*)?дн", re.I), 1),
    (re.compile(r"(\d+)\s*(?:-\s*\d+\s*)?нед", re.I), 7),
    (re.compile(r"(\d+)\s*(?:-\s*\d+\s*)?мес", re.I), 30),
    (re.compile(r"(\d+)\s*(?:-\s*\d+\s*)?day", re.I), 1),
    (re.compile(r"(\d+)\s*(?:-\s*\d+\s*)?w(?:ee)?k", re.I), 7),
    (re.compile(r"(\d+)\s*(?:-\s*\d+\s*)?month", re.I), 30),
)


def lead_days(raw) -> int | None:
    """«12 Weeks», «3-5 дней», «10 wks» → дни. Неразобранное → None, сырое остаётся в отчёте."""
    if raw is None or raw == "":
        return None
    if isinstance(raw, (int, float)):
        return int(raw)
    text = str(raw)
    if re.search(r"in stock|со склада|в наличии|available", text, re.I):
        return 0
    for pattern, factor in _LEAD_PATTERNS:
        match = pattern.search(text)
        if match:
            return int(match.group(1)) * factor
    return None


_DC_YYWW = re.compile(r"^(\d{2})(\d{2})$")
_DC_YYYYWW = re.compile(r"^(20\d{2})(\d{2})$")
_DC_YYYYMMDD = re.compile(r"^(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])$")


def looks_like_date_code(raw: str) -> bool:
    """Отсеивает заглушки, которые площадки ставят вместо DC.

    Findchips, например, пишет в поле Date Code «0» или «1» у авторизованных
    дистрибьюторов — это не партия, а признак «не указано». Настоящий DC — это
    минимум две цифры: «1706», «12+», «20180423».
    """
    text = str(raw or "").strip()
    if not text or text.lower() in ("n/a", "na", "-", "—", "нет", "none"):
        return False
    return sum(ch.isdigit() for ch in text) >= 2


def norm_date_code(raw: str) -> str:
    """DC к виду ГГГГ-НН. Непонятное оставляем пустым, сырое всегда есть в колонке DC.

    Четыре цифры в date code — это ГГНН (год и неделя), а не год: «2149» = 49-я
    неделя 2021 года. Поэтому «2023» читается как 23-я неделя 2020-го, а не как
    год 2023 — таков отраслевой формат маркировки.
    """
    text = re.sub(r"[^0-9]", "", str(raw or ""))
    if not text:
        return ""
    match = _DC_YYYYMMDD.match(text)
    if match:
        import datetime
        year, month, day = (int(match.group(i)) for i in (1, 2, 3))
        try:
            return f"{year}-{datetime.date(year, month, day).isocalendar().week:02d}"
        except ValueError:
            return str(year)
    match = _DC_YYYYWW.match(text)
    if match and 1 <= int(match.group(2)) <= 53:
        return f"{match.group(1)}-{int(match.group(2)):02d}"
    match = _DC_YYWW.match(text)
    if match:
        year, week = int(match.group(1)), int(match.group(2))
        if 1 <= week <= 53:
            century = 2000 if year <= 79 else 1900
            return f"{century + year}-{week:02d}"
    return ""


def price_at(breaks: list[dict], qty: int) -> tuple[float | None, int, bool]:
    """→ (цена, количество ценового брекета, ниже ли минимальной партии).

    Берём наибольший брекет с qty ≤ запрошенного. Если все брекеты выше
    запроса — берём самый нижний и поднимаем флаг: платить придётся за партию.
    """
    valid = [b for b in (breaks or [])
             if b.get("price") is not None and to_int(b.get("qty"), 0) > 0]
    if not valid:
        return None, 0, False
    valid.sort(key=lambda b: to_int(b["qty"]))
    applicable = [b for b in valid if to_int(b["qty"]) <= qty]
    if applicable:
        chosen = applicable[-1]
        return float(chosen["price"]), to_int(chosen["qty"]), False
    chosen = valid[0]
    return float(chosen["price"]), to_int(chosen["qty"]), True


def clean_breaks(raw: list[dict], currency: str = "USD") -> list[dict]:
    out = []
    for item in raw or []:
        qty = to_int(item.get("qty"), 0)
        price = to_float(item.get("price"))
        if qty > 0 and price is not None and price > 0:
            out.append({"qty": qty, "price": round(price, 6),
                        "currency": (item.get("currency") or currency).upper()})
    out.sort(key=lambda b: b["qty"])
    return out
