"""Разбор колонки «наименование» в (артикул, производитель) + варианты написания.

От качества этого модуля зависит всё остальное: неверно разобранный артикул
отравляет запросы ко всем двадцати источникам сразу.

Правила и способы починки — docs/parsing.md. Словари — config/brands.toml.

Модуль общий для двух проектов (count pars и bom-price-parser): копия должна
совпадать байт в байт, правки вносить в обоих местах. Отличие проектов — не
здесь, а в том, что они делают с качеством совпадения: обзор рынка показывает
всё и помечает, котировщик принимает только exact/mask/suffix/prefix_series.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

from . import store

# Кириллические буквы, неотличимые от латинских. Только они дают право
# латинизировать токен. «Д», «Ж», «Б» и прочие — признак отечественного
# обозначения (2Д510А), такой артикул трогать нельзя.
HOMOGLYPHS = {
    "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H", "О": "O",
    "Р": "P", "С": "C", "Т": "T", "Х": "X", "У": "Y", "І": "I", "Ѕ": "S",
    "а": "a", "е": "e", "к": "k", "м": "m", "о": "o", "р": "p", "с": "c",
    "т": "t", "х": "x", "у": "y", "і": "i",
}

# Упаковочные хвосты — снимаются при построении «корня» артикула и считаются
# допустимым расхождением в match_quality («p/n может отличаться окончанием»).
# Объединение двух регламентов: T/TR/CT (лента), K (bulk), L/LF/ELF/GLF (RoHS,
# Bourns), PBF/TRPBF (ADI, Infineon), G/GT (Vishay), RL/R7 (Yageo), E4 (TI).
PACKAGING_TAILS = {
    "PBF", "TRPBF", "TRLPBF", "TRPBFT", "WPBF", "TR", "T&R", "TAPE", "REEL",
    "REEL7", "REEL13", "CUT", "BULK", "TRAY", "TUBE", "ND", "LF", "L", "T", "CT",
    "K", "ELF", "GLF", "RL", "MRL", "GE3", "GE4", "E2", "E4", "E6327", "G", "GT",
    "G4", "TE1", "T1", "T2", "T3", "T7", "TB", "TG", "RO", "RG", "RH", "13", "7",
    "1", "2", "3", "R7", "Q1", "AA", "ES",
    # TDK: C5750X7R1H106KT ≡ C5750X7R1H106K230KB (толщина + bulk) — из регламента котировщика
    "230KB", "230KA", "230KC", "160KA", "160KB", "280KA", "280KB",
    # TI/National «/NOPB» (без свинца), ADI «Z» (RoHS) и «-R2»/«ZRL» (катушка),
    # onsemi/Fairchild «X» (лента). Хвосты могут идти цепочкой: AD8027ART → AD8027ARTZ-R2
    "NOPB", "Z", "X", "R2", "ZRL",
}

_TAILS_BY_LEN = sorted(PACKAGING_TAILS, key=len, reverse=True)


def _consume_tails(delta: str) -> bool:
    """True, если хвост целиком складывается из упаковочных суффиксов: «ZREEL7», «XNOPB»."""
    rest = delta.upper()
    while rest:
        for tail in _TAILS_BY_LEN:
            if rest.startswith(tail):
                rest = rest[len(tail):]
                break
        else:
            return False
    return True

_UNIT = re.compile(
    r"^[\d.,]+\s*("
    r"пФ|нФ|мкФ|мФ|Ф|pF|nF|uF|µF|mF|"
    r"нГн|мкГн|мГн|Гн|nH|uH|µH|mH|"
    r"мОм|кОм|МОм|Ом|Ohm|"
    r"мкА|мА|кА|А|мкВ|мВ|кВ|В|"
    r"mA|uA|kA|mV|uV|kV|"
    r"мВт|кВт|Вт|mW|kW|W|"
    r"кГц|МГц|ГГц|Гц|kHz|MHz|GHz|Hz|"
    r"%|ppm|°C|C|V|A|F|H"
    r")$", re.IGNORECASE)

# Кириллическая единица, приклеенная к числу внутри токена: «200КОМ», «32,768КГЦ».
# Такой токен — параметр, а не артикул; при отсутствии артикула единица
# нормализуется в латиницу (КОМ → K), см. normalize_cyrillic_units.
_CYR_UNIT_INSIDE = re.compile(
    r"\d\s*(МКГН|НГН|МГН|МКФ|НФ|ПФ|МОМ|КОМ|ОМ|КГЦ|МГЦ|ГГЦ|ГЦ)(?![А-ЯЁа-яё])",
    re.IGNORECASE)

_CASE_SIZE = re.compile(
    r"^(0075|0100|0201|0402|0603|0805|1008|1206|1210|1218|1806|1812|2010|2220|2512|"
    r"SOT\d*|SOD\d*|SOIC\d*|MSOP\d*|TSSOP\d*|SSOP\d*|QFN\d*|QFP\d*|BGA\d*|DFN\d*|"
    r"DIP\d*|TO\d*|SMA|SMB|SMC|LQFP\d*|TQFP\d*)$", re.IGNORECASE)

_MIN_ALIAS_LEN = 3  # «AD», «TI», «WE» слишком коротки — ловят ложное внутри артикулов


@dataclass
class Part:
    part_id: int
    raw: str
    qty: int
    mpn: str = ""                 # латинизированный, идёт в запросы
    mpn_raw: str = ""             # как в исходнике, идёт в отчёт
    manufacturer: str = ""
    manufacturer_note: str = ""   # правопреемник из «/AD»
    description_type: str = ""    # снятое русское слово-тип
    description_params: str = ""  # содержимое скобок
    brand_source: str = "none"    # dict | tail | lead | column | none
    homoglyph_fixed: bool = False
    mpn_script: str = "latin"     # latin | cyrillic
    confidence: str = "high"      # high | medium | low
    warnings: list[str] = field(default_factory=list)
    variants: list[dict] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "part_id": self.part_id, "raw": self.raw, "qty": self.qty,
            "mpn": self.mpn, "mpn_raw": self.mpn_raw,
            "manufacturer": self.manufacturer, "manufacturer_note": self.manufacturer_note,
            "description_type": self.description_type,
            "description_params": self.description_params,
            "brand_source": self.brand_source, "homoglyph_fixed": self.homoglyph_fixed,
            "mpn_script": self.mpn_script, "confidence": self.confidence,
            "warnings": self.warnings, "variants": self.variants,
        }


# ── словари ──────────────────────────────────────────────────────────────────

_dicts: dict | None = None


def _load() -> dict:
    global _dicts
    if _dicts is not None:
        return _dicts
    cfg = store.load_toml("brands")
    pairs: list[tuple[str, str]] = []          # (написание, канон)
    for canon, aliases in cfg.get("brands", {}).items():
        pairs.append((canon, canon))
        for alias in aliases:
            if len(alias) >= _MIN_ALIAS_LEN or not alias.isalnum():
                pairs.append((alias, canon))
    pairs.sort(key=lambda p: len(p[0]), reverse=True)
    _dicts = {
        "brand_pairs": pairs,
        "brand_canon": {_brand_key(spelling): canon for spelling, canon in pairs},
        "ru_types": sorted(cfg.get("ru_type_words", []), key=len, reverse=True),
        "stop_words": {w.upper() for w in cfg.get("stop_words", [])},
        "acquired_by": cfg.get("acquired_by", {}),
        "slash_notes": cfg.get("slash_notes", {}),
    }
    return _dicts


def _brand_key(name: str) -> str:
    return re.sub(r"[^a-z0-9а-яё]", "", (name or "").lower())


def canonical(text: str) -> str:
    """Схлопнутая форма для сравнения: только буквы и цифры, верхний регистр.

    Здесь гомоглифы гасятся безусловно — сравнение от этого только выигрывает,
    а исходное написание артикула хранится отдельно и не страдает.
    """
    return re.sub(r"[^A-Z0-9]", "", force_latin(text or "").upper())


# ── стадии разбора ───────────────────────────────────────────────────────────

def normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    text = text.replace(" ", " ").replace(" ", " ").replace(" ", " ")
    for dash in "‐‑‒–—―−":
        text = text.replace(dash, "-")
    return re.sub(r"\s+", " ", text).strip()


def force_latin(token: str) -> str:
    """Безусловная замена гомоглифов. Только для сравнения и для отдельного варианта поиска."""
    return "".join(HOMOGLYPHS.get(ch, ch) for ch in token)


def fix_homoglyphs(token: str) -> str:
    """Латинизирует токен, только если это заведомо безопасно.

    Требуется ХОТЯ БЫ ОДНА латинская буква плюс все кириллические символы —
    визуальные двойники. Тогда `М83513/01-AN` → `M83513/01-AN`.

    Одних цифр мало: `КР580ВМ80А` — настоящий отечественный артикул, целиком
    набранный кириллицей, и латинизировать его нельзя. Для таких случаев
    build_variants добавляет отдельный латинизированный вариант поиска,
    не трогая исходное написание.
    """
    cyr = [ch for ch in token if "Ѐ" <= ch <= "ӿ"]
    if not cyr:
        return token
    if not any("a" <= ch <= "z" or "A" <= ch <= "Z" for ch in token):
        return token
    if not all(ch in HOMOGLYPHS for ch in cyr):
        return token
    return force_latin(token)


def latinizable(token: str) -> bool:
    """Кириллический токен, который целиком состоит из двойников латиницы."""
    cyr = [ch for ch in token if "Ѐ" <= ch <= "ӿ"]
    return bool(cyr) and all(ch in HOMOGLYPHS for ch in cyr)


def has_cyrillic(text: str) -> bool:
    return any("Ѐ" <= ch <= "ӿ" for ch in text)


def normalize_cyrillic_units(text: str) -> str:
    """«200КОМ» → «200K», «32,768КГЦ» → «32,768kHz», «10ОМ» → «10R».

    Меняются только единицы после числа. Прочая кириллица не трогается —
    отечественные обозначения (2Д510А) остаются как есть.
    """
    value = str(text or "")
    for pattern, replacement in (
        (r"(?<=\d)\s*МКГН(?![А-ЯЁа-яё])", "uH"), (r"(?<=\d)\s*НГН(?![А-ЯЁа-яё])", "nH"),
        (r"(?<=\d)\s*МГН(?![А-ЯЁа-яё])", "mH"),
        (r"(?<=\d)\s*МКФ(?![А-ЯЁа-яё])", "uF"), (r"(?<=\d)\s*НФ(?![А-ЯЁа-яё])", "nF"),
        (r"(?<=\d)\s*ПФ(?![А-ЯЁа-яё])", "pF"),
        (r"(?<=\d)\s*МОМ(?![А-ЯЁа-яё])", "M"), (r"(?<=\d)\s*КОМ(?![А-ЯЁа-яё])", "K"),
        (r"(?<=\d)\s*ОМ(?![А-ЯЁа-яё])", "R"),
        (r"(?<=\d)\s*ГГЦ(?![А-ЯЁа-яё])", "GHz"), (r"(?<=\d)\s*МГЦ(?![А-ЯЁа-яё])", "MHz"),
        (r"(?<=\d)\s*КГЦ(?![А-ЯЁа-яё])", "kHz"), (r"(?<=\d)\s*ГЦ(?![А-ЯЁа-яё])", "Hz"),
    ):
        value = re.sub(pattern, replacement, value, flags=re.IGNORECASE)
    return value


def strip_parens(text: str) -> tuple[str, str]:
    """Убирает скобочные блоки, ОТДЕЛЁННЫЕ ПРОБЕЛОМ.

    Пробел обязателен: в `FX11LA-80P/8-SV(71)` скобка — часть артикула Hirose,
    а в `GRM21BR71H105KA12 (0805-X7R...)` — параметры. Разница только в пробеле.
    """
    grabbed: list[str] = []

    def take(match: re.Match) -> str:
        grabbed.append(match.group(1).strip())
        return " "

    text = re.sub(r"\s+\(([^()]*)\)", take, text)
    text = re.sub(r"\s+\[([^\[\]]*)\]", take, text)
    return re.sub(r"\s+", " ", text).strip(), "; ".join(grabbed)


def extract_brand(text: str) -> tuple[str, str, str, str]:
    """→ (остаток, канон бренда, примечание-правопреемник, источник совпадения)."""
    d = _load()
    note = ""

    # «Linear Technology /AD» — пометка правопреемника после слэша
    tail = re.search(r"\s/\s*([A-Za-z]{2,4})\s*$", text)
    if tail:
        key = tail.group(1).upper()
        if key in d["slash_notes"]:
            note = d["slash_notes"][key]
            text = text[: tail.start()].strip()

    for spelling, canon in d["brand_pairs"]:
        pattern = re.compile(r"(?<![A-Za-z0-9])" + re.escape(spelling) + r"(?![A-Za-z0-9])",
                             re.IGNORECASE)
        match = pattern.search(text)
        if not match:
            continue
        rest = (text[: match.start()] + " " + text[match.end():]).strip()
        rest = re.sub(r"\s+", " ", rest)
        if not rest:            # бренд — единственное, что есть: не забираем
            continue
        return rest, canon, note, "dict"
    return text, "", note, "none"


def canon_brand(name: str) -> str:
    """Написание бренда → канон из словаря; неизвестное — как есть (обрезанные пробелы)."""
    text = re.sub(r"\s+", " ", str(name or "")).strip()
    if not text:
        return ""
    return _load()["brand_canon"].get(_brand_key(text), text)


def strip_ru_type(text: str) -> tuple[str, str]:
    """Снимает русское слово-тип изделия из НАЧАЛА строки."""
    d = _load()
    low = text.lower()
    for word in d["ru_types"]:
        if low.startswith(word):
            rest = text[len(word):].strip(" -–—:,")
            if rest:
                return rest, text[: len(word)]
    # общий случай: ведущая цепочка кириллических слов, если дальше есть цифры
    match = re.match(r"^((?:[А-Яа-яЁё][А-Яа-яЁё\-]*\s+){1,3})(?=\S*\d)", text)
    if match and re.search(r"\d", text[match.end():]):
        return text[match.end():].strip(), match.group(1).strip()
    return text, ""


def _score_token(token: str) -> int:
    d = _load()
    up = token.upper()
    score = 0
    digits = sum(ch.isdigit() for ch in token)
    latin = sum(("A" <= ch <= "Z") or ("a" <= ch <= "z") for ch in token)

    if digits >= 2 and latin >= 1:
        score += 3
    if len(token) >= 5:
        score += 2
    if re.search(r"[A-Za-z0-9][-#/][A-Za-z0-9]", token):
        score += 2
    if token.isdigit():
        # Wurth и подобные используют чисто цифровые артикулы (744230900),
        # поэтому длинные числа штрафуем слабо, а короткие — как параметр.
        score += -1 if len(token) >= 6 else -5
    if _UNIT.fullmatch(token):
        score -= 5
    if _CASE_SIZE.fullmatch(token):
        score -= 5
    if _CYR_UNIT_INSIDE.search(token):
        # «3310Y-001-10КОМ» рядом с «3310Y001-103L»: параметр проигрывает артикулу
        score -= 4
    if up in d["stop_words"]:
        score -= 3
    return score


def _tokenize(text: str) -> list[str]:
    """Дробим только по пробелам: «,», «/», «#», «.», «-» живут внутри артикулов."""
    tokens = [t for t in text.split(" ") if t]
    merged: list[str] = []
    i = 0
    while i < len(tokens):
        cur = tokens[i]
        nxt = tokens[i + 1] if i + 1 < len(tokens) else ""
        # «BAS16H, 115» → «BAS16H,115» (упаковочный код NXP/Nexperia)
        if cur.endswith(",") and re.fullmatch(r"[A-Za-z0-9]{1,4}", nxt):
            merged.append(cur + nxt)
            i += 2
            continue
        # одиночный «Q»/«O» перед числом — часть артикула (кварцы Jauch: Q 10,240-…)
        if re.fullmatch(r"[QO]", cur) and nxt[:1].isdigit():
            merged.append(cur + nxt)
            i += 2
            continue
        # «16 MMCX-50-2-13/111 OE» (Huber+Suhner): число перед кодом со слэшем
        # и 1-3 заглавные буквы после него — части одного обозначения
        if re.fullmatch(r"\d{1,3}", cur) and "/" in nxt and re.search(r"[A-Za-z]", nxt):
            combined = cur + " " + nxt
            after = tokens[i + 2] if i + 2 < len(tokens) else ""
            if re.fullmatch(r"[A-Z]{1,3}", after):
                combined += " " + after
                i += 1
            merged.append(combined)
            i += 2
            continue
        merged.append(cur.strip(",;"))
        i += 1
    return [t for t in merged if t]


def pick_mpn(text: str) -> tuple[str, int, list[str]]:
    """→ (артикул, балл, прочие токены). Победитель — максимум балла, при ничьей левый."""
    tokens = _tokenize(text)
    if not tokens:
        return "", -99, []
    scored = [(_score_token(t), -i, t) for i, t in enumerate(tokens)]
    scored.sort(reverse=True)
    best_score, _, best = scored[0]
    if best_score <= 0:
        # ничего не похоже на артикул — берём самый длинный токен
        best = max(tokens, key=len)
        best_score = _score_token(best)
    return best, best_score, [t for t in tokens if t != best]


def parse_line(raw: str, qty: int, part_id: int = 0) -> Part:
    part = Part(part_id=part_id, raw=raw, qty=qty)
    text = normalize_text(raw)
    if not text:
        part.confidence = "low"
        part.warnings.append("пустая строка")
        return part

    text, params = strip_parens(text)
    part.description_params = params

    text, brand, note, brand_src = extract_brand(text)
    part.manufacturer, part.manufacturer_note, part.brand_source = brand, note, brand_src
    if not part.manufacturer and params:
        # «Вилка 44914-0801 (Molex)»: бренд в скобках — это бренд, а не параметр
        brand = _brand_in_parens(params)
        if brand:
            part.manufacturer, part.brand_source = brand, "parens"

    text, ru_type = strip_ru_type(text)
    part.description_type = ru_type

    if not part.manufacturer:
        text, brand, note, brand_src = _brand_heuristics(text)
        if brand:
            part.manufacturer = brand
            part.manufacturer_note = part.manufacturer_note or note
            part.brand_source = brand_src

    best, score, others = pick_mpn(text)
    part.mpn_raw = best
    fixed = fix_homoglyphs(best)
    if has_cyrillic(fixed) and _CYR_UNIT_INSIDE.search(fixed):
        # артикула нет, есть параметр с кириллической единицей: «0603-200КОМ+-5%»
        units = normalize_cyrillic_units(fixed)
        if not has_cyrillic(units):
            fixed = units
            part.warnings.append("единицы приведены к латинице, это параметр, а не артикул")
    part.homoglyph_fixed = fixed != best
    part.mpn = fixed
    part.mpn_script = "cyrillic" if has_cyrillic(fixed) else "latin"

    if part.mpn_script == "cyrillic":
        part.warnings.append("кириллический артикул — только российские источники")
    if score <= 0:
        part.confidence = "low"
        part.warnings.append("ни один токен не похож на артикул")
    elif len(part.mpn) < 4:
        part.confidence = "low"
        part.warnings.append("слишком короткий артикул")
    elif not part.manufacturer and len([t for t in others if t.isalpha()]) >= 2:
        part.confidence = "medium"
        part.warnings.append("производитель не распознан")
    elif not part.manufacturer:
        part.confidence = "medium"

    part.variants = build_variants(part.mpn)
    return part


def _brand_in_parens(params: str) -> str:
    """Скобочный блок, целиком являющийся названием бренда: по словарю либо
    1-3 латинских слова без цифр и единиц («Johnson Components»)."""
    d = _load()
    for chunk in params.split("; "):
        chunk = chunk.strip(" ,;")
        if not chunk or any(ch.isdigit() for ch in chunk) or "%" in chunk:
            continue
        canon = canon_brand(chunk)
        if canon != chunk or _brand_key(chunk) in d["brand_canon"]:
            return canon
        words = chunk.split()
        if 1 <= len(words) <= 3 and all(
                re.fullmatch(r"[A-Za-z][A-Za-z&+.\-]*", w) and w.upper() not in d["stop_words"]
                and not _CASE_SIZE.fullmatch(w) for w in words):
            return chunk
    return ""


def _brand_heuristics(text: str) -> tuple[str, str, str, str]:
    """Бренда нет в словаре: пробуем хвостовые и головные словесные токены."""
    d = _load()
    tokens = _tokenize(text)
    if len(tokens) < 2:
        return text, "", "", "none"

    def wordish(tok: str) -> bool:
        letters = sum(ch.isalpha() for ch in tok)
        return (letters >= 2 and letters > sum(ch.isdigit() for ch in tok)
                and tok.upper() not in d["stop_words"] and not _UNIT.fullmatch(tok))

    best_idx = tokens.index(pick_mpn(text)[0]) if pick_mpn(text)[0] in tokens else 0

    tail = tokens[best_idx + 1:]
    if tail and len(tail) <= 3 and all(wordish(t) for t in tail):
        return " ".join(tokens[: best_idx + 1]), " ".join(tail), "", "tail"

    lead = tokens[:best_idx]
    if lead and len(lead) <= 3 and all(wordish(t) for t in lead):
        return " ".join(tokens[best_idx:]), " ".join(lead), "", "lead"

    return text, "", "", "none"


# ── маска регламента: x в конце и в середине артикула ────────────────────────

def split_mask(requested: str) -> tuple[str, bool]:
    """Маска «x / Xx / XX в конце — любой суффикс подходит».

    - строчный x (или кириллическая х), один и более — маска (3313J-1-104x);
    - хвост из двух и более заглавных X — маска (DLC…CXX);
    - смешанный хвост (Xx) — маска (DLC70B8R2CW501Xx);
    - одиночная заглавная X — часть реального артикула (EXB28V220JX).
    База короче 3 значимых символов маской не считается.
    → (база без маски, была ли маска).
    """
    raw = str(requested or "").strip()
    match = re.search(r"[xхXХ]+$", raw)
    if match and match.start() > 0:
        tail = match.group(0)
        has_lower = any(ch in "xх" for ch in tail)
        uppers = sum(1 for ch in tail if ch in "XХ")
        if has_lower or uppers >= 2:
            base = raw[: match.start()].rstrip("-. ")
            if len(canonical(base)) >= 3:
                return base, True
    return raw, False


def mask_regex(requested: str):
    """Строчный x в СЕРЕДИНЕ артикула = «любой один символ» (TPSE477K010x0200).

    → анкерованный regex по канонической форме кандидата или None, если внутренних x нет.
    """
    base, masked_tail = split_mask(str(requested or "").strip())
    core = force_latin(base)
    if "x" not in core:
        return None
    pattern = []
    for ch in core:
        if ch == "x":
            pattern.append("[A-Z0-9]")
        elif ch.isalnum():
            pattern.append(re.escape(ch.upper()))
        # разделители опускаются — сравнение идёт с канонической формой кандидата
    if not pattern:
        return None
    return re.compile("^" + "".join(pattern) + (".*" if masked_tail else "") + "$")


# ── варианты написания ───────────────────────────────────────────────────────

def _strip_one_tail(mpn: str) -> str | None:
    """Снимает ровно один упаковочный хвост. None — снимать больше нечего."""
    if "#" in mpn[1:]:
        head = mpn[: mpn.index("#", 1)]
        if len(head) >= 4:
            return head
    match = re.search(r",[A-Za-z0-9]{1,4}$", mpn)
    if match and len(mpn[: match.start()]) >= 4:
        return mpn[: match.start()]
    if mpn.endswith("+") and len(mpn) >= 5:
        return mpn[:-1]
    match = re.search(r"[-/]([A-Za-z0-9&]{1,6})$", mpn)
    if match and match.group(1).upper() in PACKAGING_TAILS and len(mpn[: match.start()]) >= 4:
        return mpn[: match.start()]
    return None


def root_of(mpn: str) -> str:
    """Итеративно снимает все упаковочные хвосты: SI1013CX-T1-GE3 → SI1013CX."""
    cur = mpn
    for _ in range(4):
        nxt = _strip_one_tail(cur)
        if not nxt:
            break
        cur = nxt
    return cur


def build_variants(mpn: str) -> list[dict]:
    """Упорядоченный список написаний для запросов. Первое — всегда исходное."""
    if not mpn:
        return []
    limit = int(store.settings().get("run", {}).get("max_variants", 3))
    out: list[dict] = [{"id": "exact", "mpn": mpn}]

    def add(vid: str, value: str) -> None:
        if value and value != mpn and all(v["mpn"] != value for v in out):
            out.append({"id": vid, "mpn": value})

    # маска x: площадке задаём базу — она вернёт все суффиксы, маска отберёт
    base, masked = split_mask(mpn)
    if masked:
        add("base", base)
    add("root", root_of(base if masked else mpn))
    # артикул набран кириллицей, но все буквы — двойники латиницы: ищем и так, и так
    if has_cyrillic(mpn) and latinizable(mpn):
        add("latin", force_latin(mpn))
    if "," in mpn:
        add("nocomma", mpn.split(",")[0])
    # хвостовой R/T = катушка/лента у TI, ADI, ON: TPS62065DSGR → TPS62065DSG
    if len(mpn) >= 8 and mpn[-1] in "RT" and mpn[-2].isalpha():
        add("reelless", mpn[:-1])
    return out[:limit] if limit else out


# ── качество совпадения предложенного артикула с запрошенным ─────────────────

QUALITIES = ("exact", "mask", "suffix", "prefix_series", "partial", "mismatch")


def match_quality(requested: str, offered: str, strict: bool = False) -> tuple[str, str]:
    """→ (exact | mask | suffix | prefix_series | partial | mismatch, что именно отличается).

    strict=False (обзор рынка): короткий незнакомый хвост (≤4 знаков) считается suffix.
    strict=True (котировщик): suffix только для хвостов из PACKAGING_TAILS, остальное partial.

    exact          — совпадение с точностью до разделителей и регистра;
    mask           — запрос с маской x/Xx/XX, кандидат подходит под неё;
    suffix         — отличие только в упаковочном хвосте;
    prefix_series  — кандидат дополнен буквенным префиксом серии (0805-FX-1503 → CR0805-FX-1503ELF);
    partial        — один артикул содержится в другом, но расхождение больше упаковочного;
    mismatch       — другой компонент.
    """
    if not requested or not offered:
        return "mismatch", ""
    a, b = canonical(requested), canonical(offered)
    if not a or not b:
        return "mismatch", ""
    if a == b:
        return "exact", ""

    # маска: точечный wildcard в середине и/или «любой суффикс» в конце
    rx = mask_regex(requested)
    if rx and rx.match(b):
        return "mask", ""
    base, masked = split_mask(requested)
    if masked:
        base_c = canonical(base)
        if base_c and b.startswith(base_c):
            return "mask", b[len(base_c):]

    if canonical(root_of(requested)) == canonical(root_of(offered)):
        return "suffix", _delta(requested, offered)
    if b.startswith(a):
        delta = offered[len(requested):] if offered.upper().startswith(requested.upper()) else b[len(a):]
        clean = re.sub(r"[^A-Z0-9]", "", delta.upper())
        if clean in PACKAGING_TAILS or _consume_tails(clean) or (not strict and len(clean) <= 4):
            return "suffix", delta
        return "partial", delta
    # у запроса свой упаковочный хвост, у кандидата — другой: C5750X7R1H106KT ↔ …106K230KB,
    # TPS62065DSGR ↔ TPS62065DSGT. Сравниваем по базе без хвоста / reel-маркера.
    for base in _bases(requested):
        if base and len(base) >= 5 and b.startswith(base) and b != base:
            tail = b[len(base):]
            if tail in PACKAGING_TAILS or _consume_tails(tail) or (not strict and len(tail) <= 4):
                return "suffix", tail
    if a.startswith(b):
        rem = a[len(b):]
        if rem in PACKAGING_TAILS or not strict:
            return "suffix", "-" + rem
        return "partial", "-" + rem

    # достройка серии-префикса: запрошено 0805-FX-1503, предложено CR0805-FX-1503ELF.
    # Требования: запрошенный код ≥ 6 знаков, префикс 1-2 буквы, хвост пуст
    # или упаковочный. Иначе LM358 «совпал» бы с PLM358, а 104 — с CR104.
    series = _series_prefix(a, b)
    if series:
        return "prefix_series", series
    series = _series_prefix(b, a)
    if series:
        return "prefix_series", series

    if a in b or b in a:
        return "partial", _delta(requested, offered)
    return "mismatch", ""


def _bases(mpn: str) -> list[str]:
    """Канонические базы артикула: без упаковочных хвостов и без хвостового R/T."""
    out = []
    for cand in (root_of(mpn), mpn[:-1] if len(mpn) >= 8 and mpn[-1] in "RTrt" and mpn[-2].isalpha() else ""):
        c = canonical(cand)
        if c and c not in out:
            out.append(c)
    return out


def _series_prefix(short: str, long: str) -> str:
    """Префикс серии, если long = <1-2 буквы> + short + <упаковочный хвост или ничего>."""
    if len(short) < 6:
        return ""
    idx = long.find(short)
    if not (0 < idx <= 2) or not long[:idx].isalpha():
        return ""
    tail = long[idx + len(short):]
    if tail == "" or tail in PACKAGING_TAILS or _consume_tails(tail):
        return long[:idx]
    return ""


def _delta(a: str, b: str) -> str:
    return b[len(a):] if b.upper().startswith(a.upper()) else b


def _brand_keys(name: str) -> set[str]:
    """Все ключи, под которыми бренд может встретиться: само имя, канон, его алиасы,
    правопреемники и поглощённые (с их алиасами)."""
    d = _load()
    keys = {_brand_key(name)}
    canon = canon_brand(name)
    if _brand_key(canon) in d["brand_canon"]:
        keys.add(_brand_key(canon))
        keys.update(k for k, c in d["brand_canon"].items() if c == canon)
    for old, new in d["acquired_by"].items():
        old_k, new_k = _brand_key(old), _brand_key(new)
        if old_k in keys or new_k in keys:
            for extra in (old, new):
                keys.add(_brand_key(extra))
                keys.update(k for k, c in d["brand_canon"].items() if c == extra)
    keys.discard("")
    return keys


def _close_enough(a: str, b: str) -> bool:
    """Одна опечатка в имени бренда длиной от 6 знаков: «Sypress» ≡ «Cypress»,
    «SKYWOKS» ≡ «Skyworks». Короткие имена сравниваются только точно."""
    if a == b:
        return True
    if min(len(a), len(b)) < 6 or abs(len(a) - len(b)) > 1:
        return False
    if len(a) == len(b):
        return sum(x != y for x, y in zip(a, b)) == 1
    short, long = (a, b) if len(a) < len(b) else (b, a)
    for i in range(len(long)):
        if long[:i] + long[i + 1:] == short:
            return True
    return False


def manufacturer_match(requested: str, offered: str) -> str:
    """yes | no | unknown. Учитывает словарь брендов, поглощения и одну опечатку:
    «Avago Technologies» ≡ «Broadcom Limited», «Linear Tehnology» ≡ «Analog Devices».
    """
    if not requested or not offered:
        return "unknown"
    left, right = _brand_keys(requested), _brand_keys(offered)
    if left & right:
        return "yes"
    for x in left:
        for y in right:
            if x and y and (x.startswith(y) or y.startswith(x) or _close_enough(x, y)):
                return "yes"
    return "no"
