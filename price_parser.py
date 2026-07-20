"""
Парсер цен на ЭКБ — версия 4 (API вместо браузерного парсинга)
================================================================
Поток работы:
  1. Берёт Excel-запрос — сам находит строку-шапку и колонки
     (Наименование / Условное обозначение / Количество / Производитель),
     вытаскивает артикул из описания.
  2. Пишет ПРЕВЬЮ  превью_позиции.xlsx  и показывает распознанное в консоли.
     Вы проверяете/правите его в Excel, сохраняете — и скрипт продолжит уже по
     вашим правкам.
  3. Ищет цены через официальные API дистрибьюторов:
       - DigiKey Product Information API v4 (OAuth2 client_credentials)
       - Newark/Farnell (element14) Product Search API
       - Mouser Search API (включится автоматически, когда придёт одобрение
         от Mouser — сейчас статус "pending authorisation")
     Выбирает минимальную цену при ценовом брекете ≥ количества, среди
     авторизованных дистрибьюторов.
  4. Пишет результат  final/BOM_Приложение_1.xlsx  — 10 колонок:
     Part Number | Distributor Part Number | Qty | Manufacturer | Distributor |
     Minimum Order | Stock | Lead Time | Unit Price USD | Unit Price RUB
     Рублёвая цена = USD * (курс ЦБ USD/RUB на момент запуска + наценка, по умолч. +4).

Ключи API берутся из файла .env рядом со скриптом (не коммитится в git):
    MOUSER_API_KEY=...
    DIGIKEY_CLIENT_ID=...
    DIGIKEY_CLIENT_SECRET=...
    FARNELL_API_KEY=...

Установка зависимостей:
    pip install openpyxl

Запуск:
    python price_parser.py                      # авто-выбор Excel-запроса рядом
    python price_parser.py --input "149 торги.xlsx"
    python price_parser.py --preview-only        # только распознать в превью, без цен
    python price_parser.py --yes                  # не спрашивать подтверждение
    python price_parser.py --once C5750X7R1H106KT 320   # проверка одной позиции
    python price_parser.py --rub-markup 4 --rate 1.16
"""

import argparse
import json
import re
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Установите зависимости:\n  pip install openpyxl")
    sys.exit(1)

SCRIPT_DIR = Path(__file__).resolve().parent
PREVIEW_NAME = "превью_позиции.xlsx"
ENV_PATH = SCRIPT_DIR / ".env"

AUTHORIZED_NAMES = {
    "digikey": "DigiKey",
    "mouser": "Mouser",
    "newark": "Newark",
    "farnell": "Farnell",
    "element14": "Newark/Farnell",
}

# Русские «типовые» слова в начале наименования — отбрасываем при извлечении артикула
RU_TYPE_WORDS = [
    "микросхема", "модуль питания", "модуль", "трансформатор", "конденсатор",
    "резистор", "диод", "транзистор", "разъём", "разъем", "дроссель", "реле",
    "стабилизатор", "преобразователь", "фильтр", "индуктивность", "катушка",
    "вилка", "розетка", "переключатель", "кварц", "резонатор", "предохранитель",
    "генератор", "кнопка", "термистор", "оптопара", "оптореле", "датчик",
]
# Известные производители — отбрасываем из хвоста наименования (и пишем в Manufacturer).
# Список нужен, когда колонка «Производитель» в запросе пустая, а бренд указан
# в конце описания («Микросхема LM8272MM/NOPB Texas Instruments»).
KNOWN_MFR = [
    # пассивы, разъёмы, кварцы
    "Traco Power", "TracoPower", "Traco", "Murata", "Vishay", "TDK", "Bourns",
    "Panasonic", "Rohm", "Littelfuse", "Yageo", "Kemet", "AVX", "Nichicon",
    "Wurth", "Würth", "TE Connectivity", "Molex", "Samsung", "Kyocera",
    "Knowles", "Voltronics", "Dalicap", "Hitano", "Jamicon", "Epcos", "Jauch",
    "Epson", "Golledge", "Avago Technologies", "Broadcom", "Axicom", "Cosmo",
    "Susumu", "Meggitt Electronics", "Fujitsu", "Switronic", "Amphenol",
    "Harting", "JST", "Hirose", "Samtec", "Adam Tech", "ADAM TECH", "Phoenix Contact",
    "Weidmuller", "Wago", "Abracon", "Raltron", "IQD", "NDK", "Taiyo Yuden",
    "Coilcraft", "Pulse Electronics", "Laird", "Chilisin", "Sunlord", "Walsin",
    "Johanson", "Skyworks", "Qorvo", "Mini-Circuits", "Delta Electronics",
    "Mean Well", "Recom", "XP Power", "Cui", "CUI", "Aimtec",
    # полупроводники (главный пробел, из-за которого запрос 282 ушёл в RFQ)
    "Texas Instruments", "Analog Devices", "Linear Technology", "Maxim Integrated",
    "Maxim", "International Rectifier", "Infineon Technologies", "Infineon",
    "STMicroelectronics", "ST Microelectronics", "NXP Semiconductors", "NXP",
    "ON Semiconductor", "onsemi", "Microchip Technology", "Microchip", "Atmel",
    "Nexperia", "Diodes Incorporated", "Diodes Inc", "Renesas Electronics",
    "Renesas", "Intersil", "Cypress Semiconductor", "Cypress", "Spansion",
    "Toshiba", "Nuvoton", "Winbond", "Micron Technology", "Micron", "Alliance Memory",
    "ISSI", "Lattice Semiconductor", "Lattice", "AMD", "Xilinx", "Altera", "Intel",
    "Marvell", "Mitel", "Microsemi", "Semtech", "Power Integrations", "Vicor",
    "Monolithic Power Systems", "MPS", "Silicon Labs", "Silicon Laboratories",
    "Melexis", "Allegro MicroSystems", "Allegro", "Bosch Sensortec", "Sensirion",
    "Honeywell", "Omron", "Panasonic Industrial", "Littelfuse Semiconductor",
    "Central Semiconductor", "Comchip", "Taiwan Semiconductor", "MCC",
    "Micro Commercial", "Wolfspeed", "Cree", "IXYS", "Sanken", "Rohm Semiconductor",
]

CYRILLIC_LOOKALIKES = str.maketrans({
    "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H",
    "О": "O", "Р": "P", "С": "C", "Т": "T", "Х": "X",
    "а": "a", "е": "e", "к": "k", "м": "m", "о": "o", "р": "p",
    "с": "c", "т": "t", "х": "x",
})

MANUFACTURER_ALIASES = {
    "avagotechnologies": "broadcomavago",
    "avago": "broadcomavago",
    "broadcom": "broadcomavago",
    "broadcomavago": "broadcomavago",
    "teconnectfivity": "teconnectivityaxicom",
    "teconnectivity": "teconnectivityaxicom",
    "axicom": "teconnectivityaxicom",
    "teconnectivityaxicom": "teconnectivityaxicom",
    "epcos": "tdkepcos",
    "tdk": "tdkepcos",
    "tdkepcos": "tdkepcos",
    # Fujitsu продала релейный бизнес — дистрибьюторы значат его как FCL Components
    "fujitsu": "fujitsufcl",
    "fcl": "fujitsufcl",
    "fclcomponents": "fujitsufcl",
}

# Поглощения: бренд в запросе и бренд у дистрибьютора — одна и та же деталь.
# Ключ — canonical_manufacturer() запрошенного, значение — итоговый канон.
ACQUISITION_ALIASES = {
    # Infineon: International Rectifier (2015), Cypress (2020)
    ("internationalrectifier", "ir", "irf", "infineon", "infineontechnologies",
     "cypress", "cypresssemiconductor", "spansion"): "infineon",
    # Analog Devices: Linear Technology (2017), Maxim Integrated (2021)
    ("analogdevices", "analogdevicesinc", "lineartechnology", "lineartech", "ltc",
     "maxim", "maximintegrated", "maximintegratedproducts"): "analogdevices",
    # Renesas: Intersil (2017), IDT (2019)
    ("renesas", "renesaselectronics", "intersil", "idt",
     "integrateddevicetechnology"): "renesas",
    # Microchip: Atmel (2016), Microsemi (2018), Micrel
    ("microchip", "microchiptechnology", "atmel", "microsemi", "micrel",
     "mitel"): "microchip",
    # NXP: Freescale (2015)
    ("nxp", "nxpsemiconductors", "freescale", "freescalesemiconductor"): "nxp",
    # onsemi: Fairchild (2016)
    ("onsemiconductor", "onsemi", "on", "fairchild",
     "fairchildsemiconductor"): "onsemi",
    # AMD: Xilinx (2022)
    ("amd", "xilinx", "amdxilinx"): "amdxilinx",
    # Skyworks: Silicon Labs infrastructure (2021)
    ("skyworks", "skyworkssolutions", "siliconlabs", "siliconlaboratories"): "skyworks",
}


# ─────────────────────────────────────────────
# .ENV
# ─────────────────────────────────────────────

def load_env(path=ENV_PATH):
    """Простой парсер .env (KEY=VALUE построчно, без внешних зависимостей)."""
    env = {}
    if not path.exists():
        return env
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        env[key.strip()] = val.strip()
    return env


ENV = load_env()


def env_flag(name, default=False):
    """Булев флаг из .env (true/1/yes/on)."""
    return ENV.get(name, str(default)).strip().lower() in {"1", "true", "yes", "on"}


def fetch_bytes(req, timeout=15, attempts=2):
    """Повторяет только временные сетевые ошибки, не скрывая постоянные HTTP-ошибки."""
    for attempt in range(1, attempts + 1):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as response:
                return response.read()
        except urllib.error.HTTPError as exc:
            retryable = exc.code == 429 or 500 <= exc.code <= 599
            if attempt >= attempts or not retryable:
                raise
        except (urllib.error.URLError, TimeoutError):
            if attempt >= attempts:
                raise
        time.sleep(attempt)


# ─────────────────────────────────────────────
# КУРСЫ ВАЛЮТ
# ─────────────────────────────────────────────

def get_eur_usd():
    """Получает актуальный курс EUR/USD с open.er-api.com (бесплатный, без ключа)."""
    try:
        req = urllib.request.Request(
            "https://open.er-api.com/v6/latest/EUR",
            headers={"User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req, timeout=10) as r:
            data = json.loads(r.read().decode())
        rate = float(data["rates"]["USD"])
        print(f"Курс EUR/USD (open.er-api): {rate:.4f}")
        return rate
    except Exception as e:
        print(f"Не удалось получить EUR/USD ({e}). Использую 1.13.")
        return 1.13


def get_usd_rub(markup=4.0, manual=None):
    """Берёт официальный курс USD/RUB с ЦБ РФ и прибавляет наценку."""
    if manual:
        rate = float(manual)
        print(f"Курс USD/RUB (задан вручную): {rate} + {markup} = {rate + markup}")
        return rate + markup
    try:
        req = urllib.request.Request("https://www.cbr.ru/scripts/XML_daily.asp",
                                     headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as r:
            xml = r.read().decode("cp1251", "replace")
        m = re.search(r'ID="R01235".*?<Nominal>(\d+)</Nominal>.*?<Value>([\d,]+)</Value>',
                      xml, re.S)
        if m:
            nominal = int(m.group(1))
            val = float(m.group(2).replace(",", ".")) / nominal
            print(f"Курс ЦБ USD/RUB: {val:.4f} + наценка {markup} = {val + markup:.4f}")
            return val + markup
    except Exception as e:
        print(f"Не удалось получить курс ЦБ ({e}).")
    fallback = 72.45 + markup
    print(f"Использую запасной курс USD/RUB: {fallback}")
    return fallback


# ─────────────────────────────────────────────
# ОБЩИЙ ФОРМАТ ПРЕДЛОЖЕНИЯ ДИСТРИБЬЮТОРА
# ─────────────────────────────────────────────
# offer = {
#   "dist": "DigiKey", "pn": "<distributor SKU>", "stock": int, "moq": int,
#   "lead": int|None, "tiers": [{"qty": int, "price": float}], "manu": str,
# }


def canonical_mpn(value):
    """Нормализует только незначимые различия записи MPN, не характеристики."""
    text = str(value or "").translate(CYRILLIC_LOOKALIKES).upper().strip()
    text = text.replace("–", "-").replace("—", "-").replace(",", ".")
    return re.sub(r"\s+", "", text).strip(".;")


def canonical_mpn_loose(value):
    """Как canonical_mpn, но дополнительно убирает все разделители
    (дефисы, точки, слэши, пробелы). Нужно, чтобы разница в РАССТАНОВКЕ
    дефисов не мешала совпадению: '3310Y001-103L' == '3310Y-001-103L'.
    Значимые буквы/цифры сохраняются, поэтому характеристики не путаются."""
    return re.sub(r"[^A-Z0-9]", "", canonical_mpn(value))


# Суффиксы, меняющие только упаковку/исполнение поставки, не характеристики
# (регламент: T/tape&reel, K/bulk, L и LF/RoHS, CT/cut tape, TR; ELF/GLF — Bourns tape&reel lead-free)
PACKAGING_TAILS = {
    "LF", "L", "T", "TR", "CT", "K", "ELF", "GLF",
    # Бессвинцовое исполнение / лента (Analog Devices, Infineon и др.):
    # LT1763CS8 -> LT1763CS8#PBF / LT1763CS8#TRPBF — характеристики те же
    "PBF", "TRPBF", "TRLPBF", "TRPBFT", "G", "GT", "RL", "R7", "E4",
}


def split_mask(requested):
    """Маска регламента «x / X / XX в конце — любой суффикс подходит»:
    - строчный x (или кириллич. х), один и более — маска (3313J-1-104x)
    - хвост из двух и более заглавных X — маска (DLC...CXX)
    - смешанный хвост (Xx) — маска (DLC70B8R2CW501Xx)
    - одиночная заглавная X — часть реального PN (EXB28V220JX)
    База короче 3 значимых символов маской не считается (защита от абсурда).
    Возвращает (база без маски, был_ли_x)."""
    raw = str(requested or "").strip()
    m = re.search(r"[xхXХ]+$", raw)
    if m and m.start() > 0:
        tail = m.group(0)
        has_lower = any(ch in "xх" for ch in tail)
        uppers = sum(1 for ch in tail if ch in "XХ")
        if has_lower or uppers >= 2:
            base = raw[:m.start()].rstrip("-. ")
            if len(canonical_mpn_loose(base)) >= 3:
                return base, True
    return raw, False


def mask_regex(requested):
    """Строчный x в СЕРЕДИНЕ артикула = «любой один символ»
    (регламент: TPSE477K010x0200 -> TPSE477K010R0200).
    Возвращает анкерованный regex по loose-форме или None, если внутренних x нет."""
    base, masked_tail = split_mask(str(requested or "").strip())
    core = base.translate(CYRILLIC_LOOKALIKES)
    if "x" not in core:
        return None
    pattern = []
    for ch in core:
        if ch == "x":
            pattern.append("[A-Z0-9]")
        elif ch.upper().isalnum():
            pattern.append(re.escape(ch.upper()))
        # разделители опускаются — сравнение идёт с loose-формой кандидата
    if not pattern:
        return None
    suffix = ".*" if masked_tail else ""
    return re.compile("^" + "".join(pattern) + suffix + "$")


def mpn_matches(requested, candidate):
    if not requested or not candidate:
        return False
    cand_loose = canonical_mpn_loose(candidate)
    if not cand_loose:
        return False
    # x в середине — точечный wildcard
    rx = mask_regex(requested)
    if rx:
        return bool(rx.match(cand_loose))
    base, masked = split_mask(requested)
    if masked:
        base_loose = canonical_mpn_loose(base)
        return bool(base_loose) and cand_loose.startswith(base_loose)
    if canonical_mpn(requested) == canonical_mpn(candidate):
        return True
    loose = canonical_mpn_loose(requested)
    if not loose:
        return False
    if loose == cand_loose:
        return True
    # Упаковочный суффикс у дистрибьютора: CR0805-JW-390E -> CR0805-JW-390ELF
    if cand_loose.startswith(loose) and cand_loose[len(loose):] in PACKAGING_TAILS:
        return True
    # Достройка серии-префикса производителя: в запросе код серии без буквенного
    # префикса (0805-FX-1503 -> Bourns CR0805-FX-1503ELF). Требования: запрошенный
    # код достаточно длинный, префикс 1-2 буквы, хвост пуст или упаковочный.
    # Производитель дополнительно проверяется в manufacturer_matches.
    if len(loose) >= 6:
        idx = cand_loose.find(loose)
        if 0 < idx <= 2 and cand_loose[:idx].isalpha():
            tail = cand_loose[idx + len(loose):]
            if tail == "" or tail in PACKAGING_TAILS:
                return True
    return False


def canonical_manufacturer(value):
    key = re.sub(r"[^a-z0-9]", "", str(value or "").lower())
    if "avago" in key or "broadcom" in key:
        return "broadcomavago"
    if "axicom" in key or "teconnect" in key:
        return "teconnectivityaxicom"
    if "epcos" in key or key.startswith("tdk"):
        return "tdkepcos"
    if "fujitsu" in key or key.startswith("fcl"):
        return "fujitsufcl"
    # Поглощения: приводим и запрос, и ответ дистрибьютора к общему канону
    for names, canon in ACQUISITION_ALIASES.items():
        if key in names:
            return canon
    for names, canon in ACQUISITION_ALIASES.items():
        if any(len(n) >= 5 and n in key for n in names):
            return canon
    return MANUFACTURER_ALIASES.get(key, key)


def manufacturer_matches(requested, candidate):
    """Не допускает цену другого производителя при явно указанном бренде."""
    if not requested:
        return True
    left = canonical_manufacturer(requested)
    right = canonical_manufacturer(candidate)
    if not left or not right:
        return False
    return left == right or left in right or right in left


def part_search_variants(pn):
    variants = []
    raw = str(pn or "").strip()
    candidates = [raw]
    if re.search(r"[A-Za-z]{2,}", raw):
        candidates.append(
            re.sub(r"^([QO])\s+(?=\d)", r"\1", raw.replace(",", "."), flags=re.I)
        )
    base, masked = split_mask(raw)
    if masked and base:
        candidates.append(base)          # маска x: ищем по базе — API вернёт все суффиксы
    seen = set()
    for value in candidates:
        key = value.casefold()
        if value and key not in seen:
            variants.append(value)
            seen.add(key)
    return variants


def pick_tier(tiers, qty):
    """Возвращает (tier, min_not_met) — тир с наибольшим qty ≤ запрошенного,
    либо самый маленький тир, если такого нет (и флаг min_not_met=True)."""
    ts = sorted(tiers, key=lambda t: t["qty"])
    applicable = [t for t in ts if t["qty"] <= qty]
    if applicable:
        return applicable[-1], False
    return ts[0], True


# ─────────────────────────────────────────────
# DIGIKEY  (Product Information API v4, OAuth2 client_credentials)
# ─────────────────────────────────────────────

_digikey_token_cache = {"token": None, "expires_at": 0}


def _digikey_get_token():
    now = time.time()
    if _digikey_token_cache["token"] and _digikey_token_cache["expires_at"] > now + 30:
        return _digikey_token_cache["token"]

    client_id = ENV.get("DIGIKEY_CLIENT_ID")
    client_secret = ENV.get("DIGIKEY_CLIENT_SECRET")
    if not client_id or not client_secret:
        return None

    data = urllib.parse.urlencode({
        "client_id": client_id, "client_secret": client_secret,
        "grant_type": "client_credentials",
    }).encode()
    req = urllib.request.Request("https://api.digikey.com/v1/oauth2/token", data=data, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    try:
        d = json.loads(fetch_bytes(req).decode())
        _digikey_token_cache["token"] = d["access_token"]
        _digikey_token_cache["expires_at"] = now + int(d.get("expires_in", 600))
        return d["access_token"]
    except Exception as e:
        print(f"    DigiKey OAuth error: {e}")
        return None


def search_digikey(pn, qty, match_pn=None):
    match_pn = match_pn or pn
    client_id = ENV.get("DIGIKEY_CLIENT_ID")
    if not client_id:
        return []
    token = _digikey_get_token()
    if not token:
        return []

    url = "https://api.digikey.com/products/v4/search/keyword"
    body = json.dumps({"Keywords": pn, "Limit": 10}).encode()
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("Authorization", "Bearer " + token)
    req.add_header("X-DIGIKEY-Client-Id", client_id)
    req.add_header("X-DIGIKEY-Locale-Site", "US")
    req.add_header("X-DIGIKEY-Locale-Language", "en")
    req.add_header("X-DIGIKEY-Locale-Currency", "USD")
    req.add_header("Content-Type", "application/json")
    try:
        d = json.loads(fetch_bytes(req).decode())
    except urllib.error.HTTPError as e:
        print(f"    DigiKey HTTP {e.code}: {e.read().decode()[:200]}")
        return []
    except Exception as e:
        print(f"    DigiKey error: {e}")
        return []

    offers = []
    for p in d.get("Products", []) or []:
        # Берём только точное совпадение по PN производителя (без опечаток/аналогов)
        mpn = (p.get("ManufacturerProductNumber") or "").strip()
        if not mpn_matches(match_pn, mpn):
            continue
        tiers = []
        for sp in p.get("StandardPricing", []) or []:
            price = sp.get("UnitPrice")
            brk_qty = sp.get("BreakQuantity")
            if price is not None and brk_qty is not None:
                tiers.append({"qty": int(brk_qty), "price": float(price)})
        if not tiers:
            up = p.get("UnitPrice")
            if up is not None:
                tiers.append({"qty": 1, "price": float(up)})
        if not tiers:
            continue
        stock = int(p.get("QuantityAvailable") or 0)
        moq = int(p.get("MinimumOrderQuantity") or 1)
        manu = ((p.get("Manufacturer") or {}).get("Name")) or ""
        lead_weeks = p.get("ManufacturerLeadWeeks")
        try:
            lead_weeks = int(lead_weeks) if lead_weeks not in (None, "") else None
        except (ValueError, TypeError):
            lead_weeks = None
        offers.append({
            "dist": "DigiKey", "pn": p.get("ProductVariations", [{}])[0].get("DigiKeyProductNumber", "")
                    if p.get("ProductVariations") else (p.get("DigiKeyPartNumber") or ""),
            "mpn": mpn,
            "stock": stock, "moq": moq, "lead": lead_weeks, "tiers": tiers, "manu": manu,
        })
    return offers


# ─────────────────────────────────────────────
# MOUSER  (Search API)
# ─────────────────────────────────────────────

def _mouser_parts(key, pn):
    """Список Parts от Mouser: сначала точный поиск по партномеру,
    при пустом результате — нечёткий keyword-поиск (кандидатов дальше
    фильтрует mpn_matches, так что аналоги не проскочат)."""
    attempts = (
        ("partnumber", {"SearchByPartRequest": {"mouserPartNumber": pn, "partSearchOptions": ""}}),
        ("keyword", {"SearchByKeywordRequest": {"keyword": pn, "records": 15, "startingRecord": 0}}),
    )
    for endpoint, body_obj in attempts:
        url = f"https://api.mouser.com/api/v1.0/search/{endpoint}?apiKey=" + urllib.parse.quote(key)
        req = urllib.request.Request(url, data=json.dumps(body_obj).encode(), method="POST")
        req.add_header("Content-Type", "application/json")
        try:
            d = json.loads(fetch_bytes(req).decode())
        except Exception as e:
            print(f"    Mouser error: {e}")
            continue
        if d.get("Errors"):
            continue
        parts = (d.get("SearchResults") or {}).get("Parts") or []
        if parts:
            return parts
    return []


def search_mouser(pn, qty, match_pn=None):
    match_pn = match_pn or pn
    key = ENV.get("MOUSER_API_KEY")
    if not key:
        return []
    results = _mouser_parts(key, pn)
    offers = []
    for p in results:
        mpn = (p.get("ManufacturerPartNumber") or "").strip()
        if not mpn_matches(match_pn, mpn):
            continue
        tiers = []
        for pb in p.get("PriceBreaks", []) or []:
            price_str = (pb.get("Price") or "").replace("$", "").replace(",", "").strip()
            try:
                price = float(price_str)
            except ValueError:
                continue
            tiers.append({"qty": int(pb.get("Quantity", 1)), "price": price})
        if not tiers:
            continue
        stock_str = re.sub(r"[^\d]", "", p.get("Availability") or "")
        stock = int(stock_str) if stock_str else 0
        moq = int(p.get("Min") or 1)
        lead_weeks = None
        lead_str = p.get("LeadTime") or ""
        m = re.search(r"(\d+)\s*Day", lead_str, re.I)
        if m:
            lead_weeks = max(1, round(int(m.group(1)) / 7))
        else:
            m = re.search(r"(\d+)\s*Week", lead_str, re.I)
            if m:
                lead_weeks = int(m.group(1))
        offers.append({
            "dist": "Mouser", "pn": p.get("MouserPartNumber") or "", "mpn": mpn,
            "stock": stock, "moq": moq, "lead": lead_weeks, "tiers": tiers,
            "manu": p.get("Manufacturer") or "",
        })
    return offers


# ─────────────────────────────────────────────
# NEWARK / FARNELL (element14 Product Search API)
# ─────────────────────────────────────────────

def search_farnell(pn, qty, match_pn=None):
    match_pn = match_pn or pn
    key = ENV.get("FARNELL_API_KEY")
    if not key:
        return []

    def fetch_products(term_field, limit):
        params = (
            f"term={term_field}%3A" + urllib.parse.quote(pn) +
            "&storeInfo.id=www.newark.com"
            "&resultsSettings.offset=0"
            f"&resultsSettings.numberOfResults={limit}"
            "&resultsSettings.responseGroup=large"
            "&callInfo.responseDataFormat=JSON"
            "&callinfo.apiKey=" + key
        )
        req = urllib.request.Request("https://api.element14.com/catalog/products?" + params)
        try:
            d = json.loads(fetch_bytes(req).decode())
        except Exception as e:
            print(f"    Newark/Farnell error: {e}")
            return []
        # обёртка ответа зависит от типа term: manufacturerPartNumberSearchReturn / keywordSearchReturn
        for value in (d or {}).values():
            if isinstance(value, dict) and "products" in value:
                return value.get("products") or []
        return []

    # Точный поиск по MPN; при пустом — нечёткий поиск по всем полям (any),
    # кандидатов дальше фильтрует mpn_matches.
    products = fetch_products("manuPartNum", 5) or fetch_products("any", 15)
    offers = []
    for p in products:
        mpn = (p.get("translatedManufacturerPartNumber") or "").strip()
        if not mpn_matches(match_pn, mpn):
            continue
        tiers = []
        for pb in p.get("prices", []) or []:
            cost = pb.get("cost")
            frm = pb.get("from")
            if cost is not None and frm is not None:
                tiers.append({"qty": int(frm), "price": float(cost)})
        if not tiers:
            continue
        stock_info = p.get("stock") or {}
        stock = int(stock_info.get("level") or 0)
        lead = stock_info.get("leastLeadTime")
        moq = int(p.get("translatedMinimumOrderQuality") or 1)
        offers.append({
            "dist": "Newark/Farnell", "pn": p.get("sku") or "", "mpn": mpn,
            "stock": stock, "moq": moq,
            "lead": int(lead) if lead else None,
            "tiers": tiers, "manu": p.get("brandName") or "",
        })
    return offers


# ─────────────────────────────────────────────
# TME  (Product API v2, OAuth 2.0 client_credentials)
# ─────────────────────────────────────────────

_tme_token_cache = {"token": None, "expires_at": 0}


def _tme_get_token():
    now = time.time()
    if _tme_token_cache["token"] and _tme_token_cache["expires_at"] > now + 30:
        return _tme_token_cache["token"]

    token = ENV.get("TME_TOKEN")        # 50-значный private key (логин Basic Auth)
    secret = ENV.get("TME_APP_SECRET")  # 20-значный application secret (пароль)
    if not token or not secret:
        return None

    import base64
    basic = base64.b64encode(f"{token}:{secret}".encode()).decode()
    data = urllib.parse.urlencode({"grant_type": "client_credentials"}).encode()
    req = urllib.request.Request("https://api.tme.eu/auth/token", data=data, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    req.add_header("Authorization", "Basic " + basic)
    try:
        d = json.loads(fetch_bytes(req).decode())
        _tme_token_cache["token"] = d["access_token"]
        _tme_token_cache["expires_at"] = now + int(d.get("expires_in", 300))
        return d["access_token"]
    except Exception as e:
        print(f"    TME OAuth error: {e}")
        return None


def _tme_get(url, access):
    req = urllib.request.Request(url)
    req.add_header("Authorization", "Bearer " + access)
    return json.loads(fetch_bytes(req).decode())


def search_tme(pn, qty, match_pn=None):
    match_pn = match_pn or pn
    access = _tme_get_token()
    if not access:
        return []

    # 1) ищем товар по MPN производителя (точный справочник TME)
    try:
        d = _tme_get("https://api.tme.eu/products?" + urllib.parse.urlencode(
            [("country", "DE"), ("mpns[]", pn)]), access)
    except Exception as e:
        print(f"    TME error: {e}")
        return []

    def collect_matches(products):
        found = {}
        for p in products:
            symbols = [p.get("symbol") or ""] + list(p.get("manufacturer_symbols") or [])
            if any(mpn_matches(match_pn, s) for s in symbols if s):
                found[p["symbol"]] = p
        return found

    matches = collect_matches((d.get("data") or {}).get("elements") or [])

    # 1б) нечёткий fallback: полнотекстовый поиск по фразе (кандидатов
    # фильтрует тот же mpn_matches — аналоги не пройдут)
    if not matches:
        phrase = re.sub(r"\s+", " ", str(pn)).strip()[:40]
        if len(phrase) >= 2:
            try:
                d = _tme_get("https://api.tme.eu/products/search?" + urllib.parse.urlencode(
                    [("country", "DE"), ("phrase", phrase), ("scope[]", "products"), ("limit", "20")]), access)
                products = ((d.get("data") or {}).get("products") or {}).get("elements") or []
                matches = collect_matches(products)
            except Exception:
                pass
    if not matches:
        return []

    # 2) цены + сток по символам TME (валюта USD)
    try:
        qs = [("country", "DE"), ("currency", "USD"),
              ("scope[]", "prices"), ("scope[]", "stock")]
        qs += [("symbols[]", s) for s in list(matches)[:50]]
        d = _tme_get("https://api.tme.eu/products/data?" + urllib.parse.urlencode(qs), access)
    except Exception as e:
        print(f"    TME data error: {e}")
        return []

    offers = []
    for el in (d.get("data") or {}).get("elements") or []:
        sym = el.get("symbol")
        prod = matches.get(sym)
        if not prod:
            continue
        price_info = el.get("prices") or {}
        tiers = []
        for t in price_info.get("elements") or []:
            if t.get("price") is not None and t.get("amount") is not None:
                tiers.append({"qty": int(t["amount"]), "price": float(t["price"])})
        if not tiers:
            continue
        stock = int(el.get("stock_quantity") or 0)
        moq = int(prod.get("minimal_amount") or 1)
        manu = ((prod.get("manufacturer") or {}).get("name")) or ""
        matched_mpn = next((m for m in (prod.get("manufacturer_symbols") or []) if mpn_matches(match_pn, m)), pn)
        offers.append({
            "dist": "TME", "pn": sym, "mpn": matched_mpn, "stock": stock, "moq": moq,
            "lead": None, "tiers": tiers, "manu": manu,
        })
    return offers


# ─────────────────────────────────────────────
# ВЫБОР ЛУЧШЕЙ ЦЕНЫ
# ─────────────────────────────────────────────

def _collect_offers(query, qty, manufacturer, match_pn=None):
    """Опрашивает все API одним запросом и фильтрует по производителю.
    match_pn — исходный артикул (с маской x), против которого проверяются кандидаты."""
    query_offers = []
    for search_fn in (search_digikey, search_mouser, search_farnell, search_tme):
        try:
            query_offers.extend(search_fn(query, qty, match_pn=match_pn))
        except Exception as e:
            print(f"    {search_fn.__name__} ошибка: {e}")
    return [o for o in query_offers if manufacturer_matches(manufacturer, o.get("manu", ""))]


def find_best_price(pn, qty, manufacturer=""):
    print(f"    qty={qty}")
    offers = []
    for query in part_search_variants(pn):
        offers = _collect_offers(query, qty, manufacturer, match_pn=pn)
        if offers:
            break

    # Второй проход: если запрос отличался от настоящего MPN (разделители, маска,
    # суффикс) — повторяем поиск каждым найденным точным MPN (до 3 различных),
    # чтобы подтянуть дистрибьюторов с точечным поиском (Mouser/TME/Farnell,
    # часто дешевле) и выбрать реальный минимум по регламенту.
    if offers:
        seen = {(o["dist"], o.get("pn", "")) for o in offers}
        tried = set()
        for o in list(offers):
            rmpn = (o.get("mpn") or "").strip()
            if not rmpn or canonical_mpn(rmpn) == canonical_mpn(pn):
                continue
            key = canonical_mpn_loose(rmpn)
            if key in tried:
                continue
            if len(tried) >= 3:
                break
            tried.add(key)
            for extra in _collect_offers(rmpn, qty, manufacturer, match_pn=rmpn):
                sig = (extra["dist"], extra.get("pn", ""))
                if sig not in seen:
                    offers.append(extra)
                    seen.add(sig)

    if not offers:
        return {"pn": pn, "qty": qty, "manufacturer": manufacturer, "status": "RFQ"}
    return _best_candidate(offers, pn, qty)


def _best_candidate(offers, pn, qty):
    """Выбор лучшего предложения: приоритет — сток ≥ количества, затем мин. цена."""
    best = None
    for o in offers:
        tier, min_not_met = pick_tier(o["tiers"], qty)
        price_usd = tier["price"]
        stock = o.get("stock", 0)
        # В наличии = стока хватает на запрошенное кол-во.
        # Lead time (недели) актуален только при нехватке стока.
        in_stock = stock >= qty
        lead = None if in_stock else o.get("lead")
        cand = {
            "pn": pn, "distr_pn": o.get("pn", ""),
            "resolved_mpn": o.get("mpn", ""),
            "manufacturer": o.get("manu", ""),
            "distributor": o.get("dist", ""), "moq": o.get("moq", 1),
            "stock": stock, "lead": lead,
            "price_usd": price_usd,
            "qty": qty, "in_stock": in_stock, "min_not_met": min_not_met, "status": "FOUND",
        }
        key = (0 if in_stock else 1, price_usd)
        if best is None or key < best[0]:
            best = (key, cand)
    return best[1]


# ─────────────────────────────────────────────
# OEMSECRETS Part Search API — добивочный фолбэк (квота, по умолч. 10/день)
# ─────────────────────────────────────────────

QUOTA_PATH = SCRIPT_DIR / "oemsecrets-quota.json"

# Авторизованные дистрибьюторы регламента для фильтрации выдачи oemsecrets
# (маркетплейсы крупных дистрибьюторов допустимы: Verical = Arrow)
OEMSECRETS_AUTHORIZED = (
    "digikey", "digi-key", "mouser", "tme", "farnell", "newark", "element14",
    "arrow", "verical", "tti", "avnet", "rs comp", "rs-online", "rs de",
    "rs americas", "future", "heilind", "rutronik", "buerklin", "bürklin",
)


def _oemsecrets_quota_take():
    """True, если дневная квота позволяет ещё один запрос (и списывает его)."""
    import datetime as _dt
    limit = int(ENV.get("OEMSECRETS_DAILY_LIMIT", "10"))
    today = _dt.date.today().isoformat()
    try:
        state = json.loads(QUOTA_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        state = {}
    if state.get("date") != today:
        state = {"date": today, "used": 0}
    if state["used"] >= limit:
        return False
    state["used"] += 1
    QUOTA_PATH.write_text(json.dumps(state), encoding="utf-8")
    return True


def search_oemsecrets(pn, qty, match_pn=None):
    """Один запрос к oemsecrets Part Search API (140+ дистрибьюторов).
    Вызывается только для RFQ-остатка и только в пределах дневной квоты.
    В ответе может не быть поля производителя — тогда принимаем только
    строгое совпадение MPN (без масок/достроек)."""
    match_pn = match_pn or pn
    key = ENV.get("OEMSECRETS_API_KEY")
    if not key:
        return []
    url = "https://oemsecretsapi.com/partsearch?" + urllib.parse.urlencode(
        [("apiKey", key), ("searchTerm", pn), ("currency", "USD"), ("countryCode", "US")])
    try:
        d = json.loads(fetch_bytes(urllib.request.Request(url), timeout=25).decode())
    except Exception as e:
        print(f"    oemsecrets error: {e}")
        return []

    offers = []
    for item in d.get("stock") or []:
        dist_info = item.get("distributor") or {}
        dist_name = (dist_info.get("distributor_common_name")
                     or dist_info.get("distributor_name") or "").strip()
        if not any(a in dist_name.lower() for a in OEMSECRETS_AUTHORIZED):
            continue                                     # брокеры/маркетплейсы — мимо
        mpn = (item.get("part_number") or "").strip()
        manu = (item.get("manufacturer") or item.get("manufacturer_name") or "").strip()
        if manu:
            if not mpn_matches(match_pn, mpn):
                continue
        else:
            # производитель неизвестен — только строгое совпадение записи MPN
            if canonical_mpn(match_pn) != canonical_mpn(mpn):
                continue
        tiers = []
        for pb in (item.get("prices") or {}).get("USD") or []:
            try:
                tiers.append({"qty": int(float(pb.get("unit_break") or 1)),
                              "price": float(pb.get("unit_price"))})
            except (TypeError, ValueError):
                continue
        if not tiers:
            continue
        stock = int(item.get("quantity_in_stock") or 0)
        lead_digits = re.sub(r"[^\d]", "", str(item.get("lead_time") or ""))
        lead = int(lead_digits) if lead_digits else None
        offers.append({
            "dist": dist_name, "pn": item.get("source_part_number") or mpn, "mpn": mpn,
            "stock": stock, "moq": min(t["qty"] for t in tiers),
            "lead": lead, "tiers": tiers, "manu": manu,
        })
    return offers


# ─────────────────────────────────────────────
# LLM-КОНТУР: подсказка артикула + ОБЯЗАТЕЛЬНАЯ проверка через API
# ─────────────────────────────────────────────

CLAUDE_SYSTEM = (
    "Ты — детерминированный резолвер артикулов ЭКБ. Отвечай ТОЛЬКО валидным "
    "JSON-массивом в одну выдачу, без приветствий, пояснений и markdown. "
    "Никогда не задавай уточняющих вопросов. Идентифицируй ТОТ ЖЕ компонент, "
    "не аналог: другой номинал/напряжение/корпус/точность запрещены. Разрешено "
    "исправить опечатку, дописать префикс/суффикс серии, указать актуальный "
    "бренд после поглощений, развернуть неполный артикул. Цены/склад/сроки не "
    "придумывай. Не уверен — suggested_mpn=null."
)

CLAUDE_PROMPT = """Для каждой позиции верни объект. Формат ответа — только этот JSON-массив:
[{"requested_pn":"<как в запросе>","suggested_mpn":"<полный артикул>"|null,"suggested_manufacturer":"<бренд>"|null,"reason":"<кратко>"}]

Позиции:
"""


def _claude_binary():
    """Путь к CLI. На Windows npm ставит обёртку claude.cmd, которую
    subprocess без расширения не находит."""
    configured = ENV.get("CLAUDE_BIN")
    if configured:
        return configured
    import shutil
    for name in ("claude.cmd", "claude.exe", "claude"):
        found = shutil.which(name)
        if found:
            return found
    return "claude"


def claude_suggest(items, timeout=180):
    """Спрашивает headless-Claude про нерешённые позиции.
    Возвращает список подсказок (может быть пустым). Цены НЕ запрашиваются."""
    binary = _claude_binary()
    payload = [{"requested_pn": it["pn"], "manufacturer": it.get("manufacturer", ""),
                "description": it.get("description", "")[:120]} for it in items]
    prompt = CLAUDE_PROMPT + json.dumps(payload, ensure_ascii=False, indent=1)
    # Запускаем в отдельной пустой папке, чтобы CLI не подхватывал AGENTS.md/
    # CLAUDE.md проекта и не «уходил в диалог»; системным промптом форсим JSON.
    import tempfile
    try:
        with tempfile.TemporaryDirectory() as workdir:
            # Промпт — через stdin: на Windows .cmd-обёртка ломает кавычки/скобки
            # JSON, если передавать его как аргумент командной строки.
            completed = subprocess.run(
                [binary, "-p", "--output-format", "json",
                 "--append-system-prompt", CLAUDE_SYSTEM],
                input=prompt, capture_output=True, text=True, timeout=timeout,
                encoding="utf-8", errors="replace", cwd=workdir)
    except (OSError, subprocess.TimeoutExpired) as exc:
        print(f"    LLM-контур недоступен: {exc}")
        return []
    raw = completed.stdout or ""
    if not raw.strip():
        print(f"    LLM-контур: пустой ответ ({(completed.stderr or '').strip()[:80]})")
        return []
    # claude --output-format json оборачивает текст в объект; достаём массив
    text = raw
    try:
        wrapper = json.loads(raw)
        if isinstance(wrapper, dict):
            text = wrapper.get("result") or wrapper.get("text") or raw
    except json.JSONDecodeError:
        pass
    match = re.search(r"\[.*\]", text, re.S)
    if not match:
        return []
    try:
        data = json.loads(match.group(0))
    except json.JSONDecodeError:
        return []
    return data if isinstance(data, list) else []


def claude_assisted_pass(results):
    """Для оставшихся RFQ: подсказка LLM -> ПРОВЕРКА через API дистрибьюторов.
    В файл попадает только то, что реально подтвердилось ценой у дистрибьютора;
    строка помечается llm_assisted=True, чтобы человек видел такие позиции."""
    if not env_flag("ENABLE_CLAUDE_FALLBACK"):
        return
    pending = [(i, r) for i, r in enumerate(results)
               if r.get("status") == "RFQ" and r.get("pn")]
    if not pending:
        return
    limit = int(ENV.get("CLAUDE_MAX_ITEMS", "40"))
    batch = [r for _, r in pending[:limit]]
    print(f"\nLLM-контур: уточняю {len(batch)} нерешённых позиций...")
    suggestions = claude_suggest(batch)
    if not suggestions:
        print("LLM-контур: подсказок нет")
        return

    by_pn = {}
    for s in suggestions:
        if isinstance(s, dict) and s.get("requested_pn"):
            by_pn[canonical_mpn_loose(s["requested_pn"])] = s

    recovered = 0
    for i, r in pending:
        s = by_pn.get(canonical_mpn_loose(r["pn"]))
        if not s:
            continue
        mpn = (s.get("suggested_mpn") or "").strip()
        manu = (s.get("suggested_manufacturer") or r.get("manufacturer") or "").strip()
        if not mpn or not manu:
            continue                      # без бренда цену не подставляем (регламент)
        qty = r.get("qty", 1)
        # ПРОВЕРКА: ищем предложенный артикул в API; принимаем только подтверждённое
        offers = [o for o in _collect_offers(mpn, qty, manu, match_pn=mpn)
                  if mpn_matches(mpn, o.get("mpn", ""))]
        if not offers:
            continue
        cand = _best_candidate(offers, r["pn"], qty)
        cand["description"] = r.get("description", r["pn"])
        cand["manufacturer"] = cand.get("manufacturer") or manu
        cand["llm_assisted"] = True
        cand["llm_reason"] = (s.get("reason") or "")[:120]
        results[i] = cand
        recovered += 1
        print(f"    LLM+API: {r['pn']} -> {cand.get('resolved_mpn') or mpn} "
              f"| {cand['distributor']} ${cand['price_usd']}  ({cand['llm_reason']})")
    print(f"LLM-контур подтвердил позиций: {recovered} из {len(batch)}")


def oemsecrets_pass(results):
    """Добивочный проход по RFQ-остатку через oemsecrets (в пределах квоты).
    Позиции с бо́льшим количеством — первыми (важнее для закупки)."""
    if not ENV.get("OEMSECRETS_API_KEY"):
        return
    rfq_indexes = [i for i, r in enumerate(results)
                   if r.get("status") == "RFQ" and r.get("pn")]
    rfq_indexes.sort(key=lambda i: results[i].get("qty", 0), reverse=True)
    recovered = 0
    for i in rfq_indexes:
        if not _oemsecrets_quota_take():
            print("    oemsecrets: дневная квота исчерпана — остаток завтра")
            break
        r = results[i]
        offers = [o for o in search_oemsecrets(r["pn"], r.get("qty", 1))
                  if manufacturer_matches(r.get("manufacturer", ""), o.get("manu", ""))
                  or not o.get("manu")]
        if not offers:
            continue
        cand = _best_candidate(offers, r["pn"], r.get("qty", 1))
        cand["description"] = r.get("description", r["pn"])
        if not cand.get("manufacturer"):
            cand["manufacturer"] = r.get("manufacturer", "")
        results[i] = cand
        recovered += 1
        print(f"    oemsecrets: {r['pn']} -> {cand['distributor']} ${cand['price_usd']}")
    if recovered:
        print(f"oemsecrets-фолбэк добрал позиций: {recovered}")


# ─────────────────────────────────────────────
# ИЗВЛЕЧЕНИЕ АРТИКУЛА ИЗ ОПИСАНИЯ
# ─────────────────────────────────────────────

def clean_spaces(value):
    return re.sub(r"\s+", " ", str(value or "").replace("\u00a0", " ")).strip()


def _looks_like_brand_token(token):
    """Токен похож на часть названия бренда: есть подряд ≥2 латинских буквы,
    букв больше чем цифр, нет процентов/единиц измерения. Пропускает '3Peak',
    'Rectifier', 'Instruments'; отсекает '100nF', '0805', '±5%'."""
    t = str(token or "").strip("()[],;")
    if not t or "%" in t or "±" in t:
        return False
    if not re.search(r"[A-Za-z]{2,}", t):
        return False
    letters = sum(ch.isalpha() and ch.isascii() for ch in t)
    digits = sum(ch.isdigit() for ch in t)
    if digits >= letters:
        return False
    # единицы измерения и типоразмеры не бренд
    if re.fullmatch(r"\d+(?:[.,]\d+)?\s*(?:[munpkKMG]?[FHΩVWAf]|nF|uF|pF|mH|uH|kHz|MHz|GHz|ppm|SMD|DIP|SOT\d*|SOIC\d*)",
                    t, re.I):
        return False
    return True


def _looks_like_mpn_token(token):
    """Токен похож на артикул: латиница и либо ≥2 цифр, либо разделитель (-/#.)
    рядом с цифрой. Так 'TPW4051-SR' и 'LT1763CS8' — артикулы, а '3Peak' — бренд."""
    t = str(token or "").strip("()[],;")
    if len(t) < 4 or not re.search(r"[A-Za-z]", t):
        return False
    digits = sum(ch.isdigit() for ch in t)
    if digits >= 2:
        return True
    return bool(digits and re.search(r"[-/#.]", t))


def split_manufacturer(desc, manufacturer_hint=""):
    """Возвращает (описание без бренда, бренд), используя колонку бренда как приоритет.
    Сначала ищет бренд по списку KNOWN_MFR, затем — общим правилом: замыкающие
    словоподобные токены после артикула (покрывает бренды вне списка, напр. '3Peak')."""
    text = clean_spaces(desc)
    candidates = []
    if manufacturer_hint:
        candidates.append(clean_spaces(manufacturer_hint))
    candidates.extend(KNOWN_MFR)
    seen = set()
    for mfr in sorted(candidates, key=len, reverse=True):
        key = mfr.casefold()
        if not mfr or key in seen:
            continue
        seen.add(key)
        match = re.search(r"(?:\s*[,;\-]\s*|\s+)" + re.escape(mfr) + r"\s*$", text, re.I)
        if match:
            return text[:match.start()].strip(" -,;"), clean_spaces(manufacturer_hint) or mfr

    if manufacturer_hint:
        return text, clean_spaces(manufacturer_hint)

    # Общее правило: до 3 замыкающих «словесных» токенов считаются брендом,
    # но только если ПЕРЕД ними остаётся токен, похожий на артикул
    # ("Микросхема TPW4051-SR 3Peak" -> бренд '3Peak';
    #  "Микросхема LT1763CS8" -> бренда нет, артикул не съедаем).
    tokens = text.split()
    tail = 0
    while tail < 3 and len(tokens) - tail - 1 >= 0:
        token = tokens[-1 - tail]
        # артикул брендом не считаем — иначе съедим сам MPN (TPW4051-SR)
        if _looks_like_mpn_token(token) or not _looks_like_brand_token(token):
            break
        tail += 1
    if tail:
        head = tokens[:len(tokens) - tail]
        if any(_looks_like_mpn_token(t) for t in head):
            brand = " ".join(tokens[len(tokens) - tail:]).strip(" -,;")
            return " ".join(head).strip(" -,;"), brand
    return text, None


def strip_product_type(text):
    body = clean_spaces(text)
    for word in sorted(RU_TYPE_WORDS, key=len, reverse=True):
        updated = re.sub(r"^" + re.escape(word) + r"(?:\s+|\s*[:\-]\s*)", "", body, count=1, flags=re.I)
        if updated != body:
            return updated.strip(" -:")
    # Неизвестное русское название изделия в начале также не является MPN.
    return re.sub(r"^(?:[А-ЯЁа-яё][А-ЯЁа-яё-]*\s+)+(?=.*\d)", "", body).strip(" -:")


def normalize_cyrillic_units(text):
    value = str(text or "")
    replacements = (
        (r"(?<=\d)\s*МОМ\b", "M"),
        (r"(?<=\d)\s*КОМ\b", "K"),
        (r"(?<=\d)\s*ОМ\b", "R"),
        (r"(?<=\d)\s*МГЦ\b", "MHz"),
        (r"(?<=\d)\s*КГЦ\b", "kHz"),
        (r"(?<=\d)\s*ГЦ\b", "Hz"),
    )
    for pattern, replacement in replacements:
        value = re.sub(pattern, replacement, value, flags=re.I)
    value = re.sub(r"(?<=%)\s*[Хх]\b", "", value)
    value = value.translate(CYRILLIC_LOOKALIKES)
    value = re.sub(r"[А-ЯЁа-яё]+", "", value)
    return clean_spaces(value).strip(" -,;")


def _token_value(token):
    return str(token or "").strip("()[]{};,\"")


def _append_part_context(tokens, index, candidate):
    start = index
    if index > 0 and re.fullmatch(r"[QO]", _token_value(tokens[index - 1]), re.I):
        start -= 1
    elif index > 0 and re.fullmatch(r"\d{1,3}", _token_value(tokens[index - 1])) and "/" in candidate:
        start -= 1

    selected = [_token_value(t) for t in tokens[start:index + 1]]
    tail = [_token_value(t) for t in tokens[index + 1:index + 3]]
    if tail:
        if re.fullmatch(r"\d+(?:[.,]\d+)?(?:k?Hz|MHz|GHz)", tail[0], re.I):
            selected.append(tail[0])
        elif len(tail) >= 2 and re.fullmatch(r"\d+(?:[.,]\d+)?", tail[0]) and re.fullmatch(r"(?:k?Hz|MHz|GHz)", tail[1], re.I):
            selected.extend(tail[:2])
        elif re.fullmatch(r"[A-Z]{1,3}", tail[0]) and "/" in candidate:
            selected.append(tail[0])
    return " ".join(part for part in selected if part).strip(" -,;")


def extract_pn_from_description(desc, manufacturer_hint=""):
    """Отделяет тип изделия и бренд, сохраняя запятые, дроби и значимые префиксы MPN."""
    if not desc:
        return None
    body, _ = split_manufacturer(desc, manufacturer_hint)
    body = strip_product_type(body)
    if not body:
        return None

    # Реальные отечественные обозначения не латинизируем и не режем.
    domestic = re.search(r"\b\d[А-ЯЁ]{1,5}\d[А-ЯЁ0-9.\-/]*\b", body, re.I)
    if domestic:
        return domestic.group(0).strip(" -,;")

    raw_tokens = body.split()
    for index, raw in enumerate(raw_tokens):
        token = _token_value(raw)
        if re.search(r"[А-ЯЁа-яё]", token):
            continue
        if re.search(r"[A-Za-z]", token) and re.search(r"\d", token):
            return _append_part_context(raw_tokens, index, token)

    cleaned = normalize_cyrillic_units(body)
    cleaned_tokens = cleaned.split()
    for index, raw in enumerate(cleaned_tokens):
        token = _token_value(raw)
        if re.search(r"\d", token) and (re.search(r"[A-Za-z]", token) or re.search(r"[./,+\-%]", token)):
            return _append_part_context(cleaned_tokens, index, token)
    return cleaned or None


# ─────────────────────────────────────────────
# ЧТЕНИЕ ЗАПРОСА (Excel / txt)
# ─────────────────────────────────────────────

def _find(headers, *keys):
    for idx, h in headers:
        hl = str(h).lower()
        if any(k in hl for k in keys):
            return idx
    return None


def read_request_xlsx(path):
    """Авто-определение шапки и колонок. Возвращает список позиций."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    # ищем строку-шапку: есть и «наименование/part number», и «количеств/qty»
    header_row = None
    for r in range(1, min(ws.max_row, 40) + 1):
        cells = [(c, ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        texts = " ".join(str(v).lower() for _, v in cells if v)
        if (("наименован" in texts or "part number" in texts or "partnumber" in texts or "условное обознач" in texts)
                and ("количеств" in texts or "qty" in texts or "quantity" in texts)):
            header_row = r
            headers = [(c, v) for c, v in cells if v]
            break
    if header_row is None:
        raise ValueError("Не нашёл строку-шапку в Excel. Проверьте файл или подайте input.txt.")

    col_num = _find(headers, "№", "n п/п", "п/п")
    col_pn = _find(headers, "условное обознач", "part number", "partnumber", "артикул", "парт")
    col_desc = _find(headers, "наименован", "товар")
    col_qty = _find(headers, "количеств", "qty", "quantity", "кол-во")
    col_mfr = _find(headers, "производител", "manufacturer", "изготовит", "brand")

    print(f"Шапка в строке {header_row}. Колонки: PN={col_pn}, Наименование={col_desc}, "
          f"Кол-во={col_qty}, Производитель={col_mfr}")

    items = []
    for r in range(header_row + 1, ws.max_row + 1):
        num = ws.cell(row=r, column=col_num).value if col_num else None
        desc = str(ws.cell(row=r, column=col_desc).value).strip() if col_desc and ws.cell(row=r, column=col_desc).value else ""
        pn_cell = str(ws.cell(row=r, column=col_pn).value).strip() if col_pn and ws.cell(row=r, column=col_pn).value else ""
        qty_v = ws.cell(row=r, column=col_qty).value if col_qty else None
        mfr_cell = str(ws.cell(row=r, column=col_mfr).value).strip() if col_mfr and ws.cell(row=r, column=col_mfr).value else ""

        if not desc and not pn_cell:
            continue
        # Строка-примечание: нет количества → пропускаем
        if qty_v is None or qty_v == "":
            continue
        try:
            qty = int(float(qty_v))
        except (ValueError, TypeError):
            continue                     # нечисловое кол-во — тоже примечание
        if qty <= 0:
            continue
        # Производитель: из отдельной колонки, иначе из хвоста наименования.
        mfr = mfr_cell
        if not mfr and desc:
            _, m = split_manufacturer(desc)
            mfr = m or ""
        # Даже колонка PN иногда содержит «тип + MPN + бренд», поэтому чистим оба варианта.
        pn_source = pn_cell or desc
        pn = extract_pn_from_description(pn_source, mfr) or ""
        items.append({
            "num": int(num) if isinstance(num, (int, float)) else len(items) + 1,
            "pn": pn, "qty": qty, "manufacturer": mfr, "description": desc or pn,
        })
    return items


def read_txt(path):
    items = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split()
            pn = parts[0]
            qty = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 1
            items.append({"num": len(items) + 1, "pn": pn, "qty": qty,
                          "manufacturer": "", "description": pn})
    return items


def autopick_request():
    """Находит свежий Excel-запрос рядом со скриптом (не эталон, не превью, не вывод)."""
    candidates = []
    for d in (SCRIPT_DIR, SCRIPT_DIR / "start"):
        if not d.exists():
            continue
        for f in d.glob("*.xlsx"):
            n = f.name.lower()
            if n.startswith("bom_приложение") or n.startswith("~$") or f.name == PREVIEW_NAME:
                continue
            candidates.append(f)
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


# ─────────────────────────────────────────────
# ПРЕВЬЮ
# ─────────────────────────────────────────────

def write_preview(items, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Превью"
    hdr = ["№", "Распознанный PN (можно править)", "Кол-во", "Производитель", "Исходное наименование"]
    widths = [6, 34, 10, 22, 50]
    fill = PatternFill("solid", fgColor="1F4E79")
    for c, (h, w) in enumerate(zip(hdr, widths), 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = w
    warn = PatternFill("solid", fgColor="FFF2CC")
    for i, it in enumerate(items, 2):
        ws.cell(row=i, column=1, value=it["num"])
        c2 = ws.cell(row=i, column=2, value=it["pn"])
        ws.cell(row=i, column=3, value=it["qty"])
        ws.cell(row=i, column=4, value=it.get("manufacturer") or None)
        ws.cell(row=i, column=5, value=it.get("description") or None)
        if not it["pn"]:
            c2.fill = warn
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def read_preview(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    items = []
    for r in range(2, ws.max_row + 1):
        num = ws.cell(row=r, column=1).value
        pn = ws.cell(row=r, column=2).value
        qty = ws.cell(row=r, column=3).value
        mfr = ws.cell(row=r, column=4).value
        desc = ws.cell(row=r, column=5).value
        if not pn and not desc:
            continue
        try:
            qty = int(float(qty)) if qty not in (None, "") else 1
        except (ValueError, TypeError):
            qty = 1
        mfr_text = str(mfr).strip() if mfr else ""
        # Колонка PN в превью уже содержит извлечённый (или вручную исправленный
        # пользователем) артикул — берём как есть. Повторный прогон extract_* по
        # уже чистому PN портит сложные обозначения (резонаторы: пробелы, запятые,
        # дроби), поэтому НЕ извлекаем повторно.
        pn_text = str(pn).strip() if pn else ""
        items.append({"num": num or len(items) + 1, "pn": pn_text,
                      "qty": qty, "manufacturer": mfr_text,
                      "description": str(desc).strip() if desc else (str(pn) if pn else "")})
    return items


# ─────────────────────────────────────────────
# ВЫВОД (10 колонок)
# ─────────────────────────────────────────────

def packaging_only_suffix(req_pn, dist_pn):
    if not dist_pn or not req_pn:
        return False
    a, b = req_pn.upper().replace(" ", ""), dist_pn.upper().replace(" ", "")
    if a == b:
        return False
    base = re.sub(r"X+$", "", a)
    if base and b.startswith(base):
        return True
    return b.startswith(a) or a.startswith(b)


def write_results(results, output_path, rate_rub):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM"

    F_HDR = PatternFill("solid", fgColor="1F4E79")
    F_HDRB = PatternFill("solid", fgColor="375623")
    F_R1 = PatternFill("solid", fgColor="EBF3FB")
    F_R2 = PatternFill("solid", fgColor="FFFFFF")
    F_RFQ = PatternFill("solid", fgColor="FFF2CC")
    F_BCOL = PatternFill("solid", fgColor="E2EFDA")
    # Позиции, восстановленные подсказкой LLM и подтверждённые API — отдельный
    # цвет, чтобы закупщик мог выборочно перепроверить именно их
    F_LLM = PatternFill("solid", fgColor="FDE9D9")
    THIN = Side(style="thin", color="BFBFBF")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    headers = [
        ("Part Number", 30, F_HDR), ("Distributor Part Number", 26, F_HDRB),
        ("Qty for Single BOM", 11, F_HDR), ("Manufacturer", 20, F_HDR),
        ("Distributor", 20, F_HDR), ("Minimum Order", 11, F_HDR),
        ("Stock", 10, F_HDR), ("Lead Time (weeks)", 13, F_HDR),
        ("Unit Price USD", 14, F_HDR), ("Unit Price RUB", 15, F_HDR),
    ]
    for col, (h, w, fill) in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 32

    found_idx = 0
    for i, r in enumerate(results):
        row = i + 2
        ws.row_dimensions[row].height = 16
        pn = r.get("pn") or extract_pn_from_description(
            r.get("description", ""), r.get("manufacturer", "")
        ) or ""
        if r.get("status") == "FOUND":
            base = F_LLM if r.get("llm_assisted") else (F_R1 if found_idx % 2 == 0 else F_R2)
            found_idx += 1
            # Колонка B: найденный конкретный артикул производителя, когда он
            # отличается от запрошенного (маска x, суффикс упаковки, написание).
            resolved = r.get("resolved_mpn") or ""
            dist_pn = r.get("distr_pn") or ""
            if resolved and canonical_mpn(resolved) != canonical_mpn(pn):
                b_val = resolved
            else:
                b_val = dist_pn if packaging_only_suffix(pn, dist_pn) else ""
            lead = r.get("lead")
            stock_v = r.get("stock", 0)
            if r.get("in_stock") and not lead:
                lead_val = "In stock"
            elif lead is not None:
                lead_val = f"{lead} weeks" if lead != 1 else "1 week"
            elif not stock_v:
                lead_val = "уточнить у поставщика"
            else:
                lead_val = ""
            usd = r.get("price_usd")
            rub = round(usd * rate_rub, 2) if usd is not None else None
            vals = [
                (pn, base, None, False, "000000"),
                (b_val or None, F_BCOL if b_val else base, None, False, "000000"),
                (r.get("qty"), base, None, False, "000000"),
                (r.get("manufacturer") or None, base, None, False, "000000"),
                (r.get("distributor"), base, None, False, "000000"),
                (r.get("moq"), base, None, False, "000000"),
                (r.get("stock"), base, None, False, "000000"),
                (lead_val, base, None, False, "000000"),
                (usd, base, "#,##0.000", False, "000000"),
                (rub, base, "#,##0.00", False, "000000"),
            ]
        else:
            vals = [
                (pn, F_RFQ, None, False, "000000"),
                (None, F_RFQ, None, False, "000000"),
                (r.get("qty"), F_RFQ, None, False, "000000"),
                (r.get("manufacturer") or None, F_RFQ, None, False, "000000"),
                (None, F_RFQ, None, False, "000000"), (None, F_RFQ, None, False, "000000"),
                (None, F_RFQ, None, False, "000000"), (None, F_RFQ, None, False, "000000"),
                ("RFQ", F_RFQ, None, True, "7F4B00"), ("RFQ", F_RFQ, None, True, "7F4B00"),
            ]
        for col, (v, fill, fmt, bold, color) in enumerate(vals, 1):
            c = ws.cell(row=row, column=col, value=v)
            c.font = Font(name="Arial", size=10, bold=bold, color=color)
            c.fill = fill
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = BORDER
            if fmt:
                c.number_format = fmt

    found = sum(1 for r in results if r.get("status") == "FOUND")
    rfq = len(results) - found
    note_row = len(results) + 3
    llm_rows = sum(1 for r in results if r.get("llm_assisted"))
    note = (f"Найдено: {found} | RFQ: {rfq} | "
            f"Курс USD/RUB (ЦБ + наценка): {rate_rub:.4f} | "
            f"Цена по ценовому брекету ≥ количества, приоритет позиций в наличии. "
            f"Источник цен — только официальные API дистрибьюторов "
            f"(DigiKey, Mouser, TME, Newark/Farnell).")
    if llm_rows:
        note += (f" Позиций, где написание артикула уточнено автоматически и затем "
                 f"подтверждено у дистрибьютора: {llm_rows} — выделены персиковым, "
                 f"рекомендуется выборочная проверка.")
    nc = ws.cell(row=note_row, column=1, value=note)
    nc.font = Font(name="Arial", size=9, color="595959")
    nc.alignment = Alignment(horizontal="left", wrap_text=True)
    ws.merge_cells(f"A{note_row}:J{note_row}")
    ws.row_dimensions[note_row].height = 30

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f"\nГотово: {output_path}")
    print(f"Найдено: {found} | RFQ: {rfq}")


# ───────────────────────────────────────────────
# MAIN
# ───────────────────────────────────────────────

def load_items(args):
    if args.once:
        return [{"num": 1, "pn": args.once[0], "qty": int(args.once[1]),
                 "manufacturer": args.manufacturer or "", "description": args.once[0]}]
    if args.input:
        path = Path(args.input)
        if not path.is_absolute():
            path = SCRIPT_DIR / path
    else:
        path = autopick_request()
        if path is None:
            txt = SCRIPT_DIR / "input.txt"
            if txt.exists():
                print(f"Excel-запрос не найден, читаю {txt.name}")
                return read_txt(txt)
            print("Не найден ни Excel-запрос, ни input.txt. Положите файл запроса рядом со скриптом.")
            sys.exit(1)
        print(f"Авто-выбран запрос: {path.name}")
    if not path.exists():
        print(f"Файл не найден: {path}")
        sys.exit(1)
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        return read_request_xlsx(path)
    return read_txt(path)


def main():
    ap = argparse.ArgumentParser(description="Парсер цен ЭКБ (DigiKey + Mouser + Newark/Farnell API)")
    ap.add_argument("--input", "-i", default=None, help="Excel/txt запрос. По умолч. авто-выбор")
    ap.add_argument("--output", "-o", default=None)
    ap.add_argument("--once", nargs=2, metavar=("PN", "QTY"), help="Одна позиция для проверки")
    ap.add_argument("--manufacturer", default="", help="Производитель для режима --once")
    ap.add_argument("--rub-rate", type=float, default=None,
                    help="Курс USD/RUB вручную (иначе берётся с ЦБ)")
    ap.add_argument("--rub-markup", type=float, default=4.0,
                    help="Наценка к курсу USD/RUB (по умолч. +4%%)")
    ap.add_argument("--preview-only", action="store_true",
                    help="Только распознать в превью, без поиска цен")
    ap.add_argument("--from-preview", action="store_true",
                    help="Сразу читать превью_позиции.xlsx")
    ap.add_argument("--preview-path", type=Path, default=None,
                    help="Путь к превью; нужен для изолированной обработки очереди")
    ap.add_argument("--yes", "-y", action="store_true", help="Не спрашивать подтверждение")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    preview_path = args.preview_path or (SCRIPT_DIR / PREVIEW_NAME)

    # 1) получить позиции
    if args.from_preview and preview_path.exists():
        items = read_preview(preview_path)
        print(f"Читаю из превью: {preview_path.name}")
    else:
        items = load_items(args)

    if args.limit:
        items = items[:args.limit]

    # 2) превью + чекпоинт
    if not args.once and not args.from_preview:
        write_preview(items, preview_path)
        print(f"\nРаспознано позиций: {len(items)}. Превью: {preview_path}")
        for it in items[:30]:
            flag = "" if it["pn"] else "  <-- ПУСТО, впишите артикул!"
            print(f"  #{it['num']:>3}  {it['pn'] or '(нет)':<28} x{it['qty']:<6} "
                  f"{it.get('manufacturer', '')}{flag}")
        if len(items) > 30:
            print(f"  ... ещё {len(items) - 30}")
        if args.preview_only:
            print("\nРежим превью: файл сохранён, цены не искались.")
            return
        if not args.yes:
            print(f"\nПроверьте/поправьте {PREVIEW_NAME} в Excel и СОХРАНИТЕ его.")
            ans = input("Нажмите Enter для поиска цен, или N для выхода: ").strip().lower()
            if ans == "n":
                print("Остановлено. Поправьте превью и запустите снова.")
                return
            if preview_path.exists():
                items = read_preview(preview_path)

    # 3) курс
    rate_rub = get_usd_rub(markup=args.rub_markup, manual=args.rub_rate)

    # 4) наличие API-ключей
    have_keys = [name for name, val in (
        ("DigiKey", ENV.get("DIGIKEY_CLIENT_ID")),
        ("Mouser", ENV.get("MOUSER_API_KEY")),
        ("Newark/Farnell", ENV.get("FARNELL_API_KEY")),
        ("TME", ENV.get("TME_TOKEN")),
    ) if val]
    if not have_keys:
        print("\nВНИМАНИЕ: в .env не найдено ни одного API-ключа. "
              "Все позиции уйдут в RFQ. Проверьте файл .env рядом со скриптом.")
    else:
        print(f"\nАктивные источники: {', '.join(have_keys)}")

    # 5) поиск цен
    print(f"USD/RUB: {rate_rub:.4f}")
    print(f"Позиций к поиску: {len(items)}")
    out_path = Path(args.output) if args.output else (SCRIPT_DIR / "final" / "BOM_Приложение_1.xlsx")
    results = []
    for i, it in enumerate(items, 1):
        pn = it.get("pn")
        qty = it.get("qty", 1)
        num = it.get("num", i)
        manufacturer = it.get("manufacturer", "")
        print(f"\n[{i}/{len(items)}] #{num}: {pn or '(нет артикула)'}")
        if not pn or not manufacturer:
            results.append({
                "pn": pn or "", "qty": qty,
                "manufacturer": manufacturer,
                "description": it.get("description", ""),
                "status": "RFQ",
            })
            if pn and not manufacturer:
                print("    -> RFQ: производитель не указан")
            continue
        try:
            r = find_best_price(pn, qty, manufacturer)
        except Exception as e:
            print(f"    Ошибка: {e}")
            r = {"pn": pn, "qty": qty, "status": "RFQ"}
        r["description"] = it.get("description", pn)
        if not r.get("manufacturer"):
            r["manufacturer"] = it.get("manufacturer", "")
        if r.get("status") == "FOUND":
            tag = "В наличии" if r.get("in_stock") else "под заказ"
            print(f"    -> {r['distributor']}  ${r['price_usd']:.3f}"
                  f"  ({tag}, склад {r['stock']}, MOQ {r['moq']})")
        else:
            print("    -> RFQ")
        results.append(r)

    # Добивочный проход по RFQ-остатку через oemsecrets (лимитированная квота)
    try:
        oemsecrets_pass(results)
    except Exception as e:
        print(f"oemsecrets-фолбэк пропущен: {e}")

    # LLM-контур: подсказка артикула -> обязательная проверка через API
    try:
        claude_assisted_pass(results)
    except Exception as e:
        print(f"LLM-контур пропущен: {e}")

    write_results(results, out_path, rate_rub)


if __name__ == "__main__":
    main()
