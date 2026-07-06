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
]
# Известные производители — отбрасываем из хвоста наименования (и пишем в Manufacturer)
KNOWN_MFR = [
    "Traco Power", "TracoPower", "Traco", "Murata", "Vishay", "TDK", "Bourns",
    "Panasonic", "Rohm", "Littelfuse", "Yageo", "Kemet", "AVX", "Nichicon",
    "Wurth", "Würth", "TE Connectivity", "Molex", "Samsung", "Kyocera",
    "Knowles", "Voltronics", "Dalicap", "Hitano", "Jamicon", "Epcos",
]


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
        with urllib.request.urlopen(req, timeout=15) as r:
            d = json.loads(r.read().decode())
        _digikey_token_cache["token"] = d["access_token"]
        _digikey_token_cache["expires_at"] = now + int(d.get("expires_in", 600))
        return d["access_token"]
    except Exception as e:
        print(f"    DigiKey OAuth error: {e}")
        return None


def search_digikey(pn, qty):
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
        with urllib.request.urlopen(req, timeout=15) as r:
            d = json.loads(r.read().decode())
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
        if mpn.upper() != pn.strip().upper():
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
        offers.append({
            "dist": "DigiKey", "pn": p.get("ProductVariations", [{}])[0].get("DigiKeyProductNumber", "")
                    if p.get("ProductVariations") else (p.get("DigiKeyPartNumber") or ""),
            "stock": stock, "moq": moq, "lead": None, "tiers": tiers, "manu": manu,
        })
    return offers


# ─────────────────────────────────────────────
# MOUSER  (Search API)
# ─────────────────────────────────────────────

def search_mouser(pn, qty):
    key = ENV.get("MOUSER_API_KEY")
    if not key:
        return []
    url = "https://api.mouser.com/api/v1.0/search/partnumber?apiKey=" + urllib.parse.quote(key)
    body = json.dumps({"SearchByPartRequest": {"mouserPartNumber": pn, "partSearchOptions": ""}}).encode()
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            d = json.loads(r.read().decode())
    except Exception as e:
        print(f"    Mouser error: {e}")
        return []

    errors = d.get("Errors") or []
    if errors:
        # Тихо пропускаем — например "pending authorisation" пока не одобрен ключ
        return []

    results = (d.get("SearchResults") or {}).get("Parts") or []
    offers = []
    for p in results:
        mpn = (p.get("ManufacturerPartNumber") or "").strip()
        if mpn.upper() != pn.strip().upper():
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
        offers.append({
            "dist": "Mouser", "pn": p.get("MouserPartNumber") or "",
            "stock": stock, "moq": moq, "lead": None, "tiers": tiers,
            "manu": p.get("Manufacturer") or "",
        })
    return offers


# ─────────────────────────────────────────────
# NEWARK / FARNELL (element14 Product Search API)
# ─────────────────────────────────────────────

def search_farnell(pn, qty):
    key = ENV.get("FARNELL_API_KEY")
    if not key:
        return []
    params = (
        "term=manuPartNum%3A" + urllib.parse.quote(pn) +
        "&storeInfo.id=www.newark.com"
        "&resultsSettings.offset=0"
        "&resultsSettings.numberOfResults=5"
        "&resultsSettings.responseGroup=large"
        "&callInfo.responseDataFormat=JSON"
        "&callinfo.apiKey=" + key
    )
    url = "https://api.element14.com/catalog/products?" + params
    req = urllib.request.Request(url)
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            d = json.loads(r.read().decode())
    except Exception as e:
        print(f"    Newark/Farnell error: {e}")
        return []

    products = (d.get("manufacturerPartNumberSearchReturn") or {}).get("products") or []
    offers = []
    for p in products:
        mpn = (p.get("translatedManufacturerPartNumber") or "").strip()
        if mpn.upper() != pn.strip().upper():
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
            "dist": "Newark/Farnell", "pn": p.get("sku") or "",
            "stock": stock, "moq": moq,
            "lead": int(lead) if lead else None,
            "tiers": tiers, "manu": p.get("brandName") or "",
        })
    return offers


# ─────────────────────────────────────────────
# ВЫБОР ЛУЧШЕЙ ЦЕНЫ
# ─────────────────────────────────────────────

def find_best_price(pn, qty):
    print(f"    qty={qty}")
    offers = []
    for search_fn in (search_digikey, search_mouser, search_farnell):
        try:
            offers.extend(search_fn(pn, qty))
        except Exception as e:
            print(f"    {search_fn.__name__} ошибка: {e}")

    if not offers:
        return {"pn": pn, "qty": qty, "status": "RFQ"}

    best = None
    for o in offers:
        tier, min_not_met = pick_tier(o["tiers"], qty)
        price_usd = tier["price"]
        stock = o.get("stock", 0)
        lead = o.get("lead")
        in_stock = stock > 0 and lead is None
        cand = {
            "pn": pn, "distr_pn": o.get("pn", ""),
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
# ИЗВЛЕЧЕНИЕ АРТИКУЛА ИЗ ОПИСАНИЯ
# ─────────────────────────────────────────────

def split_manufacturer(desc):
    """Возвращает (описание_без_производителя, производитель|None)."""
    for mfr in sorted(KNOWN_MFR, key=len, reverse=True):
        m = re.search(re.escape(mfr) + r"\s*$", desc, re.I)
        if m:
            return desc[:m.start()].strip(" -,"), mfr
    return desc, None


def extract_pn_from_description(desc):
    """
    Извлекает партномер из текстового наименования.
    Правила:
    - Убираем ведущее типовое слово (КОНДЕНСАТОР, РЕЗИСТОР и т.п.)
    - Убираем производителя из хвоста
    - Из оставшихся токенов берём первый, похожий на артикул:
      содержит буквы+цифры, без "/" и "(", без кириллицы
    - Токены со спецсимволами "/" или "," (спецификации: NM-0,5/8) отбрасываем
    """
    if not desc:
        return None
    body, _ = split_manufacturer(desc.strip())
    low = body.lower()
    for w in sorted(RU_TYPE_WORDS, key=len, reverse=True):
        if low.startswith(w):
            body = body[len(w):].strip(" -:")
            break

    pn_candidates = []
    spec_candidates = []  # токены выглядят как спецификации — запасной вариант

    for t in body.split():
        t = t.strip("();,")
        if not t:
            continue
        if re.search(r"[а-яёА-ЯЁ]", t):
            continue                              # кириллица — пропуск
        if len(t) < 3:
            continue
        if re.fullmatch(r"[\d.,+\-±%]+", t):
            continue                              # чистое число/процент — пропуск
        if "/" in t or ("," in t and re.search(r"\d,\d", t)):
            spec_candidates.append(t)            # спецификация (NM-0,5/8ПФ)
            continue
        if t.startswith("(") or t.endswith(")"):
            t = t.strip("()")
        if re.search(r"[A-Za-z]", t) and re.search(r"\d", t):
            pn_candidates.append(t)              # содержит буквы И цифры → PN
        elif re.fullmatch(r"[A-Z]{2,}", t):
            pn_candidates.append(t)              # аббревиатура заглавными

    if pn_candidates:
        return pn_candidates[0]                  # первый подходящий токен
    if spec_candidates:
        return None                              # только спецификации — нет PN
    return None


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
        if (("наименован" in texts or "part number" in texts or "условное обознач" in texts)
                and ("количеств" in texts or "qty" in texts or "quantity" in texts)):
            header_row = r
            headers = [(c, v) for c, v in cells if v]
            break
    if header_row is None:
        raise ValueError("Не нашёл строку-шапку в Excel. Проверьте файл или подайте input.txt.")

    col_num = _find(headers, "№", "n п/п", "п/п")
    col_pn = _find(headers, "условное обознач", "part number", "артикул", "парт")
    col_desc = _find(headers, "наименован", "товар")
    col_qty = _find(headers, "количеств", "qty", "quantity", "кол-во")
    col_mfr = _find(headers, "производител", "manufacturer", "изготовит")

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
        # производитель: из колонки, иначе из хвоста наименования
        mfr = mfr_cell
        if not mfr and desc:
            _, m = split_manufacturer(desc)
            mfr = m or ""
        # PN: из колонки, иначе из описания
        pn = pn_cell or (extract_pn_from_description(desc) or "")
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
        items.append({"num": num or len(items) + 1, "pn": str(pn).strip() if pn else "",
                      "qty": qty, "manufacturer": str(mfr).strip() if mfr else "",
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
        pn = r.get("pn") or r.get("description") or ""
        if r.get("status") == "FOUND":
            base = F_R1 if found_idx % 2 == 0 else F_R2
            found_idx += 1
            dist_pn = r.get("distr_pn") or ""
            b_val = dist_pn if packaging_only_suffix(pn, dist_pn) else ""
            lead = r.get("lead")
            lead_val = "In stock" if (r.get("in_stock") and not lead) else (lead if lead is not None else "")
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
    note = (f"Найдено: {found} | RFQ: {rfq} | "
            f"Курс USD/RUB (ЦБ + наценка): {rate_rub:.4f} | "
            f"Цена по ценовому брекету. Источник: официальные API дистрибьюторов "
            f"(DigiKey Product Information API v4, Newark/Farnell Product Search API, "
            f"Mouser Search API). Авторизованные: DigiKey, Mouser, Newark/Farnell.")
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
                 "manufacturer": "", "description": args.once[0]}]
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
    ap.add_argument("--rub-rate", type=float, default=None,
                    help="Курс USD/RUB вручную (иначе берётся с ЦБ)")
    ap.add_argument("--rub-markup", type=float, default=4.0,
                    help="Наценка к курсу USD/RUB (по умолч. +4%%)")
    ap.add_argument("--preview-only", action="store_true",
                    help="Только распознать в превью, без поиска цен")
    ap.add_argument("--from-preview", action="store_true",
                    help="Сразу читать превью_позиции.xlsx")
    ap.add_argument("--yes", "-y", action="store_true", help="Не спрашивать подтверждение")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    preview_path = SCRIPT_DIR / PREVIEW_NAME

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
        print(f"\n[{i}/{len(items)}] #{num}: {pn or '(нет артикула)'}")
        if not pn:
            results.append({
                "pn": "", "qty": qty,
                "manufacturer": it.get("manufacturer", ""),
                "description": it.get("description", ""),
                "status": "RFQ",
            })
            continue
        try:
            r = find_best_price(pn, qty)
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

    write_results(results, out_path, rate_rub)


if __name__ == "__main__":
    main()
