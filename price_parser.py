"""
Парсер цен на ЭКБ (oemsecrets.com)  — версия 3
================================================
Поток работы:
  1. Берёт Excel-запрос (например "149 торги.xlsx") — сам находит строку-шапку
     и колонки (Наименование / Условное обозначение / Количество / Производитель),
     вытаскивает артикул из описания.
  2. Пишет ПРЕВЬЮ  превью_позиции.xlsx  и показывает распознанное в консоли.
     Вы проверяете/правите его в Excel, сохраняете — и скрипт продолжит уже по
     вашим правкам.
  3. Ищет цены на oemsecrets (обход Cloudflare через undetected-chromedriver),
     выбирает брекет под количество, берёт авторизованных дистрибьюторов.
  4. Пишет результат  final/BOM_Приложение_1.xlsx  — 10 колонок:
     Part Number | Distributor Part Number | Qty | Manufacturer | Distributor |
     Minimum Order | Stock | Lead Time | Unit Price USD | Unit Price RUB
     Рублёвая цена = USD * (курс ЦБ USD/RUB на момент запуска + наценка, по умолч. +4).

Установка зависимостей:
    pip install undetected-chromedriver selenium beautifulsoup4 openpyxl webdriver-manager

Запуск:
    python price_parser.py                      # авто-выбор Excel-запроса рядом
    python price_parser.py --input "149 торги.xlsx"
    python price_parser.py --preview-only        # только распознать в превью, без цен
    python price_parser.py --yes                  # не спрашивать подтверждение
    python price_parser.py --once C5750X7R1H106KT 320   # проверка одной позиции
    python price_parser.py --rub-markup 4 --rate 1.16
"""

import argparse
import time
import random
import re
import sys
import urllib.request
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Установите зависимости:\n  pip install undetected-chromedriver selenium beautifulsoup4 openpyxl webdriver-manager")
    sys.exit(1)

SCRIPT_DIR = Path(__file__).resolve().parent
PREVIEW_NAME = "превью_позиции.xlsx"

AUTHORIZED = ["mouser", "digi-key", "digikey", "arrow", "tti", "tme",
              "avnet", "newark", "farnell", "element14", "future"]

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
            import json
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
# WEBDRIVER
# ─────────────────────────────────────────────

def create_driver(headless=False):
    try:
        import undetected_chromedriver as uc
        opts = uc.ChromeOptions()
        opts.add_argument("--window-size=1300,950")
        if headless:
            opts.add_argument("--headless=new")
        driver = uc.Chrome(options=opts)
        print("Драйвер: undetected-chromedriver")
        return driver
    except Exception as e:
        print(f"undetected-chromedriver недоступен ({e.__class__.__name__}), пробую обычный Selenium…")
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    opts = Options()
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument("--window-size=1300,950")
    if headless:
        opts.add_argument("--headless=new")
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        from selenium.webdriver.chrome.service import Service
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    except Exception:
        driver = webdriver.Chrome(options=opts)
    driver.execute_script("Object.defineProperty(navigator,'webdriver',{get:()=>undefined})")
    print("Драйвер: Selenium")
    return driver


# ─────────────────────────────────────────────
# OEMSECRETS
# ─────────────────────────────────────────────

EXTRACT_JS = r"""
return (function(){
  function toFloat(s){
    // европейский формат: 1.234,56 → 1234.56
    s = String(s).trim();
    if(/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)){
      s = s.replace(/\./g,'').replace(',','.');
    } else {
      s = s.replace(',','.');
    }
    return parseFloat(s);
  }
  function toInt(s){ s=String(s).replace(/[^\d]/g,''); return parseInt(s)||0; }
  var rows=[].slice.call(document.querySelectorAll('[data-ps-row-index]'));
  return rows.map(function(r){
    var cells=(r.innerText||'').split(/\n|\t/).map(function(s){return s.trim();}).filter(Boolean);
    var dist=r.getAttribute('data-delog-distributor-common-name')||cells[0]||'';
    var stock=toInt(r.getAttribute('data-delog-quantity-in-stock')||'0');
    var pn=cells[1]||'';
    var moq=1, pkg='';
    for(var i=0;i<cells.length;i++){
      var m=cells[i].match(/^(\d[\d.]*|-)\s+(Bulk|Reel|Tape\s*&?\s*Reel|Tape|Box|Tube|Tray|Each|Cut\s*Tape|Bag|Pack|Ammo|Digi-?Reel|MiniReel)/i);
      if(m){ if(m[1]!=='-') moq=toInt(m[1]); pkg=m[2]; break; }
    }
    var lead=null;
    for(var i=0;i<cells.length;i++){ var lm=cells[i].match(/Lead\s*Time:\s*(\d+)\s*weeks?/i); if(lm){ lead=parseInt(lm[1]); break; } }
    // Тиры цен: oemsecrets выводит чередующиеся ячейки: "qty" затем "€ price"
    // Пример: ["1", "€ 0,3960", "10", "€ 0,2790"]
    var tiers=[];
    var pricePat=/^([€$£])\s*([\d.,]+)$/;
    var intPat=/^[\d.,]+$/;
    for(var i=1;i<cells.length;i++){
      var pm=cells[i].match(pricePat);
      if(pm){
        var price=toFloat(pm[2]);
        if(price<=0||price>9999999) continue;
        // предыдущая ячейка должна быть целым числом (кол-во)
        var prev=cells[i-1];
        if(intPat.test(prev)){
          var q=toInt(prev);
          if(q>0){
            tiers.push({qty:q, cur:pm[1], price:price});
          }
        }
      }
    }
    var manu=''; var p=r;
    for(var k=0;k<14 && p;k++){
      if(p.getAttribute && p.hasAttribute && p.hasAttribute('data-ps-product-index')){
        var mm=(p.innerText||'').match(/\bby\s+([A-Za-z][\w .,&\/\-]{1,40})/);
        if(mm){ manu=mm[1].trim().replace(/\s+(Add to BoM|Get Quote|Datasheet).*$/,'').trim(); }
        break;
      }
      p=p.parentElement;
    }
    return {dist:dist, stock:stock, pn:pn, moq:moq, pkg:pkg, lead:lead, tiers:tiers, manu:manu};
  });
})();
"""


def wait_for_results(driver, timeout=55):
    """
    Ждёт появления строк с ценами ([data-ps-row-index]).
    Увеличенный таймаут — oemsecrets грузит данные через XHR, не сразу.
    """
    deadline = time.time() + timeout
    last = ""
    while time.time() < deadline:
        # Редирект на request-stock = PN не найден
        try:
            url = driver.current_url or ""
            if "request-stock" in url:
                return "empty"
        except Exception:
            pass
        try:
            n = driver.execute_script("return document.querySelectorAll('[data-ps-row-index]').length;")
        except Exception:
            n = 0
        body = ""
        try:
            body = driver.execute_script("return document.body ? document.body.innerText : '';") or ""
        except Exception:
            pass
        if n and n > 0:
            time.sleep(2)          # ждём дозагрузку всех строк
            return "ok"
        if "No Distributor Results Found" in body or "request-stock" in body:
            return "empty"
        if "Just a moment" in body or "Один момент" in body or "Verifying you are human" in body:
            last = "cloudflare"
        time.sleep(2.5)            # пауза между проверками
    return last or "timeout"


def set_quantity(driver, qty):
    try:
        ok = driver.execute_script("""
            var inp=document.querySelector('input[name="requested_quantity"]');
            if(!inp) return false;
            inp.value=arguments[0];
            inp.dispatchEvent(new Event('input',{bubbles:true}));
            inp.dispatchEvent(new Event('change',{bubbles:true}));
            var btn=null, scope=inp.closest('form')||inp.parentElement;
            for(var hop=0; hop<4 && scope; hop++){
              var b=[].slice.call(scope.querySelectorAll('button,a')).find(function(x){return /apply/i.test(x.textContent);});
              if(b){ btn=b; break; } scope=scope.parentElement;
            }
            if(btn){ btn.click(); return true; }
            if(inp.form){ inp.form.submit(); return true; }
            return false;
        """, str(qty))
        if ok:
            time.sleep(3)
            wait_for_results(driver, timeout=25)
        return bool(ok)
    except Exception:
        return False


def parse_oemsecrets(driver, pn, qty):
    driver.get(f"https://www.oemsecrets.com/compare/{pn}")
    time.sleep(random.uniform(10, 13))  # случайная пауза: ждём XHR после навигации
    status = wait_for_results(driver, timeout=55)
    if status == "cloudflare":
        print("    Cloudflare блокирует. Поставьте undetected-chromedriver и не используйте --headless.")
        return []
    if status == "empty":
        print("    Нет результатов на oemsecrets.")
        return []
    if status != "ok":
        print(f"    Таймаут ожидания ({status}).")
        return []
    set_quantity(driver, qty)
    try:
        raw = driver.execute_script(EXTRACT_JS) or []
    except Exception as e:
        print(f"    Ошибка извлечения: {e}")
        return []
    offers = []
    for o in raw:
        dist = (o.get("dist") or "").strip()
        if not dist or not any(a in dist.lower() for a in AUTHORIZED):
            continue
        if not (o.get("tiers") or []):
            continue
        offers.append(o)
    return offers


def pick_tier(tiers, qty):
    ts = sorted(tiers, key=lambda t: t["qty"])
    applicable = [t for t in ts if t["qty"] <= qty]
    if applicable:
        return applicable[-1], False
    return ts[0], True


def to_usd(price, cur, rate_eur):
    if cur == "$":
        return price
    if cur == "£":
        return price * 1.27
    return price * rate_eur


def _parse_octopart_json(data, pn):
    """
    Извлекает предложения дистрибьюторов из Next.js JSON Octopart.
    data = parsed __NEXT_DATA__ dict
    """
    import json

    def deep_find(obj, key):
        """Рекурсивно ищет все значения ключа в произвольной структуре."""
        results = []
        if isinstance(obj, dict):
            if key in obj:
                results.append(obj[key])
            for v in obj.values():
                results.extend(deep_find(v, key))
        elif isinstance(obj, list):
            for item in obj:
                results.extend(deep_find(item, key))
        return results

    # Octopart Next.js: offers находятся в dehydratedState → queries → data → results → parts → sellers
    offers = []
    sellers_lists = deep_find(data, "sellers")
    for sellers in sellers_lists:
        if not isinstance(sellers, list):
            continue
        for seller in sellers:
            company = seller.get("company") or {}
            dist = company.get("name") or seller.get("name") or ""
            if not any(a in dist.lower() for a in AUTHORIZED):
                continue
            offers_raw = seller.get("offers") or []
            for off in offers_raw:
                prices_raw = off.get("prices") or {}
                # prices_raw может быть dict {currency: [[qty, price], ...]} или list
                tiers = []
                if isinstance(prices_raw, dict):
                    usd_tiers = prices_raw.get("USD") or prices_raw.get("usd") or next(iter(prices_raw.values()), [])
                    for item in usd_tiers:
                        if isinstance(item, (list, tuple)) and len(item) >= 2:
                            tiers.append({"qty": int(item[0]), "cur": "$", "price": float(item[1])})
                elif isinstance(prices_raw, list):
                    for item in prices_raw:
                        if isinstance(item, dict):
                            qty = item.get("quantity") or item.get("qty") or 1
                            price = item.get("price") or item.get("convertedPrice") or 0
                            cur = item.get("currency") or "$"
                            if price:
                                tiers.append({"qty": int(qty), "cur": cur, "price": float(price)})
                if not tiers:
                    continue
                stock = int(off.get("in_stock_quantity") or off.get("quantity") or 0)
                moq = int(off.get("moq") or off.get("min_order_qty") or 1)
                distr_pn = off.get("sku") or off.get("part_number") or pn
                offers.append({
                    "dist": dist, "pn": distr_pn, "stock": stock,
                    "moq": moq, "lead": None, "tiers": tiers, "manu": "",
                })
    return offers


def _parse_octopart_text(text, pn):
    """
    Резервный текстовый парсер для Octopart.innerText.
    Ищет блоки: ДистрибьюторИмя → PN дистрибьютора → остаток → MOQ → цены.
    """
    DIST_CANONICAL = {
        "mouser": "Mouser", "digi-key": "DigiKey", "digikey": "DigiKey",
        "arrow": "Arrow Electronics", " tti ": "TTI", "tme": "TME",
        "avnet": "Avnet", "newark": "Newark", "farnell": "Farnell",
        "element14": "Newark", "future electronics": "Future Electronics",
    }
    QTY_BREAKS = [1, 10, 25, 100, 250, 1000, 2500, 10000]
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    offers = []
    i = 0
    while i < len(lines):
        line = lines[i]
        matched_dist = next(
            (DIST_CANONICAL[k] for k in DIST_CANONICAL if k.strip() in line.lower()),
            None
        )
        if not matched_dist:
            i += 1
            continue
        distr_pn, stock, moq, prices = "", 0, 1, []
        j = i + 1
        while j < min(i + 25, len(lines)):
            l = lines[j]
            # Партномер дистрибьютора: латиница+цифры, без пробелов
            if not distr_pn and re.fullmatch(r"[A-Za-z0-9\-_.]+", l) and 3 <= len(l) <= 45:
                distr_pn = l
            # Остаток: целое число
            elif re.fullmatch(r"\d{1,8}", l) and not stock:
                stock = int(l)
            # MOQ строка: "1 Bulk USD"
            elif re.match(r"^\d+\s+(Bulk|Reel|Tape|Box|Tube|Tray|Each|Cut)", l, re.I):
                m = re.match(r"^(\d+)", l)
                if m:
                    moq = int(m.group(1))
            # Цена: "12.345" или "1,234.56"
            elif re.fullmatch(r"\d{1,6}(?:[,]\d{3})*[.]\d{2,4}", l):
                val = float(l.replace(",", ""))
                if 0.001 < val < 999999:
                    prices.append(val)
            # Конец блока
            elif any(k.strip() in l.lower() for k in DIST_CANONICAL) and l != line:
                break
            j += 1
        if prices:
            tiers = [{"qty": QTY_BREAKS[k] if k < len(QTY_BREAKS) else QTY_BREAKS[-1],
                      "cur": "$", "price": p} for k, p in enumerate(prices)]
            offers.append({
                "dist": matched_dist, "pn": distr_pn or pn,
                "stock": stock, "moq": moq, "lead": None, "tiers": tiers, "manu": "",
            })
        i = j
    return offers


def parse_octopart(driver, pn, qty):
    """
    Fallback: парсит Octopart через Selenium.
    Сначала пробует __NEXT_DATA__ JSON (надёжно), потом innerText (резервный).
    """
    try:
        driver.get(f"https://octopart.com/search?q={pn}&currency=USD")
    except Exception as e:
        print(f"    Octopart navigate error: {e}")
        return []

    # Случайная задержка 7–11 сек + прокрутка — имитация человека, обход антибота
    time.sleep(random.uniform(10, 13))
    try:
        driver.execute_script("window.scrollTo(0, 400);")
        time.sleep(random.uniform(1.5, 3.0))
        driver.execute_script("window.scrollTo(0, 800);")
        time.sleep(random.uniform(1.0, 2.0))
    except Exception:
        pass

    # Попытка 1: Next.js JSON (структурированные данные)
    try:
        import json
        raw_json = driver.execute_script(
            "var el=document.getElementById('__NEXT_DATA__'); return el ? el.textContent : null;"
        )
        if raw_json:
            data = json.loads(raw_json)
            offers = _parse_octopart_json(data, pn)
            if offers:
                print(f"    Octopart JSON: {len(offers)} предложений")
                return offers
    except Exception as e:
        print(f"    Octopart JSON parse error: {e}")

    # Попытка 2: текстовый парсер
    try:
        text = driver.execute_script("return document.body ? document.body.innerText : '';") or ""
    except Exception:
        return []
    offers = _parse_octopart_text(text, pn)
    if offers:
        print(f"    Octopart текст: {len(offers)} предложений")
    else:
        print(f"    Octopart: ничего не найдено для {pn}")
    return offers


def normalize_pn_variants(pn):
    """
    Возвращает список вариантов PN для поиска на oemsecrets (в порядке приоритета).
    - /NOPB, /N0PB → убираем слэш (oemsecrets редиректит при слэше)
    - суффикс x/X (wildcard упаковки) → пробуем R (reel) и T (tape&reel)
    """
    variants = [pn]
    # /NOPB или /N0PB → убрать слэш
    nopb = re.sub(r'/N[O0]PB$', 'NOPB', pn, flags=re.I)
    if nopb != pn:
        variants.append(nopb)
        # ещё без NOPB совсем (иногда на oemsecrets только базовый PN)
        base_no_nopb = re.sub(r'N[O0]PB$', '', nopb).rstrip('-_')
        if base_no_nopb not in variants:
            variants.append(base_no_nopb)
    # wildcard x → пробуем R и T
    if pn.endswith('x') or pn.endswith('X'):
        base = pn[:-1]
        for suf in ('R', 'T', 'TR'):
            v = base + suf
            if v not in variants:
                variants.append(v)
    return variants


def find_best_price(driver, pn, qty, rate_eur):
    print(f"    qty={qty}")

    # Шаг 1: oemsecrets — пробуем все варианты нормализованного PN
    offers = []
    searched_pn = pn
    for variant in normalize_pn_variants(pn):
        if variant != pn:
            print(f"    Пробую вариант PN: {variant}")
        offers = parse_oemsecrets(driver, variant, qty)
        if offers:
            searched_pn = variant
            break

    # Шаг 2 (fallback): Octopart
    if not offers:
        print(f"    oemsecrets пуст/заблокирован → пробуем Octopart...")
        offers = parse_octopart(driver, pn, qty)

    if not offers:
        return {"pn": pn, "qty": qty, "status": "RFQ"}

    best = None
    for o in offers:
        tier, min_not_met = pick_tier(o["tiers"], qty)
        price_usd = to_usd(tier["price"], tier.get("cur", "€"), rate_eur)
        stock = o.get("stock", 0)
        lead = o.get("lead")
        in_stock = stock > 0 and lead is None
        cand = {
            "pn": pn,
            # Distributor PN: ставим если отличается от запрошенного (суффикс упаковки)
            "distr_pn": o.get("pn", "") if searched_pn != pn else o.get("pn", ""),
            "manufacturer": o.get("manu", ""),
            "distributor": o.get("dist", ""), "moq": o.get("moq", 1),
            "stock": stock, "lead": lead,
            "price_usd": price_usd, "currency": tier.get("cur", "€"),
            "qty": qty, "in_stock": in_stock, "min_not_met": min_not_met, "status": "FOUND",
        }
        # Если PN с суффиксом отличается — записываем найденный вариант в distr_pn
        if searched_pn != pn and not cand["distr_pn"]:
            cand["distr_pn"] = searched_pn
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


def write_results(results, output_path, rate_eur, rate_rub):
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
    note = (f"Найдено: {found} | RFQ: {rfq} | Курс EUR/USD: {rate_eur:.4f} | "
            f"Курс USD/RUB (ЦБ + наценка): {rate_rub:.4f} | "
            f"Цена по ценовому брекету. "
            f"Авторизованные: DigiKey, Mouser, Arrow, TTI, TME, Avnet, Newark/Farnell, Future.")
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
    ap = argparse.ArgumentParser(description="Парсер цен ЭКБ (oemsecrets + Octopart)")
    ap.add_argument("--input", "-i", default=None, help="Excel/txt запрос. По умолч. авто-выбор")
    ap.add_argument("--output", "-o", default=None)
    ap.add_argument("--once", nargs=2, metavar=("PN", "QTY"), help="Одна позиция для проверки")
    ap.add_argument("--rate", type=float, default=None,
                    help="Курс EUR->USD (по умолч. авто с open.er-api.com)")
    ap.add_argument("--rub-rate", type=float, default=None,
                    help="Курс USD/RUB вручную (иначе берётся с ЦБ)")
    ap.add_argument("--rub-markup", type=float, default=4.0,
                    help="Наценка к курсу USD/RUB (по умолч. +4%%)")
    ap.add_argument("--preview-only", action="store_true",
                    help="Только распознать в превью, без поиска цен")
    ap.add_argument("--from-preview", action="store_true",
                    help="Сразу читать превью_позиции.xlsx")
    ap.add_argument("--yes", "-y", action="store_true", help="Не спрашивать подтверждение")
    ap.add_argument("--headless", action="store_true")
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

    # 3) курсы
    rate_eur = args.rate if args.rate else get_eur_usd()
    rate_rub = get_usd_rub(markup=args.rub_markup, manual=args.rub_rate)

    # 4) поиск цен
    print(f"\nКурс EUR/USD: {rate_eur:.4f}  |  USD/RUB: {rate_rub:.4f}")
    print(f"Позиций к поиску: {len(items)}")
    out_path = Path(args.output) if args.output else (SCRIPT_DIR / "final" / "BOM_Приложение_1.xlsx")
    driver = create_driver(headless=args.headless)
    results = []
    try:
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
                r = find_best_price(driver, pn, qty, rate_eur)
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
            import time as _t; _t.sleep(2)
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    write_results(results, out_path, rate_eur, rate_rub)


if __name__ == "__main__":
    main()
