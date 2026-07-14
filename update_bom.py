"""
Обновляет final/BOM_Приложение_1.xlsx найденными ценами.
Запускать: python update_bom.py
"""
import urllib.request, json, re
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT = SCRIPT_DIR / "final" / "BOM_Приложение_1.xlsx"

# ── Курсы ──────────────────────────────────────────────────────────────────────

def get_eur_usd():
    try:
        req = urllib.request.Request("https://open.er-api.com/v6/latest/EUR",
                                     headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as r:
            return float(json.loads(r.read().decode())["rates"]["USD"])
    except Exception:
        return 1.1362

def get_usd_rub(markup=4.0):
    try:
        req = urllib.request.Request("https://www.cbr.ru/scripts/XML_daily.asp",
                                     headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as r:
            xml = r.read().decode("cp1251", "replace")
        m = re.search(r'ID="R01235".*?<Nominal>(\d+)</Nominal>.*?<Value>([\d,]+)</Value>', xml, re.S)
        if m:
            return float(m.group(2).replace(",", ".")) / int(m.group(1)) + markup
    except Exception:
        pass
    return 79.6347 + markup

# ── Данные (собраны вручную с DigiKey, 2026-06-26) ────────────────────────────
# Ключи: pn, qty, manufacturer, distributor, moq, stock, lead (None=In stock), price_usd, status

RESULTS = [
    # 1
    {"pn": "C 09 03 164 6921",  "qty": 1300, "manufacturer": "Harting",      "distributor": "DigiKey", "moq": 1,   "stock": 6185,  "lead": None, "price_usd": 2.49970, "status": "FOUND"},
    # 2
    {"pn": "C 09 03 196 6921",  "qty": 550,  "manufacturer": "Harting",      "distributor": "DigiKey", "moq": 1,   "stock": 7251,  "lead": None, "price_usd": 3.60750, "status": "FOUND"},
    # 3
    {"pn": "2B 09 22 132 6921", "qty": 250,  "manufacturer": "Harting",      "distributor": "DigiKey", "moq": 1,   "stock": 1385,  "lead": None, "price_usd": 3.80400, "status": "FOUND"},
    # 4
    {"pn": "2C 09 23 148 6921", "qty": 400,  "manufacturer": "Harting",      "distributor": "DigiKey", "moq": 1,   "stock": 854,   "lead": None, "price_usd": 3.06485, "status": "FOUND"},
    # 5 – Arrow: 1083 в наличии, цена при брекете 100+
    {"pn": "64 02 01 160 2101", "qty": 250,  "manufacturer": "Harting",      "distributor": "Arrow",   "moq": 1,   "stock": 1083,  "lead": None, "price_usd": 22.83700,"status": "FOUND"},
    # 6
    {"pn": "16FMZ-BT",          "qty": 700,  "manufacturer": "JST",          "distributor": "DigiKey", "moq": 780, "stock": 0,     "lead": 16,   "price_usd": 0.74601, "status": "FOUND"},
    # 7
    {"pn": "125CTG-B-G-R",      "qty": 7000, "manufacturer": "ADAM TECH",    "distributor": "DigiKey", "moq": 1,   "stock": 12234, "lead": None, "price_usd": 0.04648, "status": "FOUND"},
    # 8 – PN исправлен: кириллическая В → латинская B
    {"pn": "B 09 03 264 6825",  "qty": 1400, "manufacturer": "Harting",      "distributor": "DigiKey", "moq": 1,   "stock": 3800,  "lead": None, "price_usd": 2.18000, "status": "FOUND"},
    # 9
    {"pn": "C 09 03 296 6825",  "qty": 650,  "manufacturer": "Harting",      "distributor": "DigiKey", "moq": 1,   "stock": 6074,  "lead": None, "price_usd": 2.96196, "status": "FOUND"},
    # 10
    {"pn": "2B 09 22 232 6825", "qty": 250,  "manufacturer": "Harting",      "distributor": "DigiKey", "moq": 1,   "stock": 1680,  "lead": None, "price_usd": 4.27215, "status": "FOUND"},
    # 11
    {"pn": "2C 09 23 248 6824", "qty": 400,  "manufacturer": "Harting",      "distributor": "DigiKey", "moq": 1,   "stock": 2563,  "lead": None, "price_usd": 3.48777, "status": "FOUND"},
    # 12
    {"pn": "64 02 02 160 2201", "qty": 250,  "manufacturer": "Harting",      "distributor": "DigiKey", "moq": 1,   "stock": 1291,  "lead": None, "price_usd": 18.65550,"status": "FOUND"},
    # 13
    {"pn": "09 45 551 1102",    "qty": 100,  "manufacturer": "Harting",      "distributor": "DigiKey", "moq": 1,   "stock": 893,   "lead": None, "price_usd": 2.60280, "status": "FOUND"},
    # 14 – PN исправлен: 09 45 551 1102 → 292304-2 (TE Connectivity)
    {"pn": "292304-2",          "qty": 500,  "manufacturer": "TE Connectivity","distributor": "DigiKey","moq": 1,  "stock": 1079,  "lead": None, "price_usd": 2.10000, "status": "FOUND"},
    # 15
    {"pn": "IDMD-10-T-05.90-G","qty": 400,  "manufacturer": "Samtec",        "distributor": "DigiKey", "moq": 1,   "stock": 45,    "lead": None, "price_usd": 6.22500, "status": "FOUND"},
    # 16
    {"pn": "125CH-G-04",        "qty": 600,  "manufacturer": "ADAM TECH",    "distributor": "DigiKey", "moq": 1,   "stock": 2156,  "lead": None, "price_usd": 0.21100, "status": "FOUND"},
    # 17
    {"pn": "125CH-G-06",        "qty": 600,  "manufacturer": "ADAM TECH",    "distributor": "DigiKey", "moq": 1,   "stock": 371,   "lead": None, "price_usd": 0.27000, "status": "FOUND"},
    # 18 – нет на DigiKey (только GR-вариант, TR не найден)
    {"pn": "125SH-G-04-TR-SMT-T/R","qty": 600,"manufacturer": "ADAM TECH",  "status": "RFQ"},
    # 19
    {"pn": "125SH-G-06-TR-SMT-T/R","qty": 600,"manufacturer": "ADAM TECH",  "distributor": "DigiKey", "moq": 1,   "stock": 1734,  "lead": None, "price_usd": 0.43400, "status": "FOUND"},
]

# ── Стили ─────────────────────────────────────────────────────────────────────

F_HDR  = PatternFill("solid", fgColor="1F4E79")
F_HDRB = PatternFill("solid", fgColor="375623")
F_R1   = PatternFill("solid", fgColor="EBF3FB")
F_R2   = PatternFill("solid", fgColor="FFFFFF")
F_RFQ  = PatternFill("solid", fgColor="FFF2CC")
F_BCOL = PatternFill("solid", fgColor="E2EFDA")
THIN   = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

HEADERS = [
    ("Part Number",          30, F_HDR),
    ("Distributor Part Number", 26, F_HDRB),
    ("Qty for Single BOM",  11, F_HDR),
    ("Manufacturer",         20, F_HDR),
    ("Distributor",          20, F_HDR),
    ("Minimum Order",        11, F_HDR),
    ("Stock",                10, F_HDR),
    ("Lead Time (weeks)",    13, F_HDR),
    ("Unit Price USD",       14, F_HDR),
    ("Unit Price RUB",       15, F_HDR),
]

def write_cell(ws, row, col, value, fill, fmt=None, bold=False, color="000000"):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Arial", size=10, bold=bold, color=color)
    c.fill = fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = BORDER
    if fmt:
        c.number_format = fmt

def build_bom(rate_eur, rate_rub):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM"

    # Заголовок
    for col, (h, w, fill) in enumerate(HEADERS, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 32

    found_idx = 0
    for i, r in enumerate(RESULTS):
        row = i + 2
        ws.row_dimensions[row].height = 16
        pn = r["pn"]

        if r["status"] == "FOUND":
            base = F_R1 if found_idx % 2 == 0 else F_R2
            found_idx += 1
            stock = r.get("stock", 0)
            qty   = r["qty"]
            lead  = r.get("lead")
            in_stock = stock > 0 and lead is None

            if in_stock:
                lead_val = "In stock"
            elif lead is not None:
                lead_val = f"{lead} weeks" if lead != 1 else "1 week"
            else:
                lead_val = "RFQ" if stock == 0 else ""

            usd = r["price_usd"]
            rub = round(usd * rate_rub, 2)

            write_cell(ws, row, 1, pn,                   base)
            write_cell(ws, row, 2, None,                  base)
            write_cell(ws, row, 3, qty,                   base)
            write_cell(ws, row, 4, r.get("manufacturer"), base)
            write_cell(ws, row, 5, r.get("distributor"),  base)
            write_cell(ws, row, 6, r.get("moq"),          base)
            write_cell(ws, row, 7, stock,                 base)
            write_cell(ws, row, 8, lead_val,              base)
            write_cell(ws, row, 9, usd,                   base, "#,##0.000")
            write_cell(ws, row,10, rub,                   base, "#,##0.00")
        else:
            for col in range(1, 11):
                v = None
                if col == 1: v = pn
                elif col == 3: v = r.get("qty")
                elif col == 4: v = r.get("manufacturer")
                elif col in (9, 10): v = "RFQ"
                bold = col in (9, 10)
                color = "7F4B00" if bold else "000000"
                write_cell(ws, row, col, v, F_RFQ, bold=bold, color=color)

    found = sum(1 for r in RESULTS if r["status"] == "FOUND")
    rfq   = len(RESULTS) - found
    note_row = len(RESULTS) + 3
    note = (
        f"Найдено: {found} | RFQ: {rfq} | Курс EUR/USD: {rate_eur:.4f} | "
        f"Курс USD/RUB (ЦБ + наценка): {rate_rub:.4f} | "
        f"Цена по ценовому брекету ≥ запрошенного кол-ва. "
        f"Авторизованные: DigiKey, Arrow. "
        f"Позиция 18 (125SH-G-04-TR-SMT-T/R): TR-вариант отсутствует у авторизованных дистрибьюторов — RFQ."
    )
    nc = ws.cell(row=note_row, column=1, value=note)
    nc.font = Font(name="Arial", size=9, color="595959")
    nc.alignment = Alignment(horizontal="left", wrap_text=True)
    ws.merge_cells(f"A{note_row}:J{note_row}")
    ws.row_dimensions[note_row].height = 45

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT)
    print(f"\nГотово: {OUTPUT}")
    print(f"Найдено: {found} | RFQ: {rfq}")
    print(f"Курс EUR/USD: {rate_eur:.4f}  |  USD/RUB: {rate_rub:.4f}")

if __name__ == "__main__":
    print("Получаю актуальные курсы валют...")
    rate_eur = get_eur_usd()
    rate_rub = get_usd_rub()
    print(f"EUR/USD: {rate_eur:.4f}  |  USD/RUB: {rate_rub:.4f}")
    build_bom(rate_eur, rate_rub)
