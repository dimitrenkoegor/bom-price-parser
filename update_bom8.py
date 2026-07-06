"""
Обновляет final/BOM_Приложение_№1_(4).xlsx найденными ценами (9 позиций: аттенюаторы, резонаторы/осцилляторы).
Источник: start/Приложение №1.xlsx
Запускать: python update_bom8.py
"""
import urllib.request, json, re
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT = SCRIPT_DIR / "final" / "BOM_Приложение_№1_(4).xlsx"

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

# ── Данные (собраны с oemsecrets через Chrome, 2026-07-02) ────────────────────

RESULTS = [
    {"pn": "YAT-12A+", "qty": 600, "manufacturer": "Mini-Circuits",
     "distributor": "Mouser", "moq": 1, "stock": 1622, "lead": None,
     "price_eur": 3.70, "status": "FOUND"},
    {"pn": "YAT-15A+", "qty": 3000, "manufacturer": "Mini-Circuits",
     "distributor": "Mouser", "moq": 1, "stock": 918, "lead": None,
     "price_eur": 3.70, "status": "FOUND"},
    {"pn": "KSE-6N50000M3C045ZA3", "qty": 1300, "manufacturer": "Keenside Electronics", "status": "RFQ"},
    {"pn": "SG-8002CA-PCM 50.0 MHz", "qty": 1300, "manufacturer": "Golledge (Epson)", "status": "RFQ"},
    {"pn": "LV0750000XMBDRX", "qty": 3600, "manufacturer": "PDI", "status": "RFQ"},
    {"pn": "CL7050-50.000-3.3-25-X-T-TR", "qty": 3600, "manufacturer": "Raltron Electronics",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 4.1610, "status": "FOUND"},
    {"pn": "511BBA50M0000AAGR", "qty": 3600, "manufacturer": "Skyworks Solutions",
     "distributor": "Mouser", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 4.03, "status": "FOUND"},
    {"pn": "XLL736050.0000001", "qty": 3600, "manufacturer": "Renesas Electronics",
     "distributor": "Avnet America", "moq": 1, "stock": 0, "lead": None,
     "bpn": "XLL736050.000000I", "price_eur": 2.5748, "status": "FOUND"},
    {"pn": "SiT9121AI-1DF-33E100.000000T", "qty": 3000, "manufacturer": "YXC (SiTime-compatible)",
     "distributor": "Avnet America", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 3.2697, "status": "FOUND"},
]

# ── Стили ─────────────────────────────────────────────────────────────────────

F_HDR  = PatternFill("solid", fgColor="1F4E79")
F_HDRB = PatternFill("solid", fgColor="375623")
F_R1   = PatternFill("solid", fgColor="EBF3FB")
F_R2   = PatternFill("solid", fgColor="FFFFFF")
F_RFQ  = PatternFill("solid", fgColor="FFF2CC")
THIN   = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

HEADERS = [
    ("Part Number",              32, F_HDR),
    ("Distributor Part Number",  26, F_HDRB),
    ("Qty for Single BOM",       12, F_HDR),
    ("Manufacturer",             22, F_HDR),
    ("Distributor",              16, F_HDR),
    ("Minimum Order",            11, F_HDR),
    ("Stock",                    11, F_HDR),
    ("Lead Time (weeks)",        13, F_HDR),
    ("Unit Price USD",           14, F_HDR),
    ("Unit Price RUB",           15, F_HDR),
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
            lead  = r.get("lead")
            in_stock = stock > 0 and lead is None

            if in_stock:
                lead_val = "In stock"
            elif lead is not None:
                lead_val = lead
            else:
                lead_val = ""

            if "price_usd" in r:
                usd = r["price_usd"]
            else:
                usd = round(r["price_eur"] * rate_eur, 5)
            rub = round(usd * rate_rub, 2)

            write_cell(ws, row, 1, pn,                   base)
            write_cell(ws, row, 2, r.get("bpn"),          base)
            write_cell(ws, row, 3, r["qty"],              base)
            write_cell(ws, row, 4, r.get("manufacturer"), base)
            write_cell(ws, row, 5, r.get("distributor"),  base)
            write_cell(ws, row, 6, r.get("moq"),          base)
            write_cell(ws, row, 7, stock,                 base)
            write_cell(ws, row, 8, lead_val,              base)
            write_cell(ws, row, 9, usd,                   base, "#,##0.0000")
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
        f"Цена по ценовому брекету ≥ закупаемого кол-ва, минимум среди авторизованных. "
        f"Авторизованные: DigiKey, Mouser, Arrow, TTI, TME, Avnet, Newark/Farnell, Future. "
        f"Позиция 7 (XLL736050.0000001): в запросе PN оканчивается на '1', корректный "
        f"символ — латинская 'I' (см. Distributor Part Number)."
    )
    nc = ws.cell(row=note_row, column=1, value=note)
    nc.font = Font(name="Arial", size=9, color="595959")
    nc.alignment = Alignment(horizontal="left", wrap_text=True)
    ws.merge_cells(f"A{note_row}:J{note_row}")
    ws.row_dimensions[note_row].height = 55

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT)
    print(f"\nГотово: {OUTPUT}")
    print(f"Найдено: {found} | RFQ: {rfq}")
    print(f"EUR/USD: {rate_eur:.4f}  |  USD/RUB: {rate_rub:.4f}")

if __name__ == "__main__":
    print("Получаю актуальные курсы валют...")
    rate_eur = get_eur_usd()
    rate_rub = get_usd_rub()
    print(f"EUR/USD: {rate_eur:.4f}  |  USD/RUB: {rate_rub:.4f}")
    build_bom(rate_eur, rate_rub)
