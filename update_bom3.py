"""
Обновляет final/BOM_Приложение_1.xlsx найденными ценами (25 IC позиций).
Запускать: python update_bom3.py
"""
import urllib.request, json, re
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT = SCRIPT_DIR / "final" / "BOM_Приложение_1.xlsx"

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

# ── Данные (собраны с oemsecrets через Chrome, 2026-06-29) ────────────────────
# bpn: Distributor Part Number (если отличается от BOM PN, напр. суффикс упаковки)
# price_eur OR price_usd
# lead: None = In stock, int = недели

RESULTS = [
    # 1 – Macronix MX25L25635FZ2I-10G, qty 30
    # DigiKey: 5245 в наличии, €7.671@10+
    {"pn": "MX25L25635FZ2I-10G",    "qty": 30,  "manufacturer": "Macronix International",
     "distributor": "DigiKey", "moq": 1, "stock": 5245,  "lead": None,
     "price_eur": 7.671,  "status": "FOUND"},

    # 2 – TI INA199B2DCKx, qty 30 (x = упаковка, найден суффикс R = рулон)
    # Mouser: 7214 в наличии, €0.359@10+
    {"pn": "INA199B2DCKx",           "qty": 30,  "manufacturer": "Texas Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 7214,  "lead": None,
     "bpn": "INA199B2DCKR",
     "price_eur": 0.359,  "status": "FOUND"},

    # 3 – STMicro LIS3DHTR, qty 20
    # Mouser: 17606 в наличии, €1.57@5+
    {"pn": "LIS3DHTR",               "qty": 20,  "manufacturer": "STMicroelectronics",
     "distributor": "Mouser", "moq": 1, "stock": 17606, "lead": None,
     "price_eur": 1.57,   "status": "FOUND"},

    # 4 – Microchip MCP1700T-3302E/TT, qty 100
    # DigiKey: 22763 в наличии, €0.3768@25+
    {"pn": "MCP1700T-3302E/TT",      "qty": 100, "manufacturer": "Microchip",
     "distributor": "DigiKey", "moq": 1, "stock": 22763, "lead": None,
     "price_eur": 0.3768, "status": "FOUND"},

    # 5 – STMicro LM258ADT, qty 20
    # Mouser: 53465 в наличии, €0.190@10+
    {"pn": "LM258ADT",               "qty": 20,  "manufacturer": "STMicroelectronics",
     "distributor": "Mouser", "moq": 1, "stock": 53465, "lead": None,
     "price_eur": 0.190,  "status": "FOUND"},

    # 6 – CML CMX994GQ4, qty 30
    # Mouser: 55 в наличии, €19.31@25+
    {"pn": "CMX994GQ4",              "qty": 30,  "manufacturer": "CML Microcircuits",
     "distributor": "Mouser", "moq": 1, "stock": 55,    "lead": None,
     "price_eur": 19.31,  "status": "FOUND"},

    # 7 – CML CMX7341Q3, qty 30
    # Mouser: 0 в наличии, 27 недель, €13.48@26+
    {"pn": "CMX7341Q3",              "qty": 30,  "manufacturer": "CML Microcircuits",
     "distributor": "Mouser", "moq": 1, "stock": 0,     "lead": 27,
     "price_eur": 13.48,  "status": "FOUND"},

    # 8 – Winbond W9825G6JB-6I, qty 30
    # Только Mouser Price on Request — нет авторизованной цены
    {"pn": "W9825G6JB-6I",           "qty": 30,  "manufacturer": "Winbond Electronics",
     "status": "RFQ"},

    # 9 – NXP LPC4337JET256, qty 30
    # Mouser: 260 в наличии, €15.51@10+
    {"pn": "LPC4337JET256",          "qty": 30,  "manufacturer": "NXP Semiconductors",
     "distributor": "Mouser", "moq": 1, "stock": 260,   "lead": None,
     "price_eur": 15.51,  "status": "FOUND"},

    # 10 – TI TLV320AIC3110IRHBR, qty 30
    # Arrow: 2926 в наличии, €1.9797@10+
    {"pn": "TLV320AIC3110IRHBR",     "qty": 30,  "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 2926,  "lead": None,
     "price_eur": 1.9797, "status": "FOUND"},

    # 11 – TI LMX2571NJKx, qty 20 (x = упаковка, найден суффикс T = лента)
    # Mouser: 1909 в наличии, €11.54@10+
    {"pn": "LMX2571NJKx",            "qty": 20,  "manufacturer": "Texas Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 1909,  "lead": None,
     "bpn": "LMX2571NJKT",
     "price_eur": 11.54,  "status": "FOUND"},

    # 12 – TI SN74LVC1G3157DCKR, qty 70
    # Mouser: 405410 в наличии, €0.120@10+
    {"pn": "SN74LVC1G3157DCKR",      "qty": 70,  "manufacturer": "Texas Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 405410,"lead": None,
     "price_eur": 0.120,  "status": "FOUND"},

    # 13 – TI LM45BIM3/N0PB, qty 30
    # Arrow: 638 в наличии, €1.514@5+ (поиск без слэша: LM45BIM3NOPB)
    {"pn": "LM45BIM3/N0PB",          "qty": 30,  "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 638,   "lead": None,
     "bpn": "LM45BIM3NOPB",
     "price_eur": 1.514,  "status": "FOUND"},

    # 14 – Mini-Circuits ADE-1L+, qty 30
    # DigiKey: 917 в наличии, €6.898@20+
    {"pn": "ADE-1L+",                "qty": 30,  "manufacturer": "Mini-Circuits",
     "distributor": "DigiKey", "moq": 1, "stock": 917,   "lead": None,
     "price_eur": 6.898,  "status": "FOUND"},

    # 15 – TI LMV431A1MF/N0PB, qty 20
    # SC-70-5 — нет у авторизованных дистрибьюторов
    {"pn": "LMV431A1MF/N0PB",        "qty": 20,  "manufacturer": "Texas Instruments",
     "status": "RFQ"},

    # 16 – TI TRF37D73IDSGx, qty 30 (x → R = рулон)
    # Mouser: 4410 в наличии, €1.08@10+
    {"pn": "TRF37D73IDSGx",          "qty": 30,  "manufacturer": "Texas Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 4410,  "lead": None,
     "bpn": "TRF37D73IDSGR",
     "price_eur": 1.08,   "status": "FOUND"},

    # 17 – ON Semi MC33063AP1, qty 20
    # Arrow: 180 в наличии, €0.6016@10+
    {"pn": "MC33063AP1",             "qty": 20,  "manufacturer": "ON Semiconductor",
     "distributor": "Arrow", "moq": 1, "stock": 180,   "lead": None,
     "price_eur": 0.6016, "status": "FOUND"},

    # 18 – Microchip PL133-37TI-R, qty 30
    # Mouser: 9680 в наличии, €0.602@1+
    {"pn": "PL133-37TI-R",           "qty": 30,  "manufacturer": "Microchip",
     "distributor": "Mouser", "moq": 1, "stock": 9680,  "lead": None,
     "price_eur": 0.602,  "status": "FOUND"},

    # 19 – TI TPS563200DDCR, qty 30
    # Mouser: 279 в наличии, €0.678@10+
    {"pn": "TPS563200DDCR",          "qty": 30,  "manufacturer": "Texas Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 279,   "lead": None,
     "price_eur": 0.678,  "status": "FOUND"},

    # 20 – TOP POWER TP5100-QFN16, qty 20
    # Китайский производитель, нет у авторизованных дистрибьюторов
    {"pn": "TP5100-QFN16",           "qty": 20,  "manufacturer": "TOP POWER",
     "status": "RFQ"},

    # 21 – TI TLV9001IDBVR, qty 110
    # DigiKey: 44957 в наличии, €0.278@10+
    {"pn": "TLV9001IDBVR",           "qty": 110, "manufacturer": "Texas Instruments",
     "distributor": "DigiKey", "moq": 1, "stock": 44957, "lead": None,
     "price_eur": 0.278,  "status": "FOUND"},

    # 22 – TI LMX2571NJKT, qty 20 (та же микросхема что и поз. 11)
    # Mouser: 1909 в наличии, €11.54@10+
    {"pn": "LMX2571NJKT",            "qty": 20,  "manufacturer": "Texas Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 1909,  "lead": None,
     "price_eur": 11.54,  "status": "FOUND"},

    # 23 – Seiko S-8252AAC-M6T1U, qty 30
    # Mouser: 237 в наличии, €1.19@10+
    {"pn": "S-8252AAC-M6T1U",        "qty": 30,  "manufacturer": "Seiko Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 237,   "lead": None,
     "price_eur": 1.19,   "status": "FOUND"},

    # 24 – ON Semi EMI7112FCTAG, qty 50
    # Mouser: 13532 в наличии, €0.272@10+
    {"pn": "EMI7112FCTAG",           "qty": 50,  "manufacturer": "ON Semiconductor",
     "distributor": "Mouser", "moq": 1, "stock": 13532, "lead": None,
     "price_eur": 0.272,  "status": "FOUND"},

    # 25 – DVSI AMBE-3000F-BGA, qty 30
    # Нишевой вокодер — нет у авторизованных дистрибьюторов
    {"pn": "AMBE-3000F-BGA",         "qty": 30,  "manufacturer": "DVSI",
     "status": "RFQ"},
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
    ("Part Number",              30, F_HDR),
    ("Distributor Part Number",  26, F_HDRB),
    ("Qty for Single BOM",       11, F_HDR),
    ("Manufacturer",             22, F_HDR),
    ("Distributor",              18, F_HDR),
    ("Minimum Order",            11, F_HDR),
    ("Stock",                    10, F_HDR),
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
        f"Цена по ценовому брекету ≥ закупаемого кол-ва. "
        f"Авторизованные: DigiKey, Mouser, Arrow. "
        f"Поз. 2/11/16 (суффикс x): найдены реальные PN с суффиксом R или T. "
        f"Поз. 8 (W9825G6JB-6I): только Mouser Price on Request. "
        f"Поз. 15 (LMV431A1MF/N0PB): SC-70-5 — нет у авторизованных. "
        f"Поз. 20 (TP5100-QFN16): китайский TOP POWER — нет у авторизованных. "
        f"Поз. 25 (AMBE-3000F-BGA): DVSI нишевой вокодер — нет у авторизованных."
    )
    nc = ws.cell(row=note_row, column=1, value=note)
    nc.font = Font(name="Arial", size=9, color="595959")
    nc.alignment = Alignment(horizontal="left", wrap_text=True)
    ws.merge_cells(f"A{note_row}:J{note_row}")
    ws.row_dimensions[note_row].height = 65

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
