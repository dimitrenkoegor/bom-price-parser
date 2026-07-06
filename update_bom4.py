"""
Обновляет final/BOM_Запрос_2053.xlsx найденными ценами (42 позиции пассивов/полупроводников).
Источник: start/Запрос_2053.xls
Запускать: python update_bom4.py
"""
import urllib.request, json, re
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT = SCRIPT_DIR / "final" / "BOM_Запрос_2053.xlsx"

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

# ── Данные (собраны с oemsecrets через Chrome, 2026-07-01) ────────────────────
# bpn: Distributor Part Number, если отличается от запрошенного PN (суффикс упаковки)
# price_eur, lead: None = In stock, int = недели (если нет данных - оставляем None и stock=0)

RESULTS = [
    # 1 – LM5176PWP_ (суффикс "_" = упаковка), qty 8802
    # Mouser: 30 в наличии, €5.32@10+
    {"pn": "LM5176PWP_", "qty": 8802, "manufacturer": "Texas Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 30, "lead": None,
     "bpn": "LM5176PWPR", "price_eur": 5.32, "status": "FOUND"},

    # 2 – Фильтр ферритовый 7427930, qty 52809
    # Mouser: 50194 в наличии, €0.567@10+
    {"pn": "7427930", "qty": 52809, "manufacturer": "Würth Elektronik",
     "distributor": "Mouser", "moq": 1, "stock": 50194, "lead": None,
     "price_eur": 0.567, "status": "FOUND"},

    # 3 – 06031A100B4T2A, qty 17604
    # Arrow: 0 в наличии, 15 нед., €0.0801@8000+
    {"pn": "06031A100B4T2A", "qty": 17604, "manufacturer": "AVX (Kyocera AVX)",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": 15,
     "price_eur": 0.0801, "status": "FOUND"},

    # 4 – 08051C105K4Z2A, qty 61608
    # TTI Europe: 146000 в наличии, €0.051@4000+
    {"pn": "08051C105K4Z2A", "qty": 61608, "manufacturer": "AVX (Kyocera AVX)",
     "distributor": "TTI Europe", "moq": 1, "stock": 146000, "lead": None,
     "price_eur": 0.051, "status": "FOUND"},

    # 5 – 12101C475K4Z2A, qty 96813
    # TTI Europe: 203000 в наличии, €0.225@2000+
    {"pn": "12101C475K4Z2A", "qty": 96813, "manufacturer": "AVX (Kyocera AVX)",
     "distributor": "TTI Europe", "moq": 1, "stock": 203000, "lead": None,
     "price_eur": 0.225, "status": "FOUND"},

    # 6 – C0603X221J1GACAUTO, qty 8802
    # Arrow: 0 в наличии, 22 нед., €0.0783@4000+
    {"pn": "C0603X221J1GACAUTO", "qty": 8802, "manufacturer": "KEMET",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": 22,
     "price_eur": 0.0783, "status": "FOUND"},

    # 7 – CGA3E1X7S1C225K080AE, qty 17604
    # Avnet America: 0 в наличии, 12 нед., €0.0429@4000+
    {"pn": "CGA3E1X7S1C225K080AE", "qty": 17604, "manufacturer": "TDK",
     "distributor": "Avnet America", "moq": 1, "stock": 0, "lead": 12,
     "price_eur": 0.0429, "status": "FOUND"},

    # 8 – CGA4C4C0G2W471J060AE, qty 26406
    # Arrow: 0 в наличии, 12 нед., €0.0574@8000+
    {"pn": "CGA4C4C0G2W471J060AE", "qty": 26406, "manufacturer": "TDK",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": 12,
     "price_eur": 0.0574, "status": "FOUND"},

    # 9 – GCJ21AR72E102K_ (суффикс "_"), qty 8802
    # Avnet America: 0 в наличии, 16 нед., €0.0228@4000+
    {"pn": "GCJ21AR72E102K_", "qty": 8802, "manufacturer": "Murata",
     "distributor": "Avnet America", "moq": 1, "stock": 0, "lead": 16,
     "bpn": "GCJ21AR72E102KXJ1D", "price_eur": 0.0228, "status": "FOUND"},

    # 10 – GCJ188R71H224K_, qty 35205
    # TME: 0 в наличии, €0.0152@10000+
    {"pn": "GCJ188R71H224K_", "qty": 35205, "manufacturer": "Murata",
     "distributor": "TME", "moq": 1, "stock": 0, "lead": None,
     "bpn": "GCJ188R71H224KA01J", "price_eur": 0.0152, "status": "FOUND"},

    # 11 – GCJ188R72A104K_, qty 61608
    # TTI Europe: 316000 в наличии, €0.0136@8000+
    {"pn": "GCJ188R72A104K_", "qty": 61608, "manufacturer": "Murata",
     "distributor": "TTI Europe", "moq": 1, "stock": 316000, "lead": None,
     "bpn": "GCJ188R72A104KA01D", "price_eur": 0.0136, "status": "FOUND"},

    # 12 – GCM188R71H393K_, qty 8802
    # Avnet America: 0 в наличии, €0.0272@8000+
    {"pn": "GCM188R71H393K_", "qty": 8802, "manufacturer": "Murata",
     "distributor": "Avnet America", "moq": 1, "stock": 0, "lead": None,
     "bpn": "GCM188R71H393KA37D", "price_eur": 0.0272, "status": "FOUND"},

    # 13 – GRM21BR72E103K_, qty 17604
    # TTI Europe: 72000 в наличии, €0.0117@6000+
    {"pn": "GRM21BR72E103K_", "qty": 17604, "manufacturer": "Murata",
     "distributor": "TTI Europe", "moq": 1, "stock": 72000, "lead": None,
     "bpn": "GRM21BR72E103KA01L", "price_eur": 0.0117, "status": "FOUND"},

    # 14 – RPF1018331M063K, qty 88011
    # Avnet America: 0 в наличии, €0.6398@3600+
    {"pn": "RPF1018331M063K", "qty": 88011, "manufacturer": "Rubycon",
     "distributor": "Avnet America", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.6398, "status": "FOUND"},

    # 15 – RPF1018471M050K, qty 70410
    # Avnet America: 0 в наличии, €0.5241@3600+
    {"pn": "RPF1018471M050K", "qty": 70410, "manufacturer": "Rubycon",
     "distributor": "Avnet America", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.5241, "status": "FOUND"},

    # 16 – RK73Z1JTTD, qty 26406
    # TTI Europe: 935000 в наличии, €0.0011@10000+
    {"pn": "RK73Z1JTTD", "qty": 26406, "manufacturer": "KOA Speer",
     "distributor": "TTI Europe", "moq": 1, "stock": 935000, "lead": None,
     "price_eur": 0.0011, "status": "FOUND"},

    # 17 – LTC4367HMS8#PBF, qty 8802
    # Arrow: 49 в наличии, €3.0621@1+
    {"pn": "LTC4367HMS8#PBF", "qty": 8802, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 49, "lead": None,
     "price_eur": 3.0621, "status": "FOUND"},

    # 18 – THJP0603AST1, qty 17604
    # TTI Asia: 10000 в наличии, €0.3287@2000+
    {"pn": "THJP0603AST1", "qty": 17604, "manufacturer": "Bourns",
     "distributor": "TTI Asia", "moq": 1, "stock": 10000, "lead": None,
     "price_eur": 0.3287, "status": "FOUND"},

    # 19 – THJP0612AST1, qty 52809
    # TTI Asia: 0 в наличии, €0.4602@3000+
    {"pn": "THJP0612AST1", "qty": 52809, "manufacturer": "Bourns",
     "distributor": "TTI Asia", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.4602, "status": "FOUND"},

    # 20 – Индуктор 7443091150, qty 8802
    # Mouser: 205 в наличии, €2.37@10+
    {"pn": "7443091150", "qty": 8802, "manufacturer": "Würth Elektronik",
     "distributor": "Mouser", "moq": 1, "stock": 205, "lead": None,
     "price_eur": 2.37, "status": "FOUND"},

    # 21 – Индуктор 7443640680, qty 8802
    # DigiKey: 457 в наличии, €7.552@5+
    {"pn": "7443640680", "qty": 8802, "manufacturer": "Würth Elektronik",
     "distributor": "DigiKey", "moq": 1, "stock": 457, "lead": None,
     "price_eur": 7.552, "status": "FOUND"},

    # 22 – Индуктор FP2-S047-R, qty 8802
    # TTI Americas: 1700 в наличии, €0.9992@3400+
    {"pn": "FP2-S047-R", "qty": 8802, "manufacturer": "Bourns",
     "distributor": "TTI Americas", "moq": 1, "stock": 1700, "lead": None,
     "price_eur": 0.9992, "status": "FOUND"},

    # 23 – CRCW06032K00FKEA, qty 8802
    # TTI Europe: 880000 в наличии, €0.0012@5000+
    {"pn": "CRCW06032K00FKEA", "qty": 8802, "manufacturer": "Vishay",
     "distributor": "TTI Europe", "moq": 1, "stock": 880000, "lead": None,
     "price_eur": 0.0012, "status": "FOUND"},

    # 24 – ERA3AEB201V, qty 17604
    # TTI Asia: 10000 в наличии, €0.0238@10000+
    {"pn": "ERA3AEB201V", "qty": 17604, "manufacturer": "Panasonic",
     "distributor": "TTI Asia", "moq": 1, "stock": 10000, "lead": None,
     "price_eur": 0.0238, "status": "FOUND"},

    # 25 – ERA3AEB1542V, qty 17604
    # TTI Americas: 0 в наличии, €0.0241@10000+
    {"pn": "ERA3AEB1542V", "qty": 17604, "manufacturer": "Panasonic",
     "distributor": "TTI Americas", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.0241, "status": "FOUND"},

    # 26 – ERA3AEB3011V, qty 17604
    # TTI Europe: 0 в наличии, €0.0248@5000+
    {"pn": "ERA3AEB3011V", "qty": 17604, "manufacturer": "Panasonic",
     "distributor": "TTI Europe", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.0248, "status": "FOUND"},

    # 27 – ERA3AEB3242V, qty 17604
    # TTI Europe: 5000 в наличии, €0.0258@10000+
    {"pn": "ERA3AEB3242V", "qty": 17604, "manufacturer": "Panasonic",
     "distributor": "TTI Europe", "moq": 1, "stock": 5000, "lead": None,
     "price_eur": 0.0258, "status": "FOUND"},

    # 28 – ERA3AEB6811V, qty 8802
    # TTI Asia: 5000 в наличии, €0.0240@5000+
    {"pn": "ERA3AEB6811V", "qty": 8802, "manufacturer": "Panasonic",
     "distributor": "TTI Asia", "moq": 1, "stock": 5000, "lead": None,
     "price_eur": 0.0240, "status": "FOUND"},

    # 29 – ERA6AEB204V, qty 44007
    # TTI Asia: 35000 в наличии, €0.0275@20000+
    {"pn": "ERA6AEB204V", "qty": 44007, "manufacturer": "Panasonic",
     "distributor": "TTI Asia", "moq": 1, "stock": 35000, "lead": None,
     "price_eur": 0.0275, "status": "FOUND"},

    # 30 – ERJ3EKF5232V, qty 8802
    # TTI Americas: 15000 в наличии, €0.0038@5000+
    {"pn": "ERJ3EKF5232V", "qty": 8802, "manufacturer": "Panasonic",
     "distributor": "TTI Americas", "moq": 1, "stock": 15000, "lead": None,
     "price_eur": 0.0038, "status": "FOUND"},

    # 31 – RC1210JR-072R2L, qty 17604
    # TTI Asia: 0 в наличии, €0.0085@5000+
    {"pn": "RC1210JR-072R2L", "qty": 17604, "manufacturer": "Yageo",
     "distributor": "TTI Asia", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.0085, "status": "FOUND"},

    # 32 – BLM18KG300TN1D, qty 44007
    # Avnet America: 0 в наличии, €0.0163@4000+
    {"pn": "BLM18KG300TN1D", "qty": 44007, "manufacturer": "Murata",
     "distributor": "Avnet America", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.0163, "status": "FOUND"},

    # 33 – RN73R1JTTD10R0D25, qty 26406
    # Arrow: 0 в наличии, €0.0281@10000+
    {"pn": "RN73R1JTTD10R0D25", "qty": 26406, "manufacturer": "KOA Speer",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.0281, "status": "FOUND"},

    # 34 – WSLP25125L000FEA, qty 17604
    # TTI Europe: 0 в наличии, €0.2720@4000+
    {"pn": "WSLP25125L000FEA", "qty": 17604, "manufacturer": "Vishay",
     "distributor": "TTI Europe", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.2720, "status": "FOUND"},

    # 35 – WSLP25126L000FEA, qty 17604
    # TTI Europe: 0 в наличии, €0.3140@10000+
    {"pn": "WSLP25126L000FEA", "qty": 17604, "manufacturer": "Vishay",
     "distributor": "TTI Europe", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.3140, "status": "FOUND"},

    # 36 – BAS316,115, qty 26406
    # Arrow: 0 в наличии, €0.0133@21000+
    {"pn": "BAS316,115", "qty": 26406, "manufacturer": "Nexperia",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.0133, "status": "FOUND"},

    # 37 – LG L29K-F2J1-24, qty 8802
    # Нет ни у одного авторизованного дистрибьютора
    {"pn": "LG L29K-F2J1-24", "qty": 8802, "manufacturer": "OSRAM",
     "status": "RFQ"},

    # 38 – PMEG100V080ELPE-QZ, qty 17604
    # Arrow: 0 в наличии, €0.1316@10000+
    {"pn": "PMEG100V080ELPE-QZ", "qty": 17604, "manufacturer": "Nexperia",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.1316, "status": "FOUND"},

    # 39 – Диод ограничительный SM6T68CA, qty 8802
    # Arrow: 367500 в наличии, €0.0959@5000+
    {"pn": "SM6T68CA", "qty": 8802, "manufacturer": "Littelfuse / Vishay",
     "distributor": "Arrow", "moq": 1, "stock": 367500, "lead": None,
     "price_eur": 0.0959, "status": "FOUND"},

    # 40 – IAUT300N10S5N015ATMA1, qty 17604
    # Avnet America: 0 в наличии, €2.1352@4000+
    {"pn": "IAUT300N10S5N015ATMA1", "qty": 17604, "manufacturer": "Infineon",
     "distributor": "Avnet America", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 2.1352, "status": "FOUND"},

    # 41 – NTMFS3D5N08XT1G, qty 44007
    # Avnet America: 0 в наличии, €0.4996@3000+
    {"pn": "NTMFS3D5N08XT1G", "qty": 44007, "manufacturer": "ON Semiconductor",
     "distributor": "Avnet America", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.4996, "status": "FOUND"},

    # 42 – Клемма AMT0440007DB0000G, qty 70410
    # TME: 9575 в наличии, €0.1797@5+
    {"pn": "AMT0440007DB0000G", "qty": 70410, "manufacturer": "AVX (Kyocera AVX)",
     "distributor": "TME", "moq": 1, "stock": 9575, "lead": None,
     "price_eur": 0.1797, "status": "FOUND"},
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
        f"Позиции с суффиксом '_' в PN (1, 9, 10, 11, 12, 13): найден реальный вариант упаковки "
        f"у дистрибьютора (см. Distributor Part Number). "
        f"Позиция 37 (LG L29K-F2J1-24, OSRAM): нет у авторизованных дистрибьюторов — RFQ. "
        f"Производитель для части позиций определён по общепринятым префиксам PN "
        f"(в исходном запросе колонка «Производитель» отсутствовала)."
    )
    nc = ws.cell(row=note_row, column=1, value=note)
    nc.font = Font(name="Arial", size=9, color="595959")
    nc.alignment = Alignment(horizontal="left", wrap_text=True)
    ws.merge_cells(f"A{note_row}:J{note_row}")
    ws.row_dimensions[note_row].height = 75

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
