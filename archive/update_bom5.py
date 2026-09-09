"""
Обновляет два файла найденными ценами:
  final/BOM_Приложение_№1.xlsx      (34 позиции, источник start/Приложение №1.xlsx)
  final/BOM_Приложение_№1_(2).xlsx  (15 позиций, источник start/Приложение №1 (2).xlsx)
Запускать: python update_bom5.py
"""
import urllib.request, json, re
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent

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

RESULTS_1 = [
    {"pn": "VI-2W1-IV", "qty": 210, "manufacturer": "Vicor",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 684.4413, "status": "FOUND"},
    {"pn": "IP178GI", "qty": 150, "manufacturer": "IC Plus", "status": "RFQ"},
    {"pn": "C8051F220", "qty": 250, "manufacturer": "Silicon Labs",
     "distributor": "TME", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 4.7945, "status": "FOUND"},
    {"pn": "DS32KHZSN", "qty": 250, "manufacturer": "Analog Devices (Maxim)",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 4.0709, "status": "FOUND"},
    {"pn": "LT8616IFE#PBF", "qty": 200, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 99, "lead": None,
     "price_eur": 6.6015, "status": "FOUND"},
    {"pn": "LP2951ACM/NOPB", "qty": 2000, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 1995, "lead": None,
     "price_eur": 0.6833, "status": "FOUND"},
    {"pn": "MC14052BDG", "qty": 250, "manufacturer": "ON Semiconductor",
     "distributor": "Avnet America", "moq": 1, "stock": 8760, "lead": None,
     "price_eur": 0.1963, "status": "FOUND"},
    {"pn": "MBC13916NT1", "qty": 250, "manufacturer": "ON Semiconductor", "status": "RFQ"},
    {"pn": "MC14053BDR2G", "qty": 250, "manufacturer": "ON Semiconductor",
     "distributor": "TME", "moq": 1, "stock": 1691, "lead": None,
     "price_eur": 0.2507, "status": "FOUND"},
    {"pn": "SN74LVC1G139DCUx", "qty": 400, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 2274, "lead": None,
     "bpn": "SN74LVC1G139DCUR", "price_eur": 0.2891, "status": "FOUND"},
    {"pn": "ADM232LARZ", "qty": 2600, "manufacturer": "Analog Devices",
     "distributor": "TME", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 2.0422, "status": "FOUND"},
    {"pn": "TOIM4232-TR3", "qty": 250, "manufacturer": "Vishay", "status": "RFQ"},
    {"pn": "LT86161EF#PBF", "qty": 10, "manufacturer": "Analog Devices", "status": "RFQ"},
    {"pn": "SA636DK", "qty": 200, "manufacturer": "NXP Semiconductors",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 2.6003, "status": "FOUND"},
    {"pn": "DS7505U+", "qty": 500, "manufacturer": "Analog Devices (Maxim)",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 2.1332, "status": "FOUND"},
    {"pn": "XCF01SV020C", "qty": 10, "manufacturer": "AMD (Xilinx)", "status": "RFQ"},
    {"pn": "SPF5043Z", "qty": 200, "manufacturer": "MACOM",
     "distributor": "DigiKey", "moq": 1, "stock": 38065, "lead": None,
     "price_eur": 3.3396, "status": "FOUND"},
    {"pn": "LMX2306TM", "qty": 200, "manufacturer": "Texas Instruments", "status": "RFQ"},
    {"pn": "LP2985IM5-2.5", "qty": 200, "manufacturer": "Texas Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 7721, "lead": None,
     "price_eur": 0.4090, "status": "FOUND"},
    {"pn": "LM6211MF", "qty": 405, "manufacturer": "Texas Instruments",
     "distributor": "DigiKey", "moq": 1, "stock": 794, "lead": None,
     "price_eur": 2.5270, "status": "FOUND"},
    {"pn": "TPS76333DBVT", "qty": 205, "manufacturer": "Texas Instruments",
     "distributor": "TME", "moq": 1, "stock": 1621, "lead": None,
     "price_eur": 0.5724, "status": "FOUND"},
    {"pn": "SKY72310-11", "qty": 400, "manufacturer": "Skyworks Solutions",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 92.44, "status": "FOUND"},
    {"pn": "MC78L05ABDR2G", "qty": 400, "manufacturer": "ON Semiconductor",
     "distributor": "Mouser", "moq": 1, "stock": 24400, "lead": None,
     "price_eur": 0.1580, "status": "FOUND"},
    {"pn": "LM2903DG", "qty": 1120, "manufacturer": "Texas Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 20442, "lead": None,
     "price_eur": 0.1770, "status": "FOUND"},
    {"pn": "SN74LVC8T245PW", "qty": 400, "manufacturer": "Texas Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 24899, "lead": None,
     "price_eur": 1.0300, "status": "FOUND"},
    {"pn": "TDA8945S", "qty": 250, "manufacturer": "NXP Semiconductors", "status": "RFQ"},
    {"pn": "LP2985IM5-3.3", "qty": 210, "manufacturer": "Texas Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 19339, "lead": None,
     "price_eur": 0.4700, "status": "FOUND"},
    {"pn": "TC620CVOA", "qty": 210, "manufacturer": "Microchip",
     "distributor": "TME", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 1.4375, "status": "FOUND"},
    {"pn": "DS1337U", "qty": 250, "manufacturer": "Analog Devices (Maxim)",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 1.7635, "status": "FOUND"},
    {"pn": "TL431AILP", "qty": 650, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 881, "lead": None,
     "price_eur": 0.3407, "status": "FOUND"},
    {"pn": "TPS76350DBVT", "qty": 405, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 200, "lead": None,
     "price_eur": 0.6769, "status": "FOUND"},
    {"pn": "AD8044ARZ-14", "qty": 150, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 4.9969, "status": "FOUND"},
    {"pn": "OP462GS", "qty": 100, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 5.5377, "status": "FOUND"},
    {"pn": "OP492GS", "qty": 120, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 5.5447, "status": "FOUND"},
]

RESULTS_2 = [
    {"pn": "LC03-6.TBT", "qty": 1000, "manufacturer": "Semtech",
     "distributor": "TME", "moq": 1, "stock": 2579, "lead": None,
     "price_eur": 0.7012, "status": "FOUND"},
    {"pn": "MT9172AN1", "qty": 1500, "manufacturer": "Microchip (Microsemi)",
     "distributor": "Newark", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 16.6447, "status": "FOUND"},
    {"pn": "MDC3105LT1G", "qty": 1300, "manufacturer": "ON Semiconductor",
     "distributor": "Farnell", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.3760, "status": "FOUND"},
    {"pn": "88E1112-XX-NNC-I000", "qty": 195, "manufacturer": "Marvell", "status": "RFQ"},
    {"pn": "TEN40-2411", "qty": 200, "manufacturer": "Traco Power",
     "distributor": "Arrow", "moq": 1, "stock": 42, "lead": None,
     "price_eur": 73.4770, "status": "FOUND"},
    {"pn": "XCV400-4HQ240I", "qty": 380, "manufacturer": "AMD (Xilinx)", "status": "RFQ"},
    {"pn": "TPS7A7300RGWx", "qty": 650, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 36, "lead": None,
     "bpn": "TPS7A7300RGWR", "price_eur": 2.1781, "status": "FOUND"},
    {"pn": "KSD-01F-45-H", "qty": 100, "manufacturer": "не определён", "status": "RFQ"},
    {"pn": "SSM2135S", "qty": 1000, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 7, "lead": None,
     "price_eur": 5.0469, "status": "FOUND"},
    {"pn": "XC2V250-4FG2561", "qty": 205, "manufacturer": "AMD (Xilinx)", "status": "RFQ"},
    {"pn": "DS7505U+", "qty": 50, "manufacturer": "Analog Devices (Maxim)",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 2.1332, "status": "FOUND"},
    {"pn": "551SDCGI", "qty": 300, "manufacturer": "API Technologies",
     "distributor": "DigiKey", "moq": 1, "stock": 820, "lead": None,
     "price_eur": 1.3990, "status": "FOUND"},
    {"pn": "MT48LC4M16A2TG-75IT:G", "qty": 200, "manufacturer": "Micron", "status": "RFQ"},
    {"pn": "QS3390QG8", "qty": 150, "manufacturer": "Renesas (IDT)",
     "distributor": "Farnell", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 1.5100, "status": "FOUND"},
    {"pn": "TPS73HD301PWPR", "qty": 500, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 59, "lead": None,
     "price_eur": 7.5924, "status": "FOUND"},
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
    ("Manufacturer",             24, F_HDR),
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

def build_bom(results, output_path, rate_eur, rate_rub, extra_note=""):
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
    for i, r in enumerate(results):
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
                lead_val = f"{lead} weeks" if lead != 1 else "1 week"
            else:
                lead_val = "RFQ" if stock == 0 else ""

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

    found = sum(1 for r in results if r["status"] == "FOUND")
    rfq   = len(results) - found
    note_row = len(results) + 3
    note = (
        f"Найдено: {found} | RFQ: {rfq} | Курс EUR/USD: {rate_eur:.4f} | "
        f"Курс USD/RUB (ЦБ + наценка): {rate_rub:.4f} | "
        f"Цена по ценовому брекету ≥ закупаемого кол-ва, минимум среди авторизованных. "
        f"Авторизованные: DigiKey, Mouser, Arrow, TTI, TME, Avnet, Newark/Farnell, Future. "
        + extra_note
    )
    nc = ws.cell(row=note_row, column=1, value=note)
    nc.font = Font(name="Arial", size=9, color="595959")
    nc.alignment = Alignment(horizontal="left", wrap_text=True)
    ws.merge_cells(f"A{note_row}:J{note_row}")
    ws.row_dimensions[note_row].height = 65

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f"\nГотово: {output_path}")
    print(f"Найдено: {found} | RFQ: {rfq}")

if __name__ == "__main__":
    print("Получаю актуальные курсы валют...")
    rate_eur = get_eur_usd()
    rate_rub = get_usd_rub()
    print(f"EUR/USD: {rate_eur:.4f}  |  USD/RUB: {rate_rub:.4f}")

    build_bom(
        RESULTS_1, SCRIPT_DIR / "final" / "BOM_Приложение_№1.xlsx",
        rate_eur, rate_rub,
        extra_note=(
            "Позиция 10 (SN74LVC1G139DCUx): суффикс 'x' — найден реальный вариант "
            "SN74LVC1G139DCUR. Позиция 13 (LT86161EF#PBF): не найден ни на одной "
            "площадке, возможна опечатка в PN (нет такой микросхемы Analog Devices)."
        ),
    )
    build_bom(
        RESULTS_2, SCRIPT_DIR / "final" / "BOM_Приложение_№1_(2).xlsx",
        rate_eur, rate_rub,
        extra_note=(
            "Позиция 7 (TPS7A7300RGWx): суффикс 'x' — найден реальный вариант "
            "TPS7A7300RGWR. Позиция 11 (DS7505U+) — тот же PN, что и в первом "
            "приложении, повторный поиск с той же ценой."
        ),
    )
