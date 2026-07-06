"""
Конвертирует экспорт из bomtool.oemsecrets.com (.xlsx) в финальный формат BOM
(тот же стиль, что и в update_bom*.py).

Запуск:
    python bomtool_to_final.py "C:\\путь\\к\\экспорту.xlsx" "final\\Имя_итогового_файла.xlsx"

Если второй аргумент не указан — сохраняет как final/<имя_исходного>.xlsx
"""
import sys
import urllib.request, json, re
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent

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

def read_bomtool_export(path):
    """Читает экспорт bomtool.oemsecrets.com, возвращает список позиций в общем формате."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    header = [str(c.value).strip() if c.value else "" for c in ws[1]]
    idx = {name: i for i, name in enumerate(header)}

    def col(name):
        return idx.get(name)

    results = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        pn = row[col("Part Number")] if col("Part Number") is not None else None
        if not pn:
            continue
        qty = row[col("Quantity for Single BOM")] if col("Quantity for Single BOM") is not None else None
        mfr = row[col("Manufacturer")] if col("Manufacturer") is not None else None
        dist = row[col("Distributor")] if col("Distributor") is not None else None
        moq = row[col("Minimum Order")] if col("Minimum Order") is not None else None
        stock = row[col("Stock")] if col("Stock") is not None else None
        lead = row[col("Lead Time on Additional Stock in Weeks")] if col("Lead Time on Additional Stock in Weeks") is not None else None
        price = row[col("Unit Price in USD")] if col("Unit Price in USD") is not None else None
        sku = row[col("Distributor SKU")] if col("Distributor SKU") is not None else None

        stock = int(stock) if isinstance(stock, (int, float)) else 0
        lead = int(lead) if isinstance(lead, (int, float)) and lead > 0 else None

        if dist and price is not None:
            bpn = None
            if sku and str(sku).strip().upper() != str(pn).strip().upper():
                # похоже на partномер дистрибьютора, отличный от исходного PN —
                # запишем только если это похоже на вариант упаковки (не всегда точно определимо)
                bpn = None  # оставляем пусто; SKU дистрибьютора обычно внутренний код, не PN с суффиксом
            results.append({
                "pn": str(pn).strip(), "qty": int(qty) if qty else 1,
                "manufacturer": str(mfr).strip() if mfr else None,
                "distributor": str(dist).strip(),
                "moq": int(moq) if isinstance(moq, (int, float)) else None,
                "stock": stock, "lead": lead,
                "price_usd": float(price), "bpn": bpn,
                "status": "FOUND",
            })
        else:
            results.append({
                "pn": str(pn).strip(), "qty": int(qty) if qty else 1,
                "manufacturer": str(mfr).strip() if mfr else None,
                "status": "RFQ",
            })
    return results

# ── Стили (те же, что в update_bom*.py) ───────────────────────────────────────

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
    ("Distributor",              18, F_HDR),
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

def build_bom(results, output_path, rate_rub, source_note=""):
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
            lead_val = "In stock" if in_stock else (lead if lead is not None else "")

            usd = r["price_usd"]
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
        f"Найдено: {found} | RFQ: {rfq} | Курс USD/RUB (ЦБ + наценка): {rate_rub:.4f} | "
        f"Источник: bomtool.oemsecrets.com (Data Autofill, Distributor Preferences ограничены "
        f"крупными авторизованными дистрибьюторами). {source_note}"
    )
    nc = ws.cell(row=note_row, column=1, value=note)
    nc.font = Font(name="Arial", size=9, color="595959")
    nc.alignment = Alignment(horizontal="left", wrap_text=True)
    ws.merge_cells(f"A{note_row}:J{note_row}")
    ws.row_dimensions[note_row].height = 45

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f"\nГотово: {output_path}")
    print(f"Найдено: {found} | RFQ: {rfq}")
    print(f"USD/RUB: {rate_rub:.4f}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Использование: python bomtool_to_final.py <экспорт.xlsx> [итоговый_файл.xlsx]")
        sys.exit(1)
    src = Path(sys.argv[1])
    if len(sys.argv) >= 3:
        out = Path(sys.argv[2])
    else:
        out = SCRIPT_DIR / "final" / (src.stem + "_final.xlsx")

    print("Читаю экспорт BoM Tool...")
    results = read_bomtool_export(src)
    print(f"Позиций: {len(results)}")
    print("Получаю курс USD/RUB...")
    rate_rub = get_usd_rub()
    print(f"USD/RUB: {rate_rub:.4f}")
    build_bom(results, out, rate_rub)
