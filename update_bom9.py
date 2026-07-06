"""
Обновляет final/BOM_Приложение1_№1.xlsx найденными ценами (99 позиций ИС).
Источник: start/Приложение1 №1.xlsx
Запускать: python update_bom9.py
"""
import urllib.request, json, re
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT = SCRIPT_DIR / "final" / "BOM_Приложение1_№1.xlsx"

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

# ── Данные (собраны с oemsecrets через Chrome, 2026-07-02..03) ────────────────

RESULTS = [
    {"pn": "74AHC02PW", "qty": 550, "manufacturer": "Nexperia",
     "distributor": "Arrow", "moq": 1, "stock": 40, "lead": None,
     "price_eur": 0.0933, "status": "FOUND"},
    {"pn": "88E1112-XX-NNC-I000", "qty": 70, "manufacturer": "Marvell", "status": "RFQ"},
    {"pn": "MT9172AN1", "qty": 1500, "manufacturer": "Mitel (Microchip)",
     "distributor": "Newark", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 16.6447, "status": "FOUND"},
    {"pn": "MDC3105LT1G", "qty": 1200, "manufacturer": "ON Semiconductor",
     "distributor": "Farnell", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.376, "status": "FOUND"},
    {"pn": "TMS320C6455BGTZA", "qty": 100, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 384.7435, "status": "FOUND"},
    {"pn": "TEN40-2411", "qty": 200, "manufacturer": "Traco Power",
     "distributor": "Arrow", "moq": 1, "stock": 42, "lead": None,
     "price_eur": 73.477, "status": "FOUND"},
    {"pn": "W681512RG", "qty": 200, "manufacturer": "Nuvoton",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 1.8897, "status": "FOUND"},
    {"pn": "OP467GS", "qty": 1400, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 18.153, "status": "FOUND"},
    {"pn": "XC6SLX100T-3FGG676I", "qty": 100, "manufacturer": "AMD (Xilinx)",
     "distributor": "Farnell", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 338.58, "status": "FOUND"},
    {"pn": "ADM3222ARWZ", "qty": 1200, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 2.6456, "status": "FOUND"},
    {"pn": "MAX3160EAP", "qty": 400, "manufacturer": "Analog Devices (Maxim)",
     "distributor": "TME", "moq": 1, "stock": 102, "lead": None,
     "price_eur": 19.665, "status": "FOUND"},
    {"pn": "THS1403IPFB", "qty": 300, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 28.7196, "status": "FOUND"},
    {"pn": "XCV400-4HQ240I", "qty": 400, "manufacturer": "AMD (Xilinx)", "status": "RFQ"},
    {"pn": "SN74AHCT244PW", "qty": 600, "manufacturer": "Texas Instruments",
     "distributor": "TME", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.3042, "status": "FOUND"},
    {"pn": "S29AL008J70TFI01x", "qty": 1400, "manufacturer": "Infineon (Spansion)",
     "distributor": "Arrow", "moq": 1, "stock": 1056, "lead": None,
     "bpn": "S29AL008J70TFI010", "price_eur": 1.8019, "status": "FOUND"},
    {"pn": "AD8132ARZ", "qty": 550, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 735, "lead": None,
     "price_eur": 2.8291, "status": "FOUND"},
    {"pn": "S29GL128P90TFIR1x", "qty": 200, "manufacturer": "Infineon (Cypress)",
     "distributor": "TME", "moq": 1, "stock": 0, "lead": None,
     "bpn": "S29GL128P90TFIR10", "price_eur": 9.0634, "status": "FOUND"},
    {"pn": "AD843AQ", "qty": 2100, "manufacturer": "Analog Devices", "status": "RFQ"},
    {"pn": "LTC1564IG#PBF", "qty": 400, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 14.5205, "status": "FOUND"},
    {"pn": "SN74AHC244PW", "qty": 2100, "manufacturer": "Texas Instruments",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.2753, "status": "FOUND"},
    {"pn": "MAX3243EAI", "qty": 370, "manufacturer": "Analog Devices (Maxim)",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 8.2611, "status": "FOUND"},
    {"pn": "TPS7A7300RGWx", "qty": 650, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 36, "lead": None,
     "bpn": "TPS7A7300RGWR", "price_eur": 2.1781, "status": "FOUND"},
    {"pn": "FM28V020-SGTR", "qty": 200, "manufacturer": "Infineon (Cypress)",
     "distributor": "Arrow", "moq": 1, "stock": 734, "lead": None,
     "price_eur": 11.9432, "status": "FOUND"},
    {"pn": "ADG601BRTZ-x", "qty": 720, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "bpn": "ADG601BRTZ", "price_eur": 1.6475, "status": "FOUND"},
    {"pn": "AMBE-2000", "qty": 400, "manufacturer": "DVSI", "status": "RFQ"},
    {"pn": "PIC32MX250F128B-I/SS", "qty": 300, "manufacturer": "Microchip",
     "distributor": "Future Electronics", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 3.4348, "status": "FOUND"},
    {"pn": "KSD-01F-45-H", "qty": 80, "manufacturer": "Dart Electronics", "status": "RFQ"},
    {"pn": "MT25QL128ABA1ESE-0SIT", "qty": 260, "manufacturer": "Micron",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 15.213, "status": "FOUND"},
    {"pn": "SSM2135S", "qty": 920, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 7, "lead": None,
     "price_eur": 5.0469, "status": "FOUND"},
    {"pn": "74AHCT244PW", "qty": 1400, "manufacturer": "Nexperia",
     "distributor": "Arrow", "moq": 1, "stock": 800, "lead": None,
     "price_eur": 0.1337, "status": "FOUND"},
    {"pn": "MAX998EUT-T", "qty": 300, "manufacturer": "Analog Devices (Maxim)",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 5.6845, "status": "FOUND"},
    {"pn": "TPS84610RKGx", "qty": 280, "manufacturer": "Texas Instruments",
     "distributor": "Mouser", "moq": 1, "stock": 1091, "lead": None,
     "bpn": "TPS84610RKGR", "price_eur": 4.58, "status": "FOUND"},
    {"pn": "88E6185-XX-LKJ-1000", "qty": 100, "manufacturer": "Marvell", "status": "RFQ"},
    {"pn": "OP292GSZ", "qty": 1650, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 4.0606, "status": "FOUND"},
    {"pn": "DS7505U+", "qty": 100, "manufacturer": "Analog Devices (Maxim)",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 2.1332, "status": "FOUND"},
    {"pn": "XC2V250-4FG2561", "qty": 200, "manufacturer": "AMD (Xilinx)", "status": "RFQ"},
    {"pn": "551SDCGI", "qty": 280, "manufacturer": "Renesas",
     "distributor": "DigiKey", "moq": 1, "stock": 820, "lead": None,
     "price_eur": 1.399, "status": "FOUND"},
    {"pn": "LT1308AIS8#TR", "qty": 330, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 7.6999, "status": "FOUND"},
    {"pn": "XC18V04VQG44C", "qty": 380, "manufacturer": "AMD (Xilinx)", "status": "RFQ"},
    {"pn": "MT48LC4M16A2P-6AIT:G", "qty": 210, "manufacturer": "Micron",
     "distributor": "Future Electronics", "moq": 1, "stock": 0, "lead": None,
     "bpn": "MT48LC4M16A2P-6AIT", "price_eur": 15.6271, "status": "FOUND"},
    {"pn": "ADG1436YRUZ", "qty": 710, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 4.7449, "status": "FOUND"},
    {"pn": "TPS84410RKGx", "qty": 510, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 187, "lead": None,
     "bpn": "TPS84410RKGR", "price_eur": 4.439, "status": "FOUND"},
    {"pn": "PAM8615RHR", "qty": 210, "manufacturer": "Diodes Incorporated",
     "distributor": "TME", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 1.5819, "status": "FOUND"},
    {"pn": "AD823ARZ", "qty": 3500, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 14, "lead": None,
     "price_eur": 4.8839, "status": "FOUND"},
    {"pn": "74AHC244PW", "qty": 1750, "manufacturer": "Nexperia",
     "distributor": "Mouser", "moq": 1, "stock": 31693, "lead": None,
     "price_eur": 0.206, "status": "FOUND"},
    {"pn": "MAX6635MSA", "qty": 100, "manufacturer": "Analog Devices (Maxim)",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 3.465, "status": "FOUND"},
    {"pn": "SN74AHC1G04DBVR", "qty": 280, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 1639, "lead": None,
     "price_eur": 0.0486, "status": "FOUND"},
    {"pn": "QS3390QG8", "qty": 100, "manufacturer": "Renesas (IDT)",
     "distributor": "Farnell", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 1.51, "status": "FOUND"},
    {"pn": "TPS7A4700RGWx", "qty": 520, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "bpn": "TPS7A4700RGWR", "price_eur": 2.6423, "status": "FOUND"},
    {"pn": "ADM3222ARS", "qty": 100, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 2.6552, "status": "FOUND"},
    {"pn": "TPS73HD301PWPR", "qty": 600, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 59, "lead": None,
     "price_eur": 7.5924, "status": "FOUND"},
    {"pn": "512MILFT", "qty": 300, "manufacturer": "Renesas", "status": "RFQ"},
    {"pn": "74ACT163MTC", "qty": 560, "manufacturer": "Fairchild Semiconductor",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.1765, "status": "FOUND"},
    {"pn": "MT88L85AN", "qty": 200, "manufacturer": "Mitel", "status": "RFQ"},
    {"pn": "XCF04SV020C", "qty": 350, "manufacturer": "AMD (Xilinx)", "status": "RFQ"},
    {"pn": "74AHCT74PW", "qty": 550, "manufacturer": "Nexperia",
     "distributor": "Arrow", "moq": 1, "stock": 1611, "lead": None,
     "price_eur": 0.1506, "status": "FOUND"},
    {"pn": "SN74LVC8T245DGVR", "qty": 3000, "manufacturer": "Texas Instruments",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.6502, "status": "FOUND"},
    {"pn": "XC3S400-4PQ208I", "qty": 390, "manufacturer": "AMD (Xilinx)", "status": "RFQ"},
    {"pn": "MT47H128M16PT-25EIT:C", "qty": 200, "manufacturer": "Micron", "status": "RFQ"},
    {"pn": "PIC18LF452-I/PT", "qty": 100, "manufacturer": "Microchip",
     "distributor": "Farnell", "moq": 1, "stock": 34, "lead": None,
     "price_eur": 8.13, "status": "FOUND"},
    {"pn": "TPS73HD325PWPR", "qty": 380, "manufacturer": "Texas Instruments",
     "distributor": "DigiKey", "moq": 1, "stock": 953, "lead": None,
     "price_eur": 10.955, "status": "FOUND"},
    {"pn": "AD7945BR", "qty": 270, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 9.9816, "status": "FOUND"},
    {"pn": "LT1121IST-5", "qty": 550, "manufacturer": "Analog Devices (Linear Tech)",
     "distributor": "TME", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 3.4086, "status": "FOUND"},
    {"pn": "TLV5639IDW", "qty": 510, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 10.3814, "status": "FOUND"},
    {"pn": "AT45DB321D-SU", "qty": 200, "manufacturer": "Microchip (Adesto)",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 2.6503, "status": "FOUND"},
    {"pn": "LP3878MR-ADJ/NOPB", "qty": 970, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 1.9953, "status": "FOUND"},
    {"pn": "XC6SLX150T-3FGG484I", "qty": 270, "manufacturer": "AMD (Xilinx)",
     "distributor": "Farnell", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 381.72, "status": "FOUND"},
    {"pn": "LC03-6.TBT", "qty": 1000, "manufacturer": "Semtech",
     "distributor": "TME", "moq": 1, "stock": 2579, "lead": None,
     "price_eur": 0.7012, "status": "FOUND"},
    {"pn": "LM2678S-12", "qty": 250, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 4.4434, "status": "FOUND"},
    {"pn": "LP2951ACM/NOPB", "qty": 900, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 1995, "lead": None,
     "price_eur": 0.6833, "status": "FOUND"},
    {"pn": "LP2950CDT-3.0", "qty": 250, "manufacturer": "Texas Instruments (National Semi)",
     "distributor": "Arrow", "moq": 1, "stock": 779, "lead": None,
     "price_eur": 0.4043, "status": "FOUND"},
    {"pn": "NC7WZ16P6X", "qty": 2000, "manufacturer": "ON Semiconductor (Fairchild)",
     "distributor": "Arrow", "moq": 1, "stock": 2900, "lead": None,
     "price_eur": 0.0302, "status": "FOUND"},
    {"pn": "LM2664M6", "qty": 250, "manufacturer": "Texas Instruments (National Semi)",
     "distributor": "Arrow", "moq": 1, "stock": 1353, "lead": None,
     "price_eur": 0.6618, "status": "FOUND"},
    {"pn": "NC7WZ14P6X", "qty": 1100, "manufacturer": "ON Semiconductor (Fairchild)",
     "distributor": "Newark", "moq": 1, "stock": 18692, "lead": None,
     "price_eur": 0.01, "status": "FOUND"},
    {"pn": "MGA72543", "qty": 1100, "manufacturer": "Avago (Agilent)", "status": "RFQ"},
    {"pn": "AD8027ART", "qty": 250, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 2.5206, "status": "FOUND"},
    {"pn": "PIC16F747-I/PT", "qty": 250, "manufacturer": "Microchip",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 5.4625, "status": "FOUND"},
    {"pn": "LM2678S-5.0", "qty": 250, "manufacturer": "Texas Instruments",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 4.2826, "status": "FOUND"},
    {"pn": "IR4427S", "qty": 700, "manufacturer": "Infineon (International Rectifier)",
     "distributor": "Arrow", "moq": 1, "stock": 4981, "lead": None,
     "price_eur": 1.192, "status": "FOUND"},
    {"pn": "SKY13286-359LF", "qty": 1500, "manufacturer": "Skyworks Solutions",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 2.92, "status": "FOUND"},
    {"pn": "PIC12F675-I/SN", "qty": 250, "manufacturer": "Microchip",
     "distributor": "Future Electronics", "moq": 1, "stock": 14319, "lead": None,
     "price_eur": 1.11, "status": "FOUND"},
    {"pn": "AD8313ARM", "qty": 500, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 22.8206, "status": "FOUND"},
    {"pn": "LMV321M5", "qty": 500, "manufacturer": "Texas Instruments (National Semi)",
     "distributor": "Arrow", "moq": 1, "stock": 122, "lead": None,
     "price_eur": 0.3381, "status": "FOUND"},
    {"pn": "LP2981IM5-3.0", "qty": 1100, "manufacturer": "Texas Instruments (National Semi)",
     "distributor": "Arrow", "moq": 1, "stock": 1, "lead": None,
     "price_eur": 0.0719, "status": "FOUND"},
    {"pn": "LM6211MF", "qty": 450, "manufacturer": "Texas Instruments",
     "distributor": "DigiKey", "moq": 1, "stock": 794, "lead": None,
     "price_eur": 2.527, "status": "FOUND"},
    {"pn": "74HC132D", "qty": 250, "manufacturer": "Nexperia (Philips)",
     "distributor": "TME", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.1652, "status": "FOUND"},
    {"pn": "MASWSS0161", "qty": 250, "manufacturer": "M/A-COM", "status": "RFQ"},
    {"pn": "LP2981IM5-3.3", "qty": 250, "manufacturer": "Texas Instruments (National Semi)",
     "distributor": "Mouser", "moq": 1, "stock": 19339, "lead": None,
     "price_eur": 0.47, "status": "FOUND"},
    {"pn": "LMX2434TM", "qty": 250, "manufacturer": "Texas Instruments (National Semi)", "status": "RFQ"},
    {"pn": "LM2940S-8.0", "qty": 450, "manufacturer": "Texas Instruments (National Semi)",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 0.9765, "status": "FOUND"},
    {"pn": "HMC391LP4", "qty": 250, "manufacturer": "Analog Devices (Hittite)",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 12.8281, "status": "FOUND"},
    {"pn": "LP2981IM5-5.0", "qty": 1100, "manufacturer": "Texas Instruments (National Semi)",
     "distributor": "DigiKey", "moq": 1, "stock": 12561, "lead": None,
     "price_eur": 1.0255, "status": "FOUND"},
    {"pn": "ADCMP600BKSZ", "qty": 900, "manufacturer": "Analog Devices",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 3.1175, "status": "FOUND"},
    {"pn": "HMC468ALP3E", "qty": 250, "manufacturer": "Analog Devices (Hittite)",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 5.5832, "status": "FOUND"},
    {"pn": "LM2671M-ADJ", "qty": 450, "manufacturer": "Texas Instruments (National Semi)",
     "distributor": "DigiKey", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 3.8651, "status": "FOUND"},
    {"pn": "24LC128-I/SN", "qty": 250, "manufacturer": "Microchip",
     "distributor": "Future Electronics", "moq": 1, "stock": 15402, "lead": None,
     "price_eur": 0.3103, "status": "FOUND"},
    {"pn": "HMC546MS8G", "qty": 250, "manufacturer": "Analog Devices (Hittite)",
     "distributor": "Arrow", "moq": 1, "stock": 0, "lead": None,
     "price_eur": 5.1295, "status": "FOUND"},
    {"pn": "LP2981IM5-2.5", "qty": 250, "manufacturer": "Texas Instruments (National Semi)",
     "distributor": "Arrow", "moq": 1, "stock": 23, "lead": None,
     "price_eur": 0.1027, "status": "FOUND"},
    {"pn": "LM57BISD-5", "qty": 650, "manufacturer": "Texas Instruments (National Semi)",
     "distributor": "Arrow", "moq": 1, "stock": 123, "lead": None,
     "price_eur": 1.4028, "status": "FOUND"},
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
    ("Manufacturer",             28, F_HDR),
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
        f"Позиции с суффиксом 'x' (S29AL008J70TFI01x, S29GL128P90TFIR1x, ADG601BRTZ-x, "
        f"TPS84610RKGx, TPS84410RKGx, TPS7A7300RGWx, TPS7A4700RGWx): найден реальный "
        f"вариант упаковки — см. Distributor Part Number."
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
