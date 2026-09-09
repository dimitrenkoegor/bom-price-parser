# -*- coding: utf-8 -*-
"""Golden-тесты разбора: 82 реальные строки одного тендерного BOM (лист «360»).

Запуск:  python run.py parse --selftest
   или:  PYTHONPATH=src python -m unittest tests.test_mpn -v

Если источник изменил написание — правь config/brands.toml, а не этот файл.
Тест защищает от регрессий парсера, а не от изменений в данных.
"""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

try:                                    # файл общий для двух проектов
    from countpars import mpn as m      # noqa: E402  count pars
except ImportError:
    from bomprice import mpn as m       # noqa: E402  bom-price-parser

# (исходная строка, ожидаемый артикул, ожидаемый производитель)
GOLDEN = [
    ("TPA6211A1DGN Texas Instruments", "TPA6211A1DGN", "Texas Instruments"),
    ("GRM21BR71H105KA12 (0805-X7R-50 В-1,0 мкФ±10%) Murata", "GRM21BR71H105KA12", "Murata"),
    ("DF40B-30DS-0.4V Hirose", "DF40B-30DS-0.4V", "Hirose"),
    ("LT3650IDD-4.2#PBF Linear Technology", "LT3650IDD-4.2#PBF", "Linear Technology"),
    ("XC6SLX75-3CSG484I Xilinx", "XC6SLX75-3CSG484I", "Xilinx"),
    ("QPA4263A Qorvo", "QPA4263A", "Qorvo"),
    ("1008CS-472XG (1008-4700 нГн±2%) Coilcraft", "1008CS-472XG", "Coilcraft"),
    ("BAR90-02EL Infineon Technologies", "BAR90-02EL", "Infineon Technologies"),
    ("0603AF-222XJ (0603-2200 нГн±5%) Coilcraft", "0603AF-222XJ", "Coilcraft"),
    ("0805CS-471XG Coilcraft", "0805CS-471XG", "Coilcraft"),
    ("0451005.MRL Littelfuse", "0451005.MRL", "Littelfuse"),
    ("LB550LVF Wayon", "LB550LVF", "Wayon"),
    ("AC-DC MPM-30-24 Mean Well", "MPM-30-24", "Mean Well"),
    ("TAJB107M010RNJ (Тип B - 10 B-100 мкФ±20%) AVX", "TAJB107M010RNJ", "AVX"),
    ("132-13SMG (206 нГн±2%) Coilcraft", "132-13SMG", "Coilcraft"),
    ("A04TG (12,5 нГн±2%) Coilcraft", "A04TG", "Coilcraft"),
    ("BBY57-02V Infineon Technologies", "BBY57-02V", "Infineon Technologies"),
    ("WSLP1206R0300FEA (1206-0,03 Ом±1%) Vishay", "WSLP1206R0300FEA", "Vishay"),
    ("Микросхема TPS62065DSG Texas Instruments", "TPS62065DSG", "Texas Instruments"),
    ("Конденсатор GRM1555C1H332GE01 (0402-C0G-50 В-3300 пФ±2%) Murata",
     "GRM1555C1H332GE01", "Murata"),
    ("Переключатель KSU213WLFG C&K", "KSU213WLFG", "C&K"),
    ("Микросхема TPD8F003DQDR Texas Instruments", "TPD8F003DQDR", "Texas Instruments"),
    ("Конденсатор GRM155R61A105KE15 (0402-X5R-10 B-1,0 мкФ±10%) Murata",
     "GRM155R61A105KE15", "Murata"),
    ("Транзистор BC847CT Diodes", "BC847CT", "Diodes Incorporated"),
    ("Микросхема INA190A1IRSWT Texas Instruments", "INA190A1IRSWT", "Texas Instruments"),
    ("Микросхема TCA6418EYFPR Texas Instruments", "TCA6418EYFPR", "Texas Instruments"),
    ("Микросхема TPS61099YFFT Texas Instruments", "TPS61099YFFT", "Texas Instruments"),
    ("Конденсатор GCM188R71H104KA57 (0603-X7R-50 В-0,1 мкФ±10%) Murata",
     "GCM188R71H104KA57", "Murata"),
    ("Конденсатор GRM188R61E106MA3 (0603-X5R-25 В-10 мкФ±20%) Murata",
     "GRM188R61E106MA3", "Murata"),
    ("Конденсатор GRM21BR6YA106ME43 (0805-X5R-35 В-10 мкФ±20%) Murata",
     "GRM21BR6YA106ME43", "Murata"),
    ("Конденсатор TAJA107M004RNJ (Тип A - 4,0 В-100 мкФ±20%) AVX", "TAJA107M004RNJ", "AVX"),
    ("Микросхема SKY13416-485LF Skyworks", "SKY13416-485LF", "Skyworks"),
    ("Микросхема SKY13418-485LF Skyworks", "SKY13418-485LF", "Skyworks"),
    ("Микросхема SN74LVC1G11YZP Texas Instruments", "SN74LVC1G11YZP", "Texas Instruments"),
    ("Микросхема TPS55340RTE Texas Instruments", "TPS55340RTE", "Texas Instruments"),
    ("Резистор YC124-JR-0722RL (R4MATR-22 Ом±5%) Yageo", "YC124-JR-0722RL", "Yageo"),
    ("Трансформатор 744230900 Wurth Electronics", "744230900", "Wurth Electronics"),
    ("Катушка индуктивности 0603CS-12NXGL (0603-12 нГн±2%) Coilcraft",
     "0603CS-12NXGL", "Coilcraft"),
    ("Катушка индуктивности IHLP1212ABEVR22M01 (1212-220 нГн±20%) Vishay",
     "IHLP1212ABEVR22M01", "Vishay"),
    ("Диод APHF1608LSEEQBDZGKC Kingbright", "APHF1608LSEEQBDZGKC", "Kingbright"),
    ("Катушка индуктивности IHLP2020BZER2R2M01 (2020-2,2 мкГн±20%) Vishay",
     "IHLP2020BZER2R2M01", "Vishay"),
    ("Конденсатор GRM033R60J474KE90 (0201-X5R-6,3 В-0,47 мкФ±10%) Murata",
     "GRM033R60J474KE90", "Murata"),
    ("Микросхема LMX2571NJKT Texas Instruments", "LMX2571NJKT", "Texas Instruments"),
    ("Конденсатор GRM21BR61C226ME44 (0805-X5R-16 В-22 мкФ±20%) Murata",
     "GRM21BR61C226ME44", "Murata"),
    ("Резистор WSLP2512R0100FEA (2512-10 мОм±1%) Vishay", "WSLP2512R0100FEA", "Vishay"),
    ("Предохранитель MF-MSMF150/24X Bourns", "MF-MSMF150/24X", "Bourns"),
    ("Розетка USB3.1 TYPE-C 24PF-036 Ruichi", "24PF-036", "Ruichi"),
    ("Диод LTW-C194TS5 LITE-ON", "LTW-C194TS5", "LITE-ON"),
    ("Резистор NTCS0402E3103JLT (0402-10 кОм±5%) Vishay", "NTCS0402E3103JLT", "Vishay"),
    ("0603CS-39NXGRW (0603-39 нГн±2%) Coilcraft", "0603CS-39NXGRW", "Coilcraft"),
    ("A05TG (18,5 нГн±2%) Coilcraft", "A05TG", "Coilcraft"),
    ("ADC-26-52+ Mini-Circuits", "ADC-26-52+", "Mini-Circuits"),
    ("AS4C4M32S-6BIN Alliance Memory", "AS4C4M32S-6BIN", "Alliance Memory"),
    ("BAR90-02ELE6327 Infineon Technologies", "BAR90-02ELE6327", "Infineon Technologies"),
    ("FIN1001M5X Fairchild", "FIN1001M5X", "Fairchild"),
    ("FX11LA-80P/8-SV(71) Hirose", "FX11LA-80P/8-SV(71)", "Hirose"),
    ("HMC433E Hittite", "HMC433E", "Hittite Microwave"),
    ("KSU213W C&K", "KSU213W", "C&K"),
    ("KT2520F27456ZAW28TBA AVX ", "KT2520F27456ZAW28TBA", "AVX"),
    ("LT3650IDD-4.2#PBF Linear Technology /AD", "LT3650IDD-4.2#PBF", "Linear Technology"),
    ("LTC7103IUHE#PBF Analog Devices", "LTC7103IUHE#PBF", "Analog Devices"),
    ("MLC1565-153ML (15,4 мкГн±20%) Coilcraft", "MLC1565-153ML", "Coilcraft"),
    ("MX29GL512FLXGI-10Q Macronix", "MX29GL512FLXGI-10Q", "Macronix"),
    ("SI1013CX-T1-GE3 Vishay", "SI1013CX-T1-GE3", "Vishay"),
    ("SIS436DN Vishay", "SIS436DN", "Vishay"),
    ("SUM90P10-19L Vishay", "SUM90P10-19L", "Vishay"),
    ("М83513/01-AN Amphenol", "M83513/01-AN", "Amphenol"),
    ("ADF4351BCPZ Analog Devices", "ADF4351BCPZ", "Analog Devices"),
    ("ADL5561ACPZ Analog Devices", "ADL5561ACPZ", "Analog Devices"),
    ("Чип-конденсатор MURATA ZRB18AR6YA475KE05L", "ZRB18AR6YA475KE05L", "Murata"),
    ("Усилитель операционный ANALOG DEVICES AD8606ARMZ-REEL",
     "AD8606ARMZ-REEL", "Analog Devices"),
    ("Диод NXP SEMICONDUCTORS BAS16H, 115", "BAS16H,115", "NXP Semiconductors"),
    ("Катушка индуктивности 0603CS-82NXGLU Coilcraft (0603-82 нГн2%)",
     "0603CS-82NXGLU", "Coilcraft"),
    ("Вилка М83513/01-AN Amphenol", "M83513/01-AN", "Amphenol"),
    ("Микросхема MT47H128M16RT-25E IT Micron", "MT47H128M16RT-25E", "Micron Technology"),
    ("Микросхема XC7A100T-2CSG324I Xilinx", "XC7A100T-2CSG324I", "Xilinx"),
    ("Микросхема STM32L451REI6 ST Microelectronics", "STM32L451REI6", "STMicroelectronics"),
    ("Конденсатор GRM188R60J476ME15 (0603-X5R-6,3 В-47 мкФ±20%) Murata",
     "GRM188R60J476ME15", "Murata"),
    ("Микросхема TPA6211A1DRBR Texas Instruments", "TPA6211A1DRBR", "Texas Instruments"),
    ("Микросхема LT3650IDD-8.4#PBF Linear Technology", "LT3650IDD-8.4#PBF", "Linear Technology"),
    ("Микросхема MAX44260AXT+ Maxim", "MAX44260AXT+", "Maxim Integrated"),
    ("Транзистор SiA408DJ Vishay", "SiA408DJ", "Vishay"),
]


class TestGolden(unittest.TestCase):
    def test_count(self):
        self.assertEqual(len(GOLDEN), 82, "в листе «360» ровно 82 позиции")

    def test_mpn_and_brand(self):
        bad = []
        for raw, want_mpn, want_brand in GOLDEN:
            part = m.parse_line(raw, 1)
            if part.mpn != want_mpn or part.manufacturer != want_brand:
                bad.append(f"  {raw!r}\n    ожидали: {want_mpn!r} / {want_brand!r}\n"
                           f"    вышло:   {part.mpn!r} / {part.manufacturer!r}")
        if bad:
            self.fail(f"расхождений {len(bad)} из {len(GOLDEN)}:\n" + "\n".join(bad))


class TestHomoglyphs(unittest.TestCase):
    def test_latinizes_when_latin_present(self):
        # есть латинские A и N → кириллическая М заведомо опечатка
        self.assertEqual(m.fix_homoglyphs("М83513/01-AN"), "M83513/01-AN")

    def test_keeps_russian_designations(self):
        # «Д» не гомоглиф — отечественный артикул трогать нельзя
        self.assertEqual(m.fix_homoglyphs("2Д510А"), "2Д510А")

    def test_keeps_all_cyrillic_designations(self):
        # все буквы — двойники латиницы, но латиницы в токене нет:
        # это настоящий отечественный артикул, а не опечатка
        self.assertEqual(m.fix_homoglyphs("КР580ВМ80А"), "КР580ВМ80А")
        self.assertEqual(m.fix_homoglyphs("МАХ232"), "МАХ232")

    def test_all_cyrillic_gets_latin_search_variant(self):
        # исходное написание сохраняется, но искать пробуем и латиницей
        variants = [v["mpn"] for v in m.build_variants("МАХ232")]
        self.assertEqual(variants[0], "МАХ232")
        self.assertIn("MAX232", variants)

    def test_leaves_pure_words_alone(self):
        self.assertEqual(m.fix_homoglyphs("Резистор"), "Резистор")


class TestParens(unittest.TestCase):
    def test_strips_only_space_separated(self):
        text, grabbed = m.strip_parens("GRM21BR71H105KA12 (0805-X7R-50 В)")
        self.assertEqual(text, "GRM21BR71H105KA12")
        self.assertEqual(grabbed, "0805-X7R-50 В")

    def test_keeps_attached_parens(self):
        # у Hirose «(71)» — часть артикула, отличает от него только пробел
        text, grabbed = m.strip_parens("FX11LA-80P/8-SV(71)")
        self.assertEqual(text, "FX11LA-80P/8-SV(71)")
        self.assertEqual(grabbed, "")


class TestVariants(unittest.TestCase):
    def test_root_strips_iteratively(self):
        self.assertEqual(m.root_of("SI1013CX-T1-GE3"), "SI1013CX")
        self.assertEqual(m.root_of("LT3650IDD-4.2#PBF"), "LT3650IDD-4.2")
        self.assertEqual(m.root_of("BAS16H,115"), "BAS16H")
        self.assertEqual(m.root_of("ADC-26-52+"), "ADC-26-52")

    def test_root_keeps_short_stems(self):
        # не должен обкусывать до неузнаваемости
        self.assertEqual(m.root_of("A04TG"), "A04TG")
        self.assertEqual(m.root_of("QPA4263A"), "QPA4263A")

    def test_exact_is_always_first(self):
        variants = m.build_variants("LT3650IDD-4.2#PBF")
        self.assertEqual(variants[0], {"id": "exact", "mpn": "LT3650IDD-4.2#PBF"})
        self.assertIn("LT3650IDD-4.2", [v["mpn"] for v in variants])

    def test_reel_suffix_variant(self):
        self.assertIn("TPA6211A1DRB", [v["mpn"] for v in m.build_variants("TPA6211A1DRBR")])


class TestMatchQuality(unittest.TestCase):
    def test_exact(self):
        self.assertEqual(m.match_quality("BAS16H", "BAS16H")[0], "exact")
        self.assertEqual(m.match_quality("bas-16h", "BAS16H")[0], "exact")

    def test_packaging_suffix(self):
        self.assertEqual(m.match_quality("LT3650IDD-4.2", "LT3650IDD-4.2#PBF")[0], "suffix")
        self.assertEqual(m.match_quality("BAS16H", "BAS16H,115")[0], "suffix")
        self.assertEqual(m.match_quality("AD8606ARMZ-REEL", "AD8606ARMZ")[0], "suffix")

    def test_mismatch(self):
        self.assertEqual(m.match_quality("BAS16H", "LM317T")[0], "mismatch")


class TestManufacturerMatch(unittest.TestCase):
    def test_acquisition_counts_as_match(self):
        self.assertEqual(m.manufacturer_match("Linear Technology", "Analog Devices"), "yes")
        self.assertEqual(m.manufacturer_match("Maxim Integrated", "Analog Devices"), "yes")

    def test_plain_mismatch(self):
        self.assertEqual(m.manufacturer_match("Murata", "Yageo"), "no")

    def test_unknown_when_absent(self):
        self.assertEqual(m.manufacturer_match("", "Murata"), "unknown")


class TestMasks(unittest.TestCase):
    """Маска регламента: x/Xx/XX в конце — любой суффикс, x в середине — один символ."""

    def test_trailing_masks(self):
        self.assertEqual(m.match_quality("3313J-1-104x", "3313J-1-104E")[0], "mask")
        self.assertEqual(m.match_quality("DLC70B8R2CW501Xx", "DLC70B8R2CW501XT")[0], "mask")
        self.assertEqual(m.match_quality("DLC70B8R2CW501XX", "DLC70B8R2CW501XK")[0], "mask")

    def test_single_capital_x_is_literal(self):
        # одиночная заглавная X — часть артикула, а не маска
        self.assertEqual(m.split_mask("EXB28V220JX"), ("EXB28V220JX", False))
        self.assertEqual(m.match_quality("EXB28V220JX", "EXB-28V220JX")[0], "exact")
        self.assertEqual(m.match_quality("EXB28V220JX", "EXB28V220JXDR", strict=True)[0], "partial")

    def test_midstring_x_wildcard(self):
        self.assertEqual(m.match_quality("TPSE477K010x0200", "TPSE477K010R0200")[0], "mask")
        self.assertEqual(m.match_quality("TPSE477K010x0200", "TPSE477K0100200")[0], "mismatch")
        self.assertEqual(m.match_quality("TPSE477K010x0200", "TPSE477K010RR0200")[0], "mismatch")

    def test_masked_mpn_gets_base_variant(self):
        variants = [v["mpn"] for v in m.build_variants("3313J-1-104x")]
        self.assertEqual(variants[0], "3313J-1-104x")
        self.assertIn("3313J-1-104", variants)


class TestSeriesPrefix(unittest.TestCase):
    def test_prefix_completion_with_packaging_tail(self):
        q, delta = m.match_quality("0805-FX-1503", "CR0805-FX-1503ELF")
        self.assertEqual((q, delta), ("prefix_series", "CR"))

    def test_short_or_unknown_tails_are_not_series(self):
        # короткий запрос не достраивается префиксом: это «частичное», а не серия
        self.assertEqual(m.match_quality("LM358", "PLM358")[0], "partial")
        self.assertEqual(m.match_quality("104", "CR104")[0], "partial")
        self.assertEqual(m.match_quality("LM358", "LM358DR", strict=True)[0], "partial")


class TestStrictSuffix(unittest.TestCase):
    def test_packaging_tail_accepted_in_both_modes(self):
        for strict in (False, True):
            self.assertEqual(m.match_quality("CR0805-JW-390E", "CR0805-JW-390ELF",
                                             strict=strict)[0], "suffix")
            self.assertEqual(m.match_quality("LT1763CS8", "LT1763CS8#TRPBF",
                                             strict=strict)[0], "suffix")

    def test_different_packaging_tails_on_both_sides(self):
        # запрос с reel-маркером T, кандидат с другим упаковочным хвостом
        self.assertEqual(m.match_quality("C5750X7R1H106KT", "C5750X7R1H106K230KB", strict=True)[0], "suffix")
        self.assertEqual(m.match_quality("TPS62065DSGR", "TPS62065DSGT", strict=True)[0], "suffix")
        # но не любой хвост: другой номинал остаётся чужим
        self.assertNotIn(m.match_quality("C5750X7R1H106KT", "C5750X7R1H226KT", strict=True)[0],
                         ("exact", "suffix", "mask", "prefix_series"))

    def test_chained_vendor_tails(self):
        # ADI: Z (RoHS) + катушка; TI: X + /NOPB; onsemi: X
        for req, off in (("AD8027ART", "AD8027ARTZ-R2"), ("OP467GS", "OP467GSZ-REEL"),
                         ("LMV321M5", "LMV321M5X/NOPB"), ("74ACT163MTC", "74ACT163MTCX"),
                         ("AD7945BR", "AD7945BRZ"), ("LT1121IST-5", "LT1121IST-5#TRPBF")):
            self.assertEqual(m.match_quality(req, off, strict=True)[0], "suffix", (req, off))
        # а вот X внутри артикула — не хвост: такое не принимается
        self.assertNotIn(m.match_quality("LM57BISD-5", "LM57BISDX-5/NOPB", strict=True)[0],
                         ("exact", "suffix", "mask", "prefix_series"))

    def test_unknown_short_tail_depends_on_mode(self):
        self.assertEqual(m.match_quality("BAS16H", "BAS16HQ")[0], "suffix")
        self.assertEqual(m.match_quality("BAS16H", "BAS16HQ", strict=True)[0], "partial")

    def test_formatting_only_differences_are_exact(self):
        self.assertEqual(m.match_quality("Q 10,240-SS3", "Q10.240-SS3")[0], "exact")
        self.assertEqual(m.match_quality("Q 10,240-SS3", "Q10.241-SS3")[0], "mismatch")
        self.assertEqual(m.match_quality("3310Y001-103L", "3310Y-001-103L")[0], "exact")


class TestCyrillicUnitsAndTokens(unittest.TestCase):
    """Кейсы из bom-price-parser: составные ячейки «тип + артикул + бренд»."""

    def test_resonator_keeps_prefix_decimal_and_slash(self):
        part = m.parse_line("РЕЗОНАТОР Q 10,240-SS3-16-30/50-FU-T1 Jauch", 300)
        self.assertEqual(part.mpn, "Q10,240-SS3-16-30/50-FU-T1")
        self.assertEqual(part.manufacturer, "Jauch")

    def test_huber_suhner_composite_code(self):
        part = m.parse_line("16 MMCX-50-2-13/111 OE Huber+Suhner", 10)
        self.assertEqual(part.mpn, "16 MMCX-50-2-13/111 OE")
        self.assertEqual(part.manufacturer, "Huber+Suhner")

    def test_relay_keeps_decimal_comma(self):
        part = m.parse_line("РЕЛЕ FTR-B4-CB-4,5Z FUJITSU", 5)
        self.assertEqual(part.mpn, "FTR-B4-CB-4,5Z")
        self.assertEqual(part.manufacturer, "Fujitsu")

    def test_catalog_part_beats_cyrillic_parameter(self):
        part = m.parse_line("РЕЗИСТОР 3310Y-001-10КОМ 3310Y001-103L Bourns", 1)
        self.assertEqual(part.mpn, "3310Y001-103L")

    def test_units_normalized_when_no_real_mpn(self):
        part = m.parse_line("РЕЗИСТОР 0603-200КОМ+-5% YAGEO", 1)
        self.assertEqual(part.mpn, "0603-200K+-5%")
        self.assertEqual(part.manufacturer, "Yageo")
        self.assertEqual(m.normalize_cyrillic_units("МС-306 32,768КГЦ"), "МС-306 32,768kHz")

    def test_brand_in_parentheses(self):
        part = m.parse_line("Вилка 44914-0801 (Molex)", 3500)
        self.assertEqual((part.mpn, part.manufacturer), ("44914-0801", "Molex"))
        part = m.parse_line("Фильтр SMA 142-0701-871 (Johnson Components)", 10)
        self.assertEqual((part.mpn, part.manufacturer), ("142-0701-871", "Johnson Components"))
        # параметры в скобках брендом не становятся
        part = m.parse_line("Конденсатор GRM1555C1H332GE01 (0402-C0G-50 В-3300 пФ±2%)", 1)
        self.assertEqual(part.manufacturer, "")

    def test_domestic_designation_untouched(self):
        part = m.parse_line("Микросхема КР580ВМ80А", 1)
        self.assertEqual(part.mpn, "КР580ВМ80А")
        self.assertEqual(part.mpn_script, "cyrillic")


class TestBrandCanon(unittest.TestCase):
    def test_alias_to_canon(self):
        self.assertEqual(m.canon_brand("Avago Technologies"), "Broadcom")
        self.assertEqual(m.canon_brand("ST Micro"), "STMicroelectronics")
        self.assertEqual(m.canon_brand("Unknown Corp"), "Unknown Corp")

    def test_typos_and_successors(self):
        self.assertEqual(m.manufacturer_match("Sypress", "Infineon Technologies"), "yes")
        self.assertEqual(m.manufacturer_match("Linear Tehnology", "Analog Devices Inc."), "yes")
        self.assertEqual(m.manufacturer_match("SKYWOKS SOLUTIONS", "Skyworks Solutions Inc."), "yes")
        self.assertEqual(m.manufacturer_match("NATIONAL SEMICONDUCTOR", "Texas Instruments"), "yes")
        self.assertEqual(m.manufacturer_match("PHILIPS", "Nexperia USA Inc."), "yes")
        # короткие и по-настоящему разные имена не сближаются
        self.assertEqual(m.manufacturer_match("Murata", "Mouser"), "no")
        self.assertEqual(m.manufacturer_match("Vishay", "Toshiba"), "no")

    def test_merged_brands_match(self):
        self.assertEqual(m.manufacturer_match("Epcos", "TDK Electronics / EPCOS"), "yes")
        self.assertEqual(m.manufacturer_match("Avago Technologies", "Broadcom Limited"), "yes")
        self.assertEqual(m.manufacturer_match("Fujitsu", "FCL Components"), "yes")
        self.assertEqual(m.manufacturer_match("Axicom", "TE Connectivity"), "yes")
        self.assertEqual(m.manufacturer_match("Jauch", "Yageo"), "no")


if __name__ == "__main__":
    unittest.main(verbosity=2)
