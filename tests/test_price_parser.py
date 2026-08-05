import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import openpyxl

import browser_bom as bbom
import price_parser as parser


class PartNumberExtractionTests(unittest.TestCase):
    def test_jauch_part_keeps_prefix_decimal_and_slash(self):
        value = "РЕЗОНАТОР Q 10,240-SS3-16-30/50-FU-T1 Jauch"
        self.assertEqual(
            parser.extract_pn_from_description(value, "Jauch"),
            "Q 10,240-SS3-16-30/50-FU-T1",
        )

    def test_cyrillic_units_are_normalized_without_losing_model(self):
        self.assertEqual(
            parser.extract_pn_from_description("РЕЗОНАТОР МС-306 32,768КГЦ Epson", "Epson"),
            "MC-306 32,768kHz",
        )
        self.assertEqual(
            parser.extract_pn_from_description("РЕЗИСТОР 0603-200КОМ+-5% YAGEO", "YAGEO"),
            "0603-200K+-5%",
        )

    def test_explicit_catalog_part_wins_over_cyrillic_description(self):
        value = "РЕЗИСТОР 3310Y-001-10КОМ 3310Y001-103L Bourns"
        self.assertEqual(parser.extract_pn_from_description(value, "Bourns"), "3310Y001-103L")

    def test_decimal_comma_and_full_huber_suhner_part_are_preserved(self):
        self.assertEqual(
            parser.extract_pn_from_description("РЕЛЕ FTR-B4-CB-4,5Z FUJITSU", "FUJITSU"),
            "FTR-B4-CB-4,5Z",
        )
        self.assertEqual(
            parser.extract_pn_from_description("16 MMCX-50-2-13/111 OE Huber+Suhner", "Huber+Suhner"),
            "16 MMCX-50-2-13/111 OE",
        )

    def test_read_request_cleans_combined_partnumber_column(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "request.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["№", "Наименование", "Partnumber", "BRAND", "QTY"])
            ws.append([1, "Резонатор", "РЕЗОНАТОР Q 10,240-SS3-16-30/50-FU-T1 Jauch", "Jauch", 300])
            wb.save(path)

            items = parser.read_request_xlsx(path)

        self.assertEqual(items[0]["pn"], "Q 10,240-SS3-16-30/50-FU-T1")
        self.assertEqual(items[0]["manufacturer"], "Jauch")


class SearchValidationTests(unittest.TestCase):
    def test_mpn_match_allows_only_formatting_variation(self):
        self.assertTrue(parser.mpn_matches("Q 10,240-SS3", "Q10.240-SS3"))
        self.assertFalse(parser.mpn_matches("Q 10,240-SS3", "Q10.241-SS3"))

    def test_trailing_masks_per_regulament(self):
        # строчный x, смешанный Xx и заглавные XX — маска «любой суффикс»
        self.assertTrue(parser.mpn_matches("3313J-1-104x", "3313J-1-104E"))
        self.assertTrue(parser.mpn_matches("DLC70B8R2CW501Xx", "DLC70B8R2CW501XT"))
        self.assertTrue(parser.mpn_matches("DLC70B8R2CW501XX", "DLC70B8R2CW501XK"))
        # одиночная заглавная X — литерал, не маска: произвольный (неупаковочный)
        # суффикс не принимается; T принялся бы по правилу упаковки — это норм
        self.assertFalse(parser.mpn_matches("EXB28V220JX", "EXB28V220JXDR"))
        self.assertTrue(parser.mpn_matches("EXB28V220JX", "EXB-28V220JX"))

    def test_midstring_x_wildcard(self):
        self.assertTrue(parser.mpn_matches("TPSE477K010x0200", "TPSE477K010R0200"))
        self.assertFalse(parser.mpn_matches("TPSE477K010x0200", "TPSE477K0100200"))
        self.assertFalse(parser.mpn_matches("TPSE477K010x0200", "TPSE477K010RR0200"))

    def test_packaging_tail_and_series_prefix(self):
        self.assertTrue(parser.mpn_matches("CR0805-JW-390E", "CR0805-JW-390ELF"))
        self.assertTrue(parser.mpn_matches("0805-FX-1503", "CR0805-FX-1503ELF"))
        self.assertFalse(parser.mpn_matches("LM358", "LM358DR"))
        self.assertFalse(parser.mpn_matches("LM358", "PLM358"))
        self.assertFalse(parser.mpn_matches("104", "CR104"))

    def test_manufacturer_aliases_cover_merged_brands(self):
        self.assertTrue(parser.manufacturer_matches("Epcos", "TDK Electronics / EPCOS"))
        self.assertTrue(parser.manufacturer_matches("Avago Technologies", "Broadcom Limited"))
        self.assertFalse(parser.manufacturer_matches("Jauch", "Yageo"))

    def test_price_search_rejects_other_manufacturer(self):
        wrong = {
            "dist": "DigiKey", "pn": "WRONG-1", "stock": 1000, "moq": 1,
            "lead": None, "tiers": [{"qty": 1, "price": 0.01}], "manu": "Yageo",
        }
        correct = {
            "dist": "DigiKey", "pn": "RIGHT-1", "stock": 1000, "moq": 1,
            "lead": None, "tiers": [{"qty": 1, "price": 0.02}], "manu": "Jauch",
        }

        def digikey(query, _qty, match_pn=None):
            return [correct] if "." in query else [wrong]

        with patch.object(parser, "search_digikey", side_effect=digikey), \
                patch.object(parser, "search_mouser", return_value=[]), \
                patch.object(parser, "search_farnell", return_value=[]), \
                patch.object(parser, "search_tme", return_value=[]):
            result = parser.find_best_price("Q 10,240-SS3", 300, "Jauch")

        self.assertEqual(result["status"], "FOUND")
        self.assertEqual(result["manufacturer"], "Jauch")
        self.assertEqual(result["distr_pn"], "RIGHT-1")

    def test_transient_timeout_is_retried(self):
        class Response:
            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

            def read(self):
                return b"ok"

        with patch("urllib.request.urlopen", side_effect=[TimeoutError(), Response()]), \
                patch.object(parser.time, "sleep"):
            self.assertEqual(parser.fetch_bytes("request"), b"ok")

    def test_once_mode_keeps_required_manufacturer(self):
        args = SimpleNamespace(once=["ASSR-1530-005E", "200"], manufacturer="Avago")
        self.assertEqual(parser.load_items(args)[0]["manufacturer"], "Avago")


OEMSECRETS_HEADERS = [
    "Part Number", "Quantity for Single BOM", "Manufacturer", "Distributor",
    "Minimum Order", "Stock", "Lead Time on Additional Stock in Weeks",
    "Unit Price in USD", "Distributor SKU",
]


def _mk_row(**kw):
    """Нормализованная строка экспорта площадки (контракт read_site_export)."""
    row = {
        "req_pn": "CR0805-JW-390E", "matched_mpn": "", "manufacturer": "Bourns",
        "distributor": "DigiKey", "price": 0.01, "currency": "USD",
        "break_qty": 1, "moq": 1, "stock": 5000, "lead": None, "sku": "SKU-1",
    }
    row.update(kw)
    return row


def _mk_results():
    """Один FOUND (API) + один RFQ с производителем + один RFQ без."""
    return [
        {"pn": "LM358D", "qty": 10, "manufacturer": "TI", "description": "LM358D",
         "status": "FOUND", "distributor": "Mouser", "price_usd": 0.5,
         "stock": 100, "moq": 1, "lead": None, "in_stock": True,
         "resolved_mpn": "", "distr_pn": ""},
        {"pn": "CR0805-JW-390E", "qty": 100, "manufacturer": "Bourns",
         "description": "CR0805-JW-390E", "status": "RFQ"},
        {"pn": "MYSTERY-1", "qty": 5, "manufacturer": "",
         "description": "MYSTERY-1", "status": "RFQ"},
    ]


class BrowserExportParsingTests(unittest.TestCase):
    def test_oemsecrets_headers_resolve_all_fields(self):
        colmap, currency = bbom.resolve_columns(OEMSECRETS_HEADERS, "oemsecrets")
        for field in ("req_pn", "manufacturer", "distributor", "price",
                      "moq", "stock", "lead", "sku", "break_qty"):
            self.assertIn(field, colmap, field)
        self.assertEqual(currency, "USD")  # валюта из заголовка 'Unit Price in USD'

    def test_distributor_part_number_does_not_steal_part_number(self):
        headers = ["Part Number", "Distributor Part Number", "Manufacturer",
                   "Distributor", "Unit Price"]
        colmap, _ = bbom.resolve_columns(headers, "octopart")
        self.assertEqual(colmap["req_pn"], 0)
        self.assertEqual(colmap["sku"], 1)   # длинный синоним выигрывает

    def test_unknown_headers_raise_and_name_them(self):
        with self.assertRaises(bbom.BrowserExportFormatError) as ctx:
            bbom.resolve_columns(["Foo", "Bar", "Baz"], "findchips")
        self.assertIn("Foo", str(ctx.exception))
        self.assertIn("aliases", str(ctx.exception))

    def test_echo_only_export_refused_without_trust_flag(self):
        results = _mk_results()
        rows = [_mk_row(matched_mpn="")]
        stats = bbom.browser_merge_pass(results, rows, "oemsecrets",
                                        echo_only=True, trust_echo=False,
                                        default_currency="USD")
        self.assertEqual(stats["accepted"], 0)
        self.assertEqual(stats["rejected"], {"echo_only": 1})
        self.assertEqual(results[1]["status"], "RFQ")


class BrowserMergeValidationTests(unittest.TestCase):
    def merge(self, rows, results=None, **kw):
        results = results if results is not None else _mk_results()
        kw.setdefault("default_currency", "USD")
        kw.setdefault("trust_echo", True)
        stats = bbom.browser_merge_pass(results, rows, "oemsecrets", **kw)
        return results, stats

    def test_missing_manufacturer_rejected(self):
        _, stats = self.merge([_mk_row(manufacturer="")])
        self.assertEqual(stats["rejected"], {"no_manufacturer": 1})

    def test_wrong_manufacturer_rejected(self):
        _, stats = self.merge([_mk_row(manufacturer="Yageo")])
        self.assertEqual(stats["rejected"], {"wrong_manufacturer": 1})

    def test_analog_mpn_rejected(self):
        results = _mk_results()
        results[1]["pn"] = "ERJP06D56R0V"
        results[1]["manufacturer"] = "Panasonic"
        _, stats = self.merge(
            [_mk_row(req_pn="ERJP06D56R0V", matched_mpn="ERJP06D56R0X",
                     manufacturer="Panasonic")],
            results=results)
        self.assertEqual(stats["rejected"], {"analog_mpn": 1})

    def test_packaging_suffix_accepted_lands_in_column_b(self):
        results, stats = self.merge([_mk_row(matched_mpn="CR0805-JW-390ELF")])
        self.assertEqual(stats["accepted"], 1)
        self.assertEqual(results[1]["resolved_mpn"], "CR0805-JW-390ELF")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "out.xlsx"
            parser.write_results(results, path, 80.0)
            ws = openpyxl.load_workbook(path).active
            # строка 3 (после FOUND-строки): колонка B = найденный артикул
            self.assertEqual(ws.cell(row=3, column=2).value, "CR0805-JW-390ELF")
            self.assertEqual(str(ws.cell(row=3, column=2).fill.fgColor.rgb)[-6:],
                             "E2EFDA")

    def test_broker_distributor_rejected(self):
        _, stats = self.merge([_mk_row(distributor="XYZ Surplus Trading")])
        self.assertEqual(stats["rejected"], {"broker": 1})

    def test_denylist_beats_whitelist_substring(self):
        _, stats = self.merge([_mk_row(distributor="Digikey Surplus Brokers")])
        self.assertEqual(stats["rejected"], {"broker": 1})

    def test_unknown_distributor_rejected_not_defaulted(self):
        _, stats = self.merge([_mk_row(distributor="Acme Components")])
        self.assertEqual(stats["rejected"], {"unknown_distributor": 1})

    def test_merge_never_overwrites_found_row(self):
        results, stats = self.merge(
            [_mk_row(req_pn="LM358D", manufacturer="TI", price=0.1)])
        self.assertEqual(stats["not_pending"], 1)
        self.assertEqual(stats["accepted"], 0)
        self.assertEqual(results[0]["price_usd"], 0.5)   # цена API не тронута

    def test_no_manufacturer_rfq_row_is_not_pending(self):
        # позиция без производителя — RFQ навсегда, каскад её не трогает
        results, stats = self.merge(
            [_mk_row(req_pn="MYSTERY-1", manufacturer="NoName")])
        self.assertEqual(stats["not_pending"], 1)
        self.assertEqual(results[2]["status"], "RFQ")

    def test_eur_converted_once_per_run(self):
        rows = [_mk_row(currency="EUR", price=2.0),
                _mk_row(currency="EUR", price=3.0, distributor="Mouser"),
                _mk_row(currency="EUR", price=4.0, distributor="TME")]
        with patch.object(parser, "get_eur_usd", return_value=1.10) as mock_rate:
            results, stats = self.merge(rows)
        self.assertEqual(mock_rate.call_count, 1)   # курс взят один раз на прогон
        self.assertEqual(stats["accepted"], 1)
        self.assertAlmostEqual(results[1]["price_usd"], 2.0 * 1.10, places=4)

    def test_unresolvable_currency_refused(self):
        # валюты нет ни в строке, ни в шапке; дефолт площадки неизвестен
        _, stats = self.merge([_mk_row(currency=None)], default_currency=None)
        self.assertEqual(stats["rejected"], {"bad_currency": 1})

    def test_equal_price_larger_stock_wins(self):
        rows = [_mk_row(stock=10, distributor="Mouser", sku="S-SMALL"),
                _mk_row(stock=10000, distributor="DigiKey", sku="S-BIG")]
        results, stats = self.merge(rows)
        self.assertEqual(stats["accepted"], 1)
        self.assertEqual(results[1]["distributor"], "DigiKey")   # регламент §4

    def test_read_results_roundtrip_preserves_rate_and_flags(self):
        results = _mk_results()
        results[1] = {"pn": "CR0805-JW-390E", "qty": 100, "manufacturer": "Bourns",
                      "description": "x", "status": "FOUND", "distributor": "TME",
                      "price_usd": 0.02, "stock": 500, "moq": 1, "lead": None,
                      "in_stock": True, "resolved_mpn": "", "distr_pn": "",
                      "browser_assisted": True, "browser_source": "oemsecrets"}
        results.append({"pn": "MAX481ESA", "qty": 2, "manufacturer": "Maxim",
                        "description": "x", "status": "FOUND", "distributor": "DigiKey",
                        "price_usd": 6.74, "stock": 10, "moq": 1, "lead": None,
                        "in_stock": True, "resolved_mpn": "MAX481ESA+T",
                        "distr_pn": "", "llm_assisted": True})
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "out.xlsx"
            parser.write_results(results, path, 83.4567)
            back, rate = bbom.read_results(path)
        self.assertAlmostEqual(rate, 83.4567, places=4)
        self.assertEqual(len(back), 4)
        self.assertTrue(back[1]["browser_assisted"])
        self.assertTrue(back[3]["llm_assisted"])
        self.assertEqual(back[2]["status"], "RFQ")
        self.assertEqual(back[0]["status"], "FOUND")

    def test_browser_row_gets_distinct_fill_and_honest_footnote(self):
        results, _ = self.merge([_mk_row()])
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "out.xlsx"
            parser.write_results(results, path, 80.0)
            ws = openpyxl.load_workbook(path).active
            self.assertEqual(str(ws.cell(row=3, column=1).fill.fgColor.rgb)[-6:],
                             "E4DFEC")
            note = ws.cell(row=len(results) + 3, column=1).value
        self.assertIn("oemsecrets: 1", note)
        self.assertNotIn("только официальные API", note)
        self.assertIn("без подтверждения API: 1", note)

    def test_cascade_pending_shrinks_by_accepted_count(self):
        results = _mk_results()
        before = bbom.collect_pending(results)
        self.assertEqual(len(before), 1)          # без производителя — не в каскад
        results, stats = self.merge([_mk_row()], results=results)
        self.assertEqual(stats["accepted"], 1)
        self.assertEqual(len(bbom.collect_pending(results)), 0)


if __name__ == "__main__":
    unittest.main()
