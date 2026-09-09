# -*- coding: utf-8 -*-
"""Тесты регламента отбора (select.py) и сбора (collect.py) на заглушках площадок.

Запуск: python -m unittest discover tests   (из корня проекта)
"""

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from bomprice import bomread, collect as collectmod, fx as fxmod, registry  # noqa: E402
from bomprice import select as selectmod  # noqa: E402

RATES = fxmod.Rates(usd_rub=84.0, usd_rub_cbr=80.0, markup=4.0,
                    per_usd={"EUR": 0.9, "GBP": 0.75}, at="2026-09-09")
CFG = {"accept_quality": ["exact", "mask", "suffix", "prefix_series"],
       "authorized_sellers": ["digikey", "mouser", "tti", "arrow"], "resolved_limit": 3}


def src(id="digikey", kind="franchise", **kw):
    base = dict(id=id, title=id.title(), url="https://x", tier="api", adapter="digikey",
                kind=kind, gives=["price", "stock"])
    base.update(kw)
    return registry.Source(**base)


DIGIKEY, MOUSER = src("digikey"), src("mouser", adapter="mouser")
OEM = src("oemsecrets", kind="aggregator", adapter="oemsecrets")
BY_ID = {s.id: s for s in (DIGIKEY, MOUSER, OEM)}


def part(pn="Q 10,240-SS3", qty=300, mfr="Jauch"):
    return bomread.literal_part(1, pn, qty, mfr)


def offer(source="digikey", pn="Q10.240-SS3", mfr="Jauch", price=0.02, stock=1000, **kw):
    from bomprice import mpn as mpnmod, offer as offermod
    p = kw.pop("part", part())
    quality, delta = mpnmod.match_quality(p.mpn, pn, strict=True)
    rec = offermod.new(
        part_id=1, req_pn=p.mpn, req_qty=p.qty, source_id=source, source_title=source,
        seller=kw.pop("seller", source), seller_kind=kw.pop("seller_kind", "franchise"),
        offer_pn=pn, offer_mpn=pn, sku=kw.pop("sku", f"{source}-{pn}"), manufacturer=mfr,
        manufacturer_match=mpnmod.manufacturer_match(p.manufacturer, mfr),
        match_quality=quality, match_delta=delta, stock_qty=stock, moq=kw.pop("moq", 1),
        lead_time_days=kw.pop("lead_days", None),
        price_breaks=kw.pop("breaks", [{"qty": 1, "price": price}]),
        currency=kw.pop("currency", "USD"), variant="exact")
    rec.update(kw)
    return rec


class TestAcceptance(unittest.TestCase):
    def test_rejects_other_manufacturer(self):
        wrong = offer(pn="WRONG-1", mfr="Yageo", price=0.01)
        right = offer(pn="Q10.240-SS3", mfr="Jauch", price=0.02, sku="RIGHT-1")
        result = selectmod.best([wrong, right], part(), RATES, BY_ID, CFG)
        self.assertEqual(result["status"], "FOUND")
        self.assertEqual(result["distr_pn"], "RIGHT-1")
        self.assertEqual(result["manufacturer"], "Jauch")

    def test_partial_match_is_rfq(self):
        p = part("LM358", 10, "Texas Instruments")
        rec = offer(pn="LM358DR", mfr="Texas Instruments", part=p)
        self.assertEqual(rec["match_quality"], "partial")
        result = selectmod.best([rec], p, RATES, BY_ID, CFG)
        self.assertEqual(result["status"], "RFQ")
        self.assertIn("артикул", result["reason"])

    def test_mask_and_packaging_suffix_accepted(self):
        p = part("DLC70B8R2CW501Xx", 100, "Dalicap")
        rec = offer(pn="DLC70B8R2CW501XT", mfr="Dalicap", part=p)
        self.assertEqual(selectmod.best([rec], p, RATES, BY_ID, CFG)["resolved_mpn"], "DLC70B8R2CW501XT")
        p = part("CR0805-JW-390E", 100, "Bourns")
        rec = offer(pn="CR0805-JW-390ELF", mfr="Bourns", part=p)
        self.assertEqual(selectmod.best([rec], p, RATES, BY_ID, CFG)["status"], "FOUND")

    def test_aggregator_seller_filter(self):
        broker = offer(source="oemsecrets", seller="Some Broker Ltd", seller_kind="broker")
        authorized = offer(source="oemsecrets", seller="TTI Europe", seller_kind="broker", sku="tti-1")
        result = selectmod.best([broker, authorized], part(), RATES, BY_ID, CFG)
        self.assertEqual(result["status"], "FOUND")
        self.assertEqual(result["distributor"], "TTI Europe")

    def test_aggregator_without_manufacturer_needs_exact(self):
        p = part("BAS16H", 100, "Nexperia")
        exact = offer(source="oemsecrets", seller="Mouser", pn="BAS16H", mfr="", part=p)
        suffix = offer(source="oemsecrets", seller="Mouser", pn="BAS16H,115", mfr="", part=p, sku="s2")
        self.assertTrue(selectmod.acceptable(exact, p, OEM, CFG)[0])
        self.assertFalse(selectmod.acceptable(suffix, p, OEM, CFG)[0])

    def test_merged_brands_match(self):
        p = part("LT1763CS8#PBF", 10, "Linear Technology")
        rec = offer(pn="LT1763CS8#PBF", mfr="Analog Devices Inc.", part=p)
        self.assertEqual(selectmod.best([rec], p, RATES, BY_ID, CFG)["status"], "FOUND")


class TestBestChoice(unittest.TestCase):
    def test_stock_beats_price(self):
        cheap_no_stock = offer(price=0.01, stock=0, sku="a")
        pricier_in_stock = offer(source="mouser", price=0.02, stock=500, sku="b")
        result = selectmod.best([cheap_no_stock, pricier_in_stock], part(), RATES, BY_ID, CFG)
        self.assertEqual(result["distr_pn"], "b")
        self.assertTrue(result["in_stock"])

    def test_equal_price_prefers_bigger_stock(self):
        a = offer(price=0.02, stock=400, sku="a")
        b = offer(source="mouser", price=0.02, stock=9000, sku="b")
        self.assertEqual(selectmod.best([a, b], part(), RATES, BY_ID, CFG)["distr_pn"], "b")

    def test_price_break_and_foreign_currency(self):
        rec = offer(currency="EUR", breaks=[{"qty": 1, "price": 0.9}, {"qty": 100, "price": 0.45}], stock=10)
        result = selectmod.best([rec], part(), RATES, BY_ID, CFG)
        self.assertAlmostEqual(result["price_usd"], 0.5, places=4)
        self.assertFalse(result["in_stock"])

    def test_lead_weeks_only_when_not_in_stock(self):
        rec = offer(stock=10, lead_days=70)
        self.assertEqual(selectmod.best([rec], part(), RATES, BY_ID, CFG)["lead"], 10)
        rec = offer(stock=1000, lead_days=70)
        self.assertIsNone(selectmod.best([rec], part(), RATES, BY_ID, CFG)["lead"])

    def test_moq_kept_for_footnote(self):
        rec = offer(moq=500)
        result = selectmod.best([rec], part(), RATES, BY_ID, CFG)
        self.assertEqual(result["moq"], 500)


class TestCollect(unittest.TestCase):
    def test_variants_stop_at_first_hit_and_resolution_pass(self):
        p = part("BAS16H", 100, "Nexperia")
        calls = []

        def fake_adapter(src, query):
            calls.append((src.id, query))
            if query == "BAS16H":
                return [{"offer_mpn": "BAS16H,115", "manufacturer": "Nexperia", "stock_qty": 100,
                         "price_breaks": [{"qty": 1, "price": 0.05}]}], {}
            return [], {}

        stats = collectmod.Stats()
        with patch.object(collectmod.adapters, "get", return_value=fake_adapter):
            offers = collectmod.collect_with_resolution(p, [DIGIKEY], stats, {"exact", "suffix"},
                                                        use_cache=False, log=lambda *_: None)
        self.assertEqual(len(offers), 1)
        self.assertEqual(offers[0]["match_quality"], "suffix")
        # первый вариант дал результат — второй не спрашивали; затем повтор найденным точным p/n
        self.assertEqual(calls, [("digikey", "BAS16H"), ("digikey", "BAS16H,115")])

    def test_adapter_error_does_not_break_run(self):
        def boom(src, query):
            raise RuntimeError("HTTP 500")

        stats = collectmod.Stats()
        with patch.object(collectmod.adapters, "get", return_value=boom):
            offers = collectmod.collect_part(part(), [DIGIKEY], stats, use_cache=False,
                                             log=lambda *_: None)
        self.assertEqual(offers, [])
        self.assertEqual(stats.errors, 1)
        self.assertIn("HTTP 500", stats.by_source["digikey"]["error_text"])

    def test_tme_alt_symbols_pick_best(self):
        p = part("LT1763CS8#PBF", 5, "Analog Devices")
        raw = {"offer_mpn": "LT1763CS8", "alt_mpns": ["LT1763CS8#PBF", "LT1763CS8#TRPBF"],
               "manufacturer": "Analog Devices", "price_breaks": [{"qty": 1, "price": 1}]}
        rec = collectmod._finalize(raw, p, src("tme", adapter="tme"), "exact", "ref")
        self.assertEqual(rec["offer_pn"], "LT1763CS8#PBF")
        self.assertEqual(rec["match_quality"], "exact")


if __name__ == "__main__":
    unittest.main(verbosity=2)
