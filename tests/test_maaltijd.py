"""
Tests voor de onderdelen van maaltijd.py die geen scherm nodig hebben:
geld en datum, Excel lezen en schrijven, de ledenlijst en de opslag.

Draaien vanuit de map van het programma:
    python -m unittest discover -s tests -v

Alle namen hieronder zijn verzonnen. Gebruik in tests nooit echte leden.
"""

import json
import os
import sys
import tempfile
import unittest

# Zorgt dat "import maaltijd" werkt, ook als de tests vanuit een andere map starten.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import maaltijd as m  # noqa: E402  (moet na de sys.path regel hierboven)


class TestGeld(unittest.TestCase):
    def test_parse_prijs_komma_en_punt(self):
        self.assertEqual(m.parse_prijs("1,50"), 150)
        self.assertEqual(m.parse_prijs("1.50"), 150)
        self.assertEqual(m.parse_prijs(" € 2 "), 200)

    def test_parse_prijs_rondt_af_op_centen(self):
        self.assertEqual(m.parse_prijs("1,505"), 151)
        self.assertEqual(m.parse_prijs("1,504"), 150)

    def test_parse_prijs_weigert_onzin(self):
        for tekst in ("", "abc", "-1", "nan", "inf", "150000"):
            with self.subTest(tekst=tekst):
                with self.assertRaises(ValueError):
                    m.parse_prijs(tekst)

    def test_euro(self):
        self.assertEqual(m.euro(150), "1,50")
        self.assertEqual(m.euro(0), "0,00")
        self.assertEqual(m.euro(5), "0,05")
        self.assertEqual(m.euro(123456), "1.234,56")
        self.assertEqual(m.euro(-5), "-0,05")


class TestDatum(unittest.TestCase):
    def test_parse_datum(self):
        self.assertEqual(m.parse_datum("30-10-2026"), "30-10-2026")
        self.assertEqual(m.parse_datum(" 5/1/2026 "), "05-01-2026")

    def test_parse_datum_weigert_niet_bestaande_datum(self):
        for tekst in ("31-02-2026", "2026-10-30", "morgen"):
            with self.subTest(tekst=tekst):
                with self.assertRaises(ValueError):
                    m.parse_datum(tekst)

    def test_auto_datum(self):
        data = {"datum": "01-01-2000", "datum_handmatig": False}
        m.auto_datum(data)
        self.assertEqual(data["datum"], m.today_str())

        data = {"datum": "01-01-2000", "datum_handmatig": True}
        m.auto_datum(data)
        self.assertEqual(data["datum"], "01-01-2000")


class TestExcel(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = os.path.join(self.tmp.name, "test.xlsx")

    def tearDown(self):
        self.tmp.cleanup()

    def test_kolomletters(self):
        for index, letters in ((0, "A"), (25, "Z"), (26, "AA"), (701, "ZZ"), (702, "AAA")):
            with self.subTest(letters=letters):
                self.assertEqual(m._col_letter(index), letters)
                self.assertEqual(m._col_index(letters + "1"), index)

    def test_schrijven_en_teruglezen(self):
        rows = [
            [("Kop", m.STYLE_BOLD)],
            [],
            ["007", "A & B <c>", 12, 1.5],
            ["", "", "derde kolom"],
            ["stuur\x01code"],
        ]
        m.write_xlsx(self.path, [("Blad", rows)])
        self.assertEqual(
            m.read_xlsx_rows(self.path),
            [["Kop"], ["007", "A & B <c>", "12", "1.5"], ["", "", "derde kolom"], ["stuurcode"]],
        )
        self.assertFalse(os.path.exists(self.path + ".tmp"))

    def test_geen_xlsx(self):
        with open(self.path, "w") as f:
            f.write("dit is geen zip")
        with self.assertRaises(m.StoreError):
            m.read_xlsx_rows(self.path)


class TestLeden(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = os.path.join(self.tmp.name, "leden.xlsx")

    def tearDown(self):
        self.tmp.cleanup()

    def load(self, rows):
        m.write_xlsx(self.path, [("Leden", rows)])
        return [(x.nr, x.naam) for x in m.load_members(self.path)]

    def test_met_kopjes_voornaam_achternaam(self):
        rows = [["Lidnummer", "Voornaam", "Achternaam"], [1001, "Anna", "de Vries"], [1010, "Jelle", "Bos"]]
        self.assertEqual(self.load(rows), [("1001", "Anna de Vries"), ("1010", "Jelle Bos")])

    def test_zonder_kopjes(self):
        self.assertEqual(self.load([[1001, "Anna de Vries"]]), [("1001", "Anna de Vries")])

    def test_slaat_dubbele_en_lege_regels_over(self):
        rows = [["Nr", "Naam"], ["007", "Anna de Vries"], ["007", "Dubbel"], ["", "Zonder nummer"], ["A12", "Jelle Bos"]]
        self.assertEqual(self.load(rows), [("007", "Anna de Vries"), ("A12", "Jelle Bos")])

    def test_ontbrekend_bestand(self):
        with self.assertRaises(m.StoreError):
            m.load_members(os.path.join(self.tmp.name, "bestaat-niet.xlsx"))

    def test_zoeken(self):
        leden = [m.Member("1001", "Anna de Vries"), m.Member("1010", "Jelle Bos"), m.Member("2001", "Bas Anders")]

        def namen(tekst):
            return [x.naam for x in m.find_members(leden, tekst)]

        self.assertEqual(namen("1001"), ["Anna de Vries"])  # exact lidnummer
        self.assertEqual(namen("10"), ["Anna de Vries", "Jelle Bos"])  # begin van lidnummer
        self.assertEqual(namen("VRIES anna"), ["Anna de Vries"])  # alle woorden, hoofdletters maken niet uit
        self.assertEqual(namen("an"), ["Anna de Vries", "Bas Anders"])  # gesorteerd op naam
        self.assertEqual(namen("   "), [])


class TestOpslag(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = os.path.join(self.tmp.name, "soos.json")
        self.store = m.JsonStore(self.path, m.soos_default, m.auto_datum)

    def tearDown(self):
        self.tmp.cleanup()

    def test_leeg_bestand_geeft_standaardgegevens(self):
        self.assertEqual(self.store.read()["verbruik"], {})

    def test_update_schrijft_weg_en_ruimt_op(self):
        self.store.update(lambda data: data["drankjes"].append({"id": 1, "naam": "Bier", "cent": 150}))
        self.assertEqual(self.store.read()["drankjes"], [{"id": 1, "naam": "Bier", "cent": 150}])
        self.assertFalse(os.path.exists(self.path + ".lock"))
        self.assertFalse(os.path.exists(self.path + ".tmp"))

    def test_slot_wordt_ook_na_fout_losgelaten(self):
        def kapot(data):
            raise RuntimeError("test")

        with self.assertRaises(RuntimeError):
            self.store.update(kapot)
        self.assertFalse(os.path.exists(self.path + ".lock"))

    def test_kapot_bestand_wordt_apart_gezet(self):
        with open(self.path, "w") as f:
            f.write("{dit is geen json")
        self.assertEqual(self.store.read()["drankjes"], [])
        kapot = [n for n in os.listdir(self.tmp.name) if n.startswith("soos.json.kapot-")]
        self.assertEqual(len(kapot), 1)

    def test_bezet_slot_geeft_nette_fout(self):
        open(self.path + ".lock", "w").close()  # een vers slot van "een ander scherm"
        with self.assertRaises(m.StoreError):
            with m.FileLock(self.path + ".lock", timeout=0.2):
                pass

    def test_rekening(self):
        soos = {
            "drankjes": [{"id": 1, "naam": "Bier", "cent": 150}, {"id": 2, "naam": "Fris", "cent": 100}],
            "verbruik": {"1001": {"naam": "Anna de Vries", "drank": {"1": 3, "2": 1}}},
        }
        self.assertEqual(m.member_total_cents(soos, "1001"), 550)
        self.assertEqual(m.member_total_cents(soos, "9999"), 0)

    def test_excel_nr(self):
        self.assertEqual(m.excel_nr("1001"), 1001)
        self.assertEqual(m.excel_nr("007"), "007")
        self.assertEqual(m.excel_nr("A12"), "A12")


if __name__ == "__main__":
    unittest.main()
