"""
Tests voor de onderdelen van maaltijd_app die geen scherm nodig hebben:
geld en datum, Excel lezen en schrijven, de ledenlijst en de opslag.

Draaien vanuit de map van het programma:
    python -m unittest discover -s tests -v

Alle namen hieronder zijn verzonnen. Gebruik in tests nooit echte leden.
"""

import json
import os
import struct
import sys
import tempfile
import unittest
from unittest import mock

# Zorgt dat "import maaltijd_app" werkt, ook als de tests vanuit een andere map starten.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from maaltijd_app import excel, hulpjes, instellingen, leden, opslag  # noqa: E402  (na de sys.path regel)


class TestGeld(unittest.TestCase):
    def test_parse_prijs_komma_en_punt(self):
        self.assertEqual(hulpjes.parse_prijs("1,50"), 150)
        self.assertEqual(hulpjes.parse_prijs("1.50"), 150)
        self.assertEqual(hulpjes.parse_prijs(" € 2 "), 200)

    def test_parse_prijs_rondt_af_op_centen(self):
        self.assertEqual(hulpjes.parse_prijs("1,505"), 151)
        self.assertEqual(hulpjes.parse_prijs("1,504"), 150)

    def test_parse_prijs_weigert_onzin(self):
        for tekst in ("", "abc", "-1", "nan", "inf", "150000"):
            with self.subTest(tekst=tekst):
                with self.assertRaises(ValueError):
                    hulpjes.parse_prijs(tekst)

    def test_euro(self):
        self.assertEqual(hulpjes.euro(150), "1,50")
        self.assertEqual(hulpjes.euro(0), "0,00")
        self.assertEqual(hulpjes.euro(5), "0,05")
        self.assertEqual(hulpjes.euro(123456), "1.234,56")
        self.assertEqual(hulpjes.euro(-5), "-0,05")


class TestDatum(unittest.TestCase):
    def test_parse_datum(self):
        self.assertEqual(hulpjes.parse_datum("30-10-2026"), "30-10-2026")
        self.assertEqual(hulpjes.parse_datum(" 5/1/2026 "), "05-01-2026")

    def test_parse_datum_weigert_niet_bestaande_datum(self):
        for tekst in ("31-02-2026", "2026-10-30", "morgen"):
            with self.subTest(tekst=tekst):
                with self.assertRaises(ValueError):
                    hulpjes.parse_datum(tekst)

    def test_auto_datum(self):
        data = {"datum": "01-01-2000", "datum_handmatig": False}
        opslag.auto_datum(data)
        self.assertEqual(data["datum"], hulpjes.today_str())

        data = {"datum": "01-01-2000", "datum_handmatig": True}
        opslag.auto_datum(data)
        self.assertEqual(data["datum"], "01-01-2000")

    def test_auto_datum_blijft_staan_zodra_er_iets_op_de_lijst_staat(self):
        meal = {"datum": "04-11-2026", "datum_handmatig": False, "aanmeldingen": {"1001": {"naam": "Anna de Vries", "aantal": 1}}}
        opslag.auto_datum(meal)
        self.assertEqual(meal["datum"], "04-11-2026")

        soos = {"datum": "04-11-2026", "datum_handmatig": False, "verbruik": {"1001": {"naam": "Anna de Vries", "drank": {"1": 2}}}}
        opslag.auto_datum(soos)
        self.assertEqual(soos["datum"], "04-11-2026")

        # Alles terug naar 0 geboekt telt als leeg: dan volgt de datum weer vandaag.
        leeg = {"datum": "04-11-2026", "datum_handmatig": False, "verbruik": {"1001": {"naam": "Anna de Vries", "drank": {}}}}
        opslag.auto_datum(leeg)
        self.assertEqual(leeg["datum"], hulpjes.today_str())

    def test_soos_na_middernacht_houdt_de_datum_van_de_avond(self):
        with tempfile.TemporaryDirectory() as tmp:
            store = opslag.JsonStore(os.path.join(tmp, "soos.json"), opslag.soos_default, opslag.auto_datum)
            vandaag = {"datum": "04-11-2026"}
            with mock.patch.object(opslag, "today_str", lambda: vandaag["datum"]):
                self.assertEqual(store.read()["datum"], "04-11-2026")
                store.update(lambda d: d["verbruik"].setdefault("1001", {"naam": "Anna de Vries", "drank": {"1": 1}}))
                vandaag["datum"] = "05-11-2026"  # middernacht
                self.assertEqual(store.read()["datum"], "04-11-2026")
                store.update(lambda d: d["verbruik"]["1001"]["drank"].update({"1": 2}))
                self.assertEqual(store.read()["datum"], "04-11-2026")
                # De knop Vandaag zet de datum bewust op de nieuwe dag.
                store.update(opslag.set_auto_datum())
                self.assertEqual(store.read()["datum"], "05-11-2026")
                self.assertFalse(store.read()["datum_handmatig"])


class TestExcel(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = os.path.join(self.tmp.name, "test.xlsx")

    def tearDown(self):
        self.tmp.cleanup()

    def test_kolomletters(self):
        for index, letters in ((0, "A"), (25, "Z"), (26, "AA"), (701, "ZZ"), (702, "AAA")):
            with self.subTest(letters=letters):
                self.assertEqual(excel._col_letter(index), letters)
                self.assertEqual(excel._col_index(letters + "1"), index)

    def test_schrijven_en_teruglezen(self):
        rows = [
            [("Kop", excel.STYLE_BOLD)],
            [],
            ["007", "A & B <c>", 12, 1.5],
            ["", "", "derde kolom"],
            ["stuur\x01code"],
        ]
        excel.write_xlsx(self.path, [("Blad", rows)])
        self.assertEqual(
            excel.read_xlsx_rows(self.path),
            [["Kop"], ["007", "A & B <c>", "12", "1.5"], ["", "", "derde kolom"], ["stuurcode"]],
        )
        self.assertFalse(os.path.exists(self.path + ".tmp"))

    def test_geen_xlsx(self):
        with open(self.path, "w") as f:
            f.write("dit is geen zip")
        with self.assertRaises(instellingen.StoreError):
            excel.read_xlsx_rows(self.path)


class TestLeden(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = os.path.join(self.tmp.name, "leden.xlsx")

    def tearDown(self):
        self.tmp.cleanup()

    def load(self, rows):
        excel.write_xlsx(self.path, [("Leden", rows)])
        return [(x.nr, x.naam) for x in leden.load_members(self.path)]

    def test_met_kopjes_voornaam_achternaam(self):
        rows = [["Lidnummer", "Voornaam", "Achternaam"], [1001, "Anna", "de Vries"], [1010, "Jelle", "Bos"]]
        self.assertEqual(self.load(rows), [("1001", "Anna de Vries"), ("1010", "Jelle Bos")])

    def test_zonder_kopjes(self):
        self.assertEqual(self.load([[1001, "Anna de Vries"]]), [("1001", "Anna de Vries")])

    def test_slaat_dubbele_en_lege_regels_over(self):
        rows = [["Nr", "Naam"], ["007", "Anna de Vries"], ["007", "Dubbel"], ["", "Zonder nummer"], ["A12", "Jelle Bos"]]
        self.assertEqual(self.load(rows), [("007", "Anna de Vries"), ("A12", "Jelle Bos")])

    def test_ontbrekend_bestand(self):
        with self.assertRaises(instellingen.StoreError):
            leden.load_members(os.path.join(self.tmp.name, "bestaat-niet.xlsx"))

    def test_zoeken(self):
        lijst = [leden.Member("1001", "Anna de Vries"), leden.Member("1010", "Jelle Bos"), leden.Member("2001", "Bas Anders")]

        def namen(tekst):
            return [x.naam for x in leden.find_members(lijst, tekst)]

        self.assertEqual(namen("1001"), ["Anna de Vries"])  # exact lidnummer
        self.assertEqual(namen("10"), ["Anna de Vries", "Jelle Bos"])  # begin van lidnummer
        self.assertEqual(namen("VRIES anna"), ["Anna de Vries"])  # alle woorden, hoofdletters maken niet uit
        self.assertEqual(namen("an"), ["Anna de Vries", "Bas Anders"])  # gesorteerd op naam
        self.assertEqual(namen("   "), [])


class TestOpslag(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = os.path.join(self.tmp.name, "soos.json")
        self.store = opslag.JsonStore(self.path, opslag.soos_default, opslag.auto_datum)

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

    def test_update_schrijft_ook_een_reservekopie(self):
        self.store.update(lambda data: data["drankjes"].append({"id": 1, "naam": "Bier", "cent": 150}))
        with open(self.path + ".reserve", encoding="utf-8") as f:
            self.assertEqual(json.load(f)["drankjes"], [{"id": 1, "naam": "Bier", "cent": 150}])

    def test_kapot_bestand_wordt_hersteld_uit_reservekopie(self):
        self.store.update(lambda data: data["drankjes"].append({"id": 1, "naam": "Bier", "cent": 150}))
        with open(self.path, "w") as f:
            f.write("")  # leeg, zoals na een stroomstoring
        with self.assertRaises(instellingen.StoreError) as fout:
            self.store.read()
        self.assertIn("hersteld", str(fout.exception))
        # Daarna is alles terug, en het kapotte bestand is bewaard.
        self.assertEqual(self.store.read()["drankjes"], [{"id": 1, "naam": "Bier", "cent": 150}])
        kapot = [n for n in os.listdir(self.tmp.name) if n.startswith("soos.json.kapot-")]
        self.assertEqual(len(kapot), 1)

    def test_herstel_onder_slot_overschrijft_geen_nieuwere_versie(self):
        # Scherm A ziet een kapot bestand. Terwijl A op het slot wacht, herstelt een
        # ander scherm het en slaat er een boeking in op. A mag die nieuwere versie
        # dan niet overschrijven met de (oudere) reservekopie.
        self.store.update(lambda data: data["drankjes"].append({"id": 1, "naam": "Bier", "cent": 150}))
        with open(self.path, "w") as f:
            f.write("")
        nieuwer = dict(self.store.default_factory(), drankjes=[{"id": 1, "naam": "Bier", "cent": 150}, {"id": 2, "naam": "Fris", "cent": 125}])
        echt_enter = opslag.FileLock.__enter__

        def ander_scherm_was_eerst(lock):
            if os.path.getsize(self.path) == 0:
                with open(self.path, "w", encoding="utf-8") as f:
                    json.dump(nieuwer, f)
            return echt_enter(lock)

        with mock.patch.object(opslag.FileLock, "__enter__", ander_scherm_was_eerst):
            data = self.store.read()
        self.assertEqual([d["naam"] for d in data["drankjes"]], ["Bier", "Fris"])
        self.assertEqual([n for n in os.listdir(self.tmp.name) if ".kapot-" in n], [])

    def test_mislukte_reservekopie_geeft_waarschuwing_maar_opslaan_lukt(self):
        meldingen = []
        store = opslag.JsonStore(self.path, opslag.soos_default, opslag.auto_datum, meldingen.append)
        store.update(lambda data: data["drankjes"].append({"id": 1, "naam": "Bier", "cent": 150}))
        self.assertTrue(os.path.exists(self.path + ".reserve"))
        echt_schrijven = opslag.JsonStore._write_file

        def reserve_lukt_niet(path, data):
            if path.endswith(".reserve"):
                raise OSError("schijf vol")
            echt_schrijven(path, data)

        with mock.patch.object(opslag.JsonStore, "_write_file", staticmethod(reserve_lukt_niet)):
            store.update(lambda data: data["drankjes"].append({"id": 2, "naam": "Fris", "cent": 125}))
        self.assertEqual(len(store.read()["drankjes"]), 2)  # de wijziging zelf is opgeslagen
        self.assertEqual(len(meldingen), 1)
        self.assertIn("reservekopie", meldingen[0])
        # De verouderde reservekopie is weg, zodat een herstel niets kan terugdraaien.
        self.assertFalse(os.path.exists(self.path + ".reserve"))

    def test_kapot_bestand_zonder_reservekopie_stopt_het_opslaan(self):
        for inhoud in ("{dit is geen json", "[]", ""):
            with self.subTest(inhoud=inhoud):
                with open(self.path, "w") as f:
                    f.write(inhoud)
                with self.assertRaises(instellingen.StoreError):
                    self.store.read()
                # Een klik mag het kapotte bestand niet stil vervangen door lege gegevens.
                with self.assertRaises(instellingen.StoreError):
                    self.store.update(lambda data: data["drankjes"].append({"id": 1, "naam": "Bier", "cent": 150}))
                with open(self.path) as f:
                    self.assertEqual(f.read(), inhoud)

    def test_opslaan_probeert_opnieuw_als_het_bestand_even_bezet_is(self):
        echt = os.replace
        pogingen = []

        def soms_bezet(bron, doel):
            pogingen.append(doel)
            if len(pogingen) <= 2:
                raise PermissionError("bezet door een ander programma")
            echt(bron, doel)

        with mock.patch.object(opslag.os, "replace", soms_bezet), mock.patch.object(opslag, "REPLACE_WAIT", 0):
            self.store.update(lambda data: data["drankjes"].append({"id": 1, "naam": "Bier", "cent": 150}))
        self.assertEqual(self.store.read()["drankjes"], [{"id": 1, "naam": "Bier", "cent": 150}])

    def test_opslaan_geeft_nette_fout_als_het_bestand_bezet_blijft(self):
        def altijd_bezet(bron, doel):
            raise PermissionError("bezet door een ander programma")

        with mock.patch.object(opslag.os, "replace", altijd_bezet), mock.patch.object(opslag, "REPLACE_WAIT", 0):
            with self.assertRaises(instellingen.StoreError):
                self.store.update(lambda data: data["drankjes"].append({"id": 1, "naam": "Bier", "cent": 150}))

    def test_bezet_slot_geeft_nette_fout(self):
        open(self.path + ".lock", "w").close()  # een vers slot van "een ander scherm"
        with self.assertRaises(instellingen.StoreError):
            with opslag.FileLock(self.path + ".lock", timeout=0.2):
                pass

    def test_rekening(self):
        soos = {
            "drankjes": [{"id": 1, "naam": "Bier", "cent": 150}, {"id": 2, "naam": "Fris", "cent": 100}],
            "verbruik": {"1001": {"naam": "Anna de Vries", "drank": {"1": 3, "2": 1}}},
        }
        self.assertEqual(opslag.member_total_cents(soos, "1001"), 550)
        self.assertEqual(opslag.member_total_cents(soos, "9999"), 0)

    def test_excel_nr(self):
        self.assertEqual(opslag.excel_nr("1001"), 1001)
        self.assertEqual(opslag.excel_nr("007"), "007")
        self.assertEqual(opslag.excel_nr("A12"), "A12")

    def test_nieuwe_soos_begint_met_drankjes_uit_startbestand(self):
        start = os.path.join(self.tmp.name, "soos_standaard.json")
        with open(start, "w", encoding="utf-8") as f:
            json.dump({"drankjes": [{"id": 3, "naam": "Bier", "cent": 150}], "volgende_id": 2, "verbruik": {}}, f)
        store = opslag.JsonStore(self.path, lambda: opslag.soos_default(start), opslag.auto_datum)
        data = store.read()
        self.assertEqual(data["drankjes"], [{"id": 3, "naam": "Bier", "cent": 150}])
        self.assertEqual(data["volgende_id"], 4)  # nooit een bestaand id hergebruiken
        self.assertEqual(data["verbruik"], {})
        # Daarna verandert alleen soos.json, het startbestand niet.
        store.update(lambda d: d["drankjes"].append({"id": 4, "naam": "Fris", "cent": 125}))
        with open(start, encoding="utf-8") as f:
            self.assertEqual(len(json.load(f)["drankjes"]), 1)

    def test_kapot_of_ontbrekend_startbestand_geeft_lege_drankenlijst(self):
        start = os.path.join(self.tmp.name, "soos_standaard.json")
        self.assertEqual(opslag.soos_default(start)["drankjes"], [])
        with open(start, "w", encoding="utf-8") as f:
            f.write("{kapot")
        self.assertEqual(opslag.soos_default(start)["drankjes"], [])
        # Geldige drankjes maar een ongeldige teller: liever leeg dan drankjes met teller 1,
        # want dan zouden nieuwe drankjes bestaande ids hergebruiken.
        with open(start, "w", encoding="utf-8") as f:
            json.dump({"drankjes": [{"id": 5, "naam": "Bier", "cent": 150}], "volgende_id": None}, f)
        data = opslag.soos_default(start)
        self.assertEqual((data["drankjes"], data["volgende_id"]), ([], 1))

    def test_soos_standaard_bevat_alleen_drankjes(self):
        # Het startbestand staat in git: er mogen nooit boekingen (met namen van leden) in.
        with open(instellingen.SOOS_START_PATH, encoding="utf-8") as f:
            start = json.load(f)
        self.assertLessEqual(set(start), {"drankjes", "volgende_id", "verbruik"})
        self.assertFalse(start.get("verbruik"))
        ids = [d["id"] for d in start["drankjes"]]
        self.assertTrue(ids)
        self.assertEqual(len(ids), len(set(ids)))
        self.assertLess(max(ids), start["volgende_id"])
        for d in start["drankjes"]:
            self.assertIsInstance(d["cent"], int)

    def test_export_overschrijft_nooit_een_andere_lijst(self):
        map_ = self.tmp.name

        def naam(datum, eigen=None):
            return os.path.basename(opslag.export_path(map_, "Soos", datum, eigen))

        self.assertEqual(naam("04-11-2026"), "Soos 04-11-2026.xlsx")
        open(os.path.join(map_, "Soos 04-11-2026.xlsx"), "w").close()
        # Dezelfde lijst nog een keer exporteren: eigen bestand vervangen.
        self.assertEqual(naam("04-11-2026", "Soos 04-11-2026.xlsx"), "Soos 04-11-2026.xlsx")
        # Een andere lijst met dezelfde datum: nieuw bestand ernaast.
        self.assertEqual(naam("04-11-2026"), "Soos 04-11-2026 (2).xlsx")
        open(os.path.join(map_, "Soos 04-11-2026 (2).xlsx"), "w").close()
        self.assertEqual(naam("04-11-2026", "Soos 04-11-2026 (2).xlsx"), "Soos 04-11-2026 (2).xlsx")
        self.assertEqual(naam("04-11-2026"), "Soos 04-11-2026 (3).xlsx")
        # Is de datum intussen gewijzigd, dan telt de oude naam niet meer als eigen bestand.
        self.assertEqual(naam("05-11-2026", "Soos 04-11-2026.xlsx"), "Soos 05-11-2026.xlsx")


class TestIcoon(unittest.TestCase):
    def test_visje_is_een_geldig_icoon_met_kleine_maten(self):
        # Het .ico-bestand naast maaltijd.py (zie App.ICON_FILE). Kopje: 0, 1, aantal
        # plaatjes; daarna per plaatje 16 bytes, met als eerste byte de breedte (0 = 256).
        with open(os.path.join(instellingen.APP_DIR, "ichthus-visje.ico"), "rb") as f:
            data = f.read()
        reserved, kind, count = struct.unpack("<HHH", data[:6])
        self.assertEqual((reserved, kind), (0, 1))
        sizes = {data[6 + 16 * i] or 256 for i in range(count)}
        self.assertTrue({16, 32, 48} <= sizes, sizes)


if __name__ == "__main__":
    unittest.main()
