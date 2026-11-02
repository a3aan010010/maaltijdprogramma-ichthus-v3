"""
5. Opslag met slot (FileLock, JsonStore) en het gegevensformaat van maaltijd.json en soos.json.
"""

import json
import os
import time

from .hulpjes import today_str
from .instellingen import StoreError


# ---------------------------------------------------------------------------
# 5. Opslag met slot
#    Twee programma's (maaltijd en soos, of twee computers) mogen tegelijk
#    schrijven. Een slotbestand zorgt dat steeds een tegelijk bezig is, en elke
#    wijziging is: inlezen, aanpassen, wegschrijven. Zo gaat er niets verloren.
# ---------------------------------------------------------------------------


class FileLock:
    """
    Een eenvoudig slot via een bestand, te gebruiken als:  with FileLock(pad): ...

    Werking: os.open met O_CREAT | O_EXCL maakt het bestand alleen aan als het
    nog niet bestaat; het besturingssysteem garandeert dat precies een programma
    daarin slaagt. Wie het bestand heeft gemaakt, heeft het slot. Bij het
    verlaten van het with-blok wordt het bestand weer verwijderd.

    Dit werkt ook over een netwerkschijf, zodat twee computers op dezelfde map
    elkaar niet in de weg zitten.

    timeout     : zoveel seconden wachten we maximaal op het slot. Het scherm
                  reageert zolang niet, dus houd dit kort.
    stale_after : een slot dat ouder is dan dit, is van een gecrasht programma
                  en wordt opgeruimd. Een normale schrijfactie duurt milliseconden.
    """

    def __init__(self, path, timeout=3.0, stale_after=10.0):
        self.path, self.timeout, self.stale_after = path, timeout, stale_after

    def __enter__(self):
        start = time.time()
        while True:
            try:
                os.close(os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY))
                return self
            except FileExistsError:
                # Een slot dat te lang blijft staan is van een gecrasht programma.
                try:
                    if time.time() - os.path.getmtime(self.path) > self.stale_after:
                        os.remove(self.path)
                        continue
                except OSError:
                    pass
                if time.time() - start > self.timeout:
                    raise StoreError("De gegevens zijn bezig bij een ander scherm. Probeer het over een paar seconden opnieuw.")
                time.sleep(0.05)  # even wachten en opnieuw proberen
            except OSError as e:
                raise StoreError(f"Kan niet opslaan in de opslagmap: {e}")

    def __exit__(self, *exc):
        try:
            os.remove(self.path)
        except OSError:
            pass


class JsonStore:
    """
    Een JSON-bestand met gegevens, veilig te delen tussen schermen.

    path            : het JSON-bestand (bijv. <opslagmap>/maaltijd.json)
    default_factory : functie die lege gegevens maakt als het bestand nog niet bestaat
    upgrade         : optionele functie die na elke leesbeurt wordt aangeroepen om
                      gegevens aan te vullen of bij te werken (zie auto_datum)

    Gebruik:
        data = store.read()                 # alleen lezen, geen slot nodig
        store.update(lambda data: ...)      # lezen + aanpassen + opslaan onder slot

    Belangrijk: pas gegevens ALTIJD aan via update(), nooit via read() gevolgd
    door zelf wegschrijven. Alleen update() voorkomt dat twee schermen elkaars
    wijzigingen overschrijven.
    """

    def __init__(self, path, default_factory, upgrade=None):
        self.path, self.default_factory, self.upgrade = path, default_factory, upgrade

    def read(self):
        """Leest de gegevens. Ontbreekt het bestand, dan komen er lege gegevens terug."""
        data = self._load()
        if self.upgrade:
            self.upgrade(data)  # vult velden aan die in oudere bestanden ontbreken
        return data

    def _load(self):
        """Leest het ruwe JSON-bestand. Lezen zonder slot is veilig: we schrijven
        altijd via os.replace, dus een lezer ziet het oude of het nieuwe bestand,
        nooit een half geschreven versie."""
        try:
            with open(self.path, encoding="utf-8") as f:
                return json.load(f)
        except FileNotFoundError:
            return self.default_factory()
        except ValueError:
            # Kapot bestand: bewaar het onder een andere naam en begin schoon.
            broken = f"{self.path}.kapot-{int(time.time())}"
            os.replace(self.path, broken)
            return self.default_factory()
        except OSError as e:
            raise StoreError(f"De gegevens konden niet worden gelezen: {e}")

    def update(self, change):
        """
        Leest, past aan met `change(data)`, en schrijft atomair weg. Geeft de nieuwe data terug.

        `change` krijgt de actuele gegevens als dict en past die ter plekke aan
        (de return-waarde wordt genegeerd). Wil je iets terugmelden aan de
        aanroeper, zoals "bestond al", zet het dan in een dict buiten de functie;
        zo gebeurt het ook in de pagina's (zie bijv. DrinksPage.add_drink).
        """
        try:
            os.makedirs(os.path.dirname(self.path), exist_ok=True)
        except OSError as e:
            raise StoreError(f"De opslagmap kon niet worden aangemaakt: {e}")
        with FileLock(self.path + ".lock"):
            # Binnen het slot opnieuw lezen: een ander scherm kan net iets hebben
            # opgeslagen. Zo bouwen we altijd voort op de nieuwste versie.
            data = self.read()
            change(data)
            # Eerst naar .tmp, dan in een keer omwisselen: nooit een half bestand.
            tmp = self.path + ".tmp"
            try:
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=1, ensure_ascii=False)
                    # fsync: echt naar de schijf, niet alleen in het geheugen van Windows.
                    # Zonder dit kan het bestand na een stroomstoring leeg zijn.
                    f.flush()
                    os.fsync(f.fileno())
                os.replace(tmp, self.path)
            except OSError as e:
                raise StoreError(f"De gegevens konden niet worden opgeslagen: {e}")
        return data


# ---------------------------------------------------------------------------
# Gegevensformaat
#
# maaltijd.json:
#   {
#     "datum": "30-10-2026",          datum van de maaltijd (dd-mm-jjjj)
#     "datum_handmatig": false,       false = volgt automatisch vandaag
#     "aanmeldingen": {               per lidnummer, in volgorde van aanmelden
#       "1001": {"naam": "Anna de Vries", "aantal": 2, "tijd": "17:45"}
#     }
#   }
#
# soos.json:
#   {
#     "datum": "30-10-2026",
#     "datum_handmatig": false,
#     "drankjes": [                   volgorde = volgorde op het scherm
#       {"id": 1, "naam": "Bier", "cent": 150}
#     ],
#     "volgende_id": 2,               id voor het volgende nieuwe drankje
#     "verbruik": {                   per lidnummer
#       "1001": {"naam": "Anna de Vries", "drank": {"1": 3}}   drank-id -> aantal
#     }
#   }
#
# De naam wordt bij elke aanmelding/boeking meebewaard. Zo kloppen exports ook
# als het ledenbestand intussen wijzigt of iemand eruit is gehaald.
# Drank-ids zijn in "drank" tekst ("1"), omdat JSON alleen tekst als sleutel kent.
# Een drankje met aantal 0 wordt uit "drank" verwijderd, zodat het bestand schoon blijft.
# ---------------------------------------------------------------------------


def meal_default():
    """Lege maaltijdgegevens; zie het gegevensformaat hierboven."""
    # datum_handmatig: False = datum volgt automatisch vandaag, True = zelf ingesteld
    # aanmeldingen: lidnummer -> {naam, aantal, tijd}
    return {"datum": today_str(), "datum_handmatig": False, "aanmeldingen": {}}


def soos_default():
    """Lege soosgegevens; zie het gegevensformaat hierboven."""
    # datum: de datum van de soos (los van de maaltijd)
    # drankjes: lijst van {id, naam, cent}; verbruik: lidnummer -> {naam, drank: {id: aantal}}
    return {"datum": today_str(), "datum_handmatig": False, "drankjes": [], "volgende_id": 1, "verbruik": {}}


def auto_datum(data):
    """
    Wordt bij elke leesbeurt toegepast: zonder handmatige datum is de datum die van vandaag.

    Zo hoeft niemand na middernacht of de volgende ochtend de datum aan te passen.
    Heeft iemand zelf een datum opgeslagen (datum_handmatig = True), dan blijft
    die staan. Oude bestanden zonder het veld datum_handmatig volgen ook vandaag.
    """
    if not data.get("datum_handmatig") or not data.get("datum"):
        data["datum"] = today_str()
        data["datum_handmatig"] = False


def set_manual_datum(datum):
    """Geeft een wijzigfunctie voor JsonStore.update die een eigen datum vastzet."""

    def change(data):
        data["datum"] = datum
        data["datum_handmatig"] = True

    return change


def excel_nr(nr):
    """
    Lidnummer voor in een export: als getal als dat kan (dan sorteert Excel goed),
    maar "007" of "A12" blijven tekst zodat er niets verloren gaat.
    """
    return int(nr) if nr.isdigit() and not nr.startswith("0") else nr


def member_total_cents(soos, nr):
    """
    Rekening van een lid in centen: per drankje aantal x huidige prijs.
    Let op: dit gebruikt de prijs van NU. Wordt een prijs gewijzigd, dan
    verandert de rekening voor al geboekte drankjes mee (de Drankjes-pagina
    waarschuwt daarvoor).
    """
    # Opzoektabel drank-id (als tekst) -> prijs in centen.
    prices = {str(d["id"]): d["cent"] for d in soos["drankjes"]}
    drinks = soos["verbruik"].get(nr, {}).get("drank", {})
    return sum(prices.get(k, 0) * n for k, n in drinks.items())
