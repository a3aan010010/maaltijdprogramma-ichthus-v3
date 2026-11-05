"""
5. Opslag met slot (FileLock, JsonStore) en het gegevensformaat van maaltijd.json en soos.json.
"""

import contextlib
import json
import os
import re
import shutil
import time

from .hulpjes import today_str
from .instellingen import StoreError


# Hoe vaak en hoe lang we het omwisselen van een bestand opnieuw proberen als
# Windows "Toegang geweigerd" geeft. Dat gebeurt als een ander programma (een
# virusscanner, een back-upprogramma, of dit programma een tweede keer gestart)
# het bestand precies op dat moment open heeft; meestal is dat na een paar
# milliseconden voorbij. 20 x 0,05 seconde: het scherm wacht hooguit een seconde.
REPLACE_TRIES = 20
REPLACE_WAIT = 0.05


def replace_file(source, target):
    """os.replace, maar bij PermissionError een paar keer opnieuw proberen (zie hierboven)."""
    for attempt in range(REPLACE_TRIES):
        try:
            os.replace(source, target)
            return
        except PermissionError:
            if attempt == REPLACE_TRIES - 1:
                raise  # blijft het bezet, dan krijgt de gebruiker een melding
            time.sleep(REPLACE_WAIT)


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
    on_warning      : optionele functie(tekst) voor een waarschuwing die niets
                      tegenhoudt: de wijziging is wel gelukt (zie _reserve_failed)

    Gebruik:
        data = store.read()                 # alleen lezen, geen slot nodig
        store.update(lambda data: ...)      # lezen + aanpassen + opslaan onder slot

    Belangrijk: pas gegevens ALTIJD aan via update(), nooit via read() gevolgd
    door zelf wegschrijven. Alleen update() voorkomt dat twee schermen elkaars
    wijzigingen overschrijven.

    Na elke geslaagde update staat dezelfde inhoud ook in <bestand>.reserve.
    Raakt het echte bestand beschadigd, dan zet _load die reservekopie terug.
    """

    def __init__(self, path, default_factory, upgrade=None, on_warning=None):
        self.path, self.default_factory, self.upgrade = path, default_factory, upgrade
        self.on_warning = on_warning
        self.reserve_path = path + ".reserve"
        self._locked = False  # True zolang deze store het slot heeft (zie _lock)

    @contextlib.contextmanager
    def _lock(self):
        """Pakt het slot en onthoudt dat we het hebben, zodat _load het niet nog eens probeert."""
        with FileLock(self.path + ".lock"):
            self._locked = True
            try:
                yield
            finally:
                self._locked = False

    def read(self):
        """Leest de gegevens. Ontbreekt het bestand, dan komen er lege gegevens terug."""
        data = self._load()
        if self.upgrade:
            self.upgrade(data)  # vult velden aan die in oudere bestanden ontbreken
        return data

    def _load(self):
        """Leest het ruwe JSON-bestand. Lezen zonder slot is veilig: we schrijven
        altijd via os.replace, dus een lezer ziet het oude of het nieuwe bestand,
        nooit een half geschreven versie.
        Een beschadigd bestand geeft altijd een StoreError (zie _repair): we gaan
        nooit stil met lege gegevens verder, want dan is de avond ongemerkt weg."""
        try:
            return self._read_file(self.path)
        except FileNotFoundError:
            return self.default_factory()
        except ValueError:
            pass  # beschadigd: hieronder herstellen
        except OSError as e:
            raise StoreError(f"De gegevens konden niet worden gelezen: {e}")
        if self._locked:
            raise self._repair()  # update() heeft het slot al
        # Herstellen gebeurt altijd onder het slot, en daarbinnen lezen we opnieuw:
        # een ander scherm kan het bestand net hersteld en er alweer iets in
        # opgeslagen hebben. Een oude reservekopie mag dat niet overschrijven.
        with self._lock():
            return self._load()

    @staticmethod
    def _read_file(path):
        """Leest een gegevensbestand. ValueError als het geen geldige gegevens zijn."""
        with open(path, encoding="utf-8") as f:
            data = json.load(f)  # ook verkeerde tekens (UnicodeDecodeError) zijn een ValueError
        if not isinstance(data, dict):
            raise ValueError("geen gegevensbestand")
        return data

    @staticmethod
    def _write_file(path, data):
        """Eerst naar .tmp, dan in een keer omwisselen: nooit een half bestand."""
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=1, ensure_ascii=False)
            # fsync: echt naar de schijf, niet alleen in het geheugen van Windows.
            # Zonder dit kan het bestand na een stroomstoring leeg zijn.
            f.flush()
            os.fsync(f.fileno())
        replace_file(tmp, path)

    def _repair(self):
        """
        Het gegevensbestand is beschadigd (bijvoorbeeld leeg na een stroomstoring,
        of in Kladblok verkeerd opgeslagen). Geeft een StoreError terug met uitleg.

        Is er een bruikbare reservekopie, dan bewaren we het kapotte bestand als
        <bestand>.kapot-<tijd> en zetten we de reservekopie terug. Zo niet, dan
        blijft het kapotte bestand staan en wordt er niets opgeslagen tot iemand
        het oplost. Wordt alleen onder het slot aangeroepen (zie _load).
        """
        name = os.path.basename(self.path)
        try:
            data = self._read_file(self.reserve_path)
            saved_at = time.strftime("%d-%m-%Y %H:%M", time.localtime(os.path.getmtime(self.reserve_path)))
        except (OSError, ValueError):
            return StoreError(
                f"Het gegevensbestand {name} is beschadigd en er is geen bruikbare reservekopie.\n\n"
                "Er wordt niets opgeslagen tot dit is opgelost. Sluit het programma en zet "
                f"{name} terug uit de map 'archief', of verwijder {name} om leeg te beginnen."
            )
        broken = f"{self.path}.kapot-{int(time.time())}"
        try:
            # Kopie in plaats van verplaatsen: zo staat er op elk moment een bestand.
            shutil.copyfile(self.path, broken)
            self._write_file(self.path, data)
        except OSError as e:
            return StoreError(f"Het gegevensbestand {name} is beschadigd en kon niet worden hersteld: {e}")
        return StoreError(
            f"Het gegevensbestand {name} was beschadigd en is hersteld uit de reservekopie van {saved_at}.\n\n"
            f"Het beschadigde bestand is bewaard als {os.path.basename(broken)}. "
            "Controleer de laatste boekingen en doe je laatste klik opnieuw."
        )

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
        with self._lock():
            # Binnen het slot opnieuw lezen: een ander scherm kan net iets hebben
            # opgeslagen. Zo bouwen we altijd voort op de nieuwste versie.
            data = self.read()
            change(data)
            try:
                self._write_file(self.path, data)
            except OSError as e:
                raise StoreError(f"De gegevens konden niet worden opgeslagen: {e}")
            try:
                self._write_file(self.reserve_path, data)
            except OSError as e:
                self._reserve_failed(e)
        return data

    def _reserve_failed(self, error):
        """
        De wijziging staat veilig op schijf, maar de reservekopie kon niet worden
        bijgewerkt. De wijziging geldt als gelukt (anders klikt iemand nog een keer);
        er komt alleen een waarschuwing via on_warning.
        De oude reservekopie halen we weg: een later herstel zou anders al opgeslagen
        wijzigingen stil terugdraaien. Lukt ook dat niet, dan noemt de herstelmelding
        in elk geval de tijd van de reservekopie.
        """
        try:
            os.remove(self.reserve_path)
        except OSError:
            pass
        if self.on_warning:
            name = os.path.basename(self.path)
            self.on_warning(
                f"De wijziging is opgeslagen, maar de reservekopie van {name} kon niet worden bijgewerkt:\n{error}\n\n"
                "Controleer of de opslagmap bereikbaar is en niet vol staat."
            )


# ---------------------------------------------------------------------------
# Gegevensformaat
#
# maaltijd.json:
#   {
#     "datum": "30-10-2026",          datum van de maaltijd (dd-mm-jjjj)
#     "datum_handmatig": false,       false = automatisch (zie auto_datum)
#     "aanmeldingen": {               per lidnummer, in volgorde van aanmelden
#       "1001": {"naam": "Anna de Vries", "aantal": 2, "tijd": "17:45"}
#     },
#     "export_bestand": "Maaltijd 30-10-2026.xlsx"   alleen na een export (zie export_path)
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
#     },
#     "export_bestand": "Soos 30-10-2026.xlsx"
#   }
#
# Naast elk bestand staat <bestand>.reserve met dezelfde inhoud (zie JsonStore).
#
# De naam wordt bij elke aanmelding/boeking meebewaard. Zo kloppen exports ook
# als het ledenbestand intussen wijzigt of iemand eruit is gehaald.
# Drank-ids zijn in "drank" tekst ("1"), omdat JSON alleen tekst als sleutel kent.
# Een drankje met aantal 0 wordt uit "drank" verwijderd, zodat het bestand schoon blijft.
# ---------------------------------------------------------------------------


def meal_default():
    """Lege maaltijdgegevens; zie het gegevensformaat hierboven."""
    # datum_handmatig: False = datum automatisch (zie auto_datum), True = zelf ingesteld
    # aanmeldingen: lidnummer -> {naam, aantal, tijd}
    return {"datum": today_str(), "datum_handmatig": False, "aanmeldingen": {}}


def soos_default(start_path=None):
    """
    Soosgegevens voor als er nog geen soos.json is; zie het gegevensformaat hierboven.

    start_path: een startbestand met drankjes en prijzen (soos_standaard.json naast
    maaltijd.py). Zo begint een nieuwe installatie meteen met de vaste drankjes.
    Daarna wordt alleen soos.json aangepast; het startbestand blijft zoals het is.
    Ontbreekt het startbestand of klopt het niet, dan begint de soos zonder drankjes.
    """
    # datum: de datum van de soos (los van de maaltijd)
    # drankjes: lijst van {id, naam, cent}; verbruik: lidnummer -> {naam, drank: {id: aantal}}
    data = {"datum": today_str(), "datum_handmatig": False, "drankjes": [], "volgende_id": 1, "verbruik": {}}
    if start_path:
        try:
            with open(start_path, encoding="utf-8") as f:
                start = json.load(f)
            drinks = [{"id": int(d["id"]), "naam": str(d["naam"]), "cent": int(d["cent"])} for d in start["drankjes"]]
            # Nooit een id hergebruiken, ook niet als volgende_id in het startbestand te laag is.
            # Eerst alles uitrekenen en pas daarna invullen: gaat er iets mis, dan blijven
            # de lege standaardgegevens heel (en nooit drankjes met teller 1).
            next_id = max([int(start.get("volgende_id", 1))] + [d["id"] + 1 for d in drinks])
            data["drankjes"], data["volgende_id"] = drinks, next_id
        except (OSError, ValueError, KeyError, TypeError):
            pass  # zonder bruikbaar startbestand: lege drankenlijst
    return data


def has_entries(data):
    """True als de lijst al iets bevat: een aanmelding (maaltijd) of een geboekt drankje (soos)."""
    if data.get("aanmeldingen"):
        return True
    return any(e.get("drank") for e in data.get("verbruik", {}).values())


def auto_datum(data):
    """
    Wordt bij elke leesbeurt toegepast. Zonder handmatige datum geldt:
      * lege lijst: de datum is die van vandaag;
      * staat er al iets op: de datum blijft die van de dag waarop de lijst begon.

    De eerste aanmelding of boeking legt die dag vast: update() leest via deze
    functie (datum = vandaag) en schrijft daarna alles weg, de datum ook.
    Zo houdt een soos die na middernacht doorloopt de datum van de avond waarop
    hij begon, en krijgen de export en het archief die datum ook. "Nieuwe soos"
    en "Nieuwe maaltijd" maken de lijst leeg; dan volgt de datum weer vandaag.
    Staat er na een week nog steeds een oude datum, dan is "Nieuwe ..." vergeten.

    Heeft iemand zelf een datum opgeslagen (datum_handmatig = True), dan blijft
    die staan. Oude bestanden zonder het veld datum_handmatig werken ook.
    """
    if data.get("datum_handmatig") and data.get("datum"):
        return
    data["datum_handmatig"] = False
    if not data.get("datum") or not has_entries(data):
        data["datum"] = today_str()


def set_auto_datum():
    """Geeft een wijzigfunctie voor de knop Vandaag: de datum wordt vandaag en weer automatisch."""

    def change(data):
        data["datum"] = today_str()
        data["datum_handmatig"] = False

    return change


def export_path(folder, prefix, datum, own_name=None):
    """
    Pad voor de export "<prefix> <datum>.xlsx" in `folder`, zonder ooit de export
    van een andere lijst te overschrijven.

    own_name is de naam van de vorige export van DEZE lijst (het veld
    export_bestand). Die mag vervangen worden: zo kun je tijdens de avond gerust
    vaker exporteren. Bestaat "<prefix> <datum>.xlsx" al, maar hoort hij bij een
    andere lijst (bijvoorbeeld twee soosavonden met dezelfde datum), dan wordt
    het "<prefix> <datum> (2).xlsx", (3), enzovoort.
    """
    safe = re.sub(r'[<>:"/\\|?*]', "-", datum)  # tekens die Windows niet toestaat in een naam
    base = f"{prefix} {safe}"
    if own_name and re.fullmatch(re.escape(base) + r"( \(\d+\))?\.xlsx", own_name):
        return os.path.join(folder, own_name)
    name, n = f"{base}.xlsx", 1
    while os.path.exists(os.path.join(folder, name)):
        n += 1
        name = f"{base} ({n}).xlsx"
    return os.path.join(folder, name)


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
