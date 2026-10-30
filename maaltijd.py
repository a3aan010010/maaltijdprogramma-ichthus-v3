"""
===============================================================================
  Maaltijd- en soosprogramma voor Ichthus Utrecht
===============================================================================

  Gemaakt door : Adriaan Abrahamse
  Datum        : 30 oktober 2026
  Versie       : 1.0
  Taal         : Python 3 (alleen standaard libraries)
  Licentie     : vrij te gebruiken en aan te passen binnen de vereniging

  Gebouwd met hulp van Claude Code (Anthropic).
  Zie README.md voor installatie, gebruik en uitleg voor ontwikkelaars.

  Wijzigingen
  -----------
  1.0  (30-10-2026)  Eerste versie: maaltijd, soos, drankjes, totalen,
                     instellingen, Excel-export en archief.

===============================================================================

Alleen standaard Python libraries (tkinter, zipfile, json, ...), dus er hoeft niets
geinstalleerd te worden. Starten met:  python maaltijd.py
Voor een gewoon (niet schermvullend) venster:  python maaltijd.py --venster

Wat doet het?
  * Maaltijd: leden melden zich aan met lidnummer OF naam. Iemand mag vaker
    aangemeld worden (bijvoorbeeld met gasten) en kan weer verwijderd worden.
  * Soos: drankjes eenvoudig toevoegen, per lid met plus en min boeken.
  * Beide kunnen tegelijk draaien, ook op twee computers die dezelfde map delen.
    Elke wijziging wordt direct opgeslagen onder een slot (lock) en de schermen
    verversen zichzelf, dus niemand overschrijft elkaars werk.
  * Export naar Excel in twee losse bestanden: "Maaltijd <datum>.xlsx" (aanmeldlijst)
    en "Soos <datum>.xlsx" (rekeningen). Maaltijd en soos hebben elk hun eigen datum,
    exportknop en "Nieuwe ..." knop en raken elkaar nooit.
  * Datum: standaard automatisch die van vandaag; los aan te passen, en met de knop
    "Vandaag" volgt hij weer automatisch.
  * Instellingen: de opslagmap en het ledenbestand (Excel) zijn instelbaar.

Opbouw van dit bestand:
  1. Instellingen (config)       5. Opslag met slot (JsonStore) en gegevensformaat
  2. Excel lezen en schrijven    6. App (gedeelde logica, export, archief)
  3. Ledenlijst en zoeken        7. Schermen (navbar, maaltijd, soos, drankjes,
  4. Geld en datum hulpjes          totalen, instellingen)

Hoe het in grote lijnen werkt (voor wie hieraan verder bouwt):
  * Er is een venster (App.root). Bovenin staat de Navbar, daaronder steeds een
    "pagina" (een ttk.Frame). App.show_page() wisselt van pagina.
  * Alle gegevens staan in twee JSON-bestanden in de opslagmap: maaltijd.json en
    soos.json. Er wordt NIETS in het geheugen bewaard dat niet ook op schijf staat.
    Elke klik doet: slot pakken, bestand lezen, aanpassen, wegschrijven, slot los
    (zie JsonStore.update). Daardoor kunnen meerdere schermen of computers veilig
    tegelijk werken.
  * Elke pagina leest elke 2 seconden het bestand opnieuw (BasePage._poll) en
    tekent zich alleen opnieuw als er echt iets veranderd is.
  * Fouten die de gebruiker moet zien, zijn StoreError. De decorator @guarded
    vangt die op en toont een venstertje in plaats van te crashen.
  * Geld wordt altijd in hele centen (int) opgeslagen, nooit als kommagetal,
    zodat er geen afrondfouten ontstaan.

Een nieuwe pagina toevoegen:
  1. Maak een klasse die erft van BasePage en implementeer refresh().
  2. Roep aan het eind van __init__ self.refresh() en self.start_polling() aan.
  3. Zet hem in App.build() in self.pages en maak een knop of menu-item in Navbar
     (en zet hem in Navbar.GROUP, zodat de juiste knop blauw kleurt).
"""

import datetime
import functools
import json
import os
import re
import shutil
import sys
import time
import tkinter as tk
import zipfile
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from tkinter import filedialog, messagebox, ttk
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

# ---------------------------------------------------------------------------
# 1. Instellingen
# ---------------------------------------------------------------------------

# De map waar dit script staat. De instellingen staan altijd hier, zodat het
# programma ze terugvindt, ook als de opslagmap zelf is verplaatst.
APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "maaltijd_config.json")

# Standaardinstellingen. Alles wat in maaltijd_config.json staat, overschrijft
# deze waarden; ontbreekt een sleutel in dat bestand, dan geldt de waarde hier.
# Een nieuwe instelling toevoegen? Zet hem hier met een standaardwaarde, dan
# werken oude config-bestanden gewoon door.
DEFAULT_CONFIG = {
    # Map waar alle gegevens, exports en het archief komen te staan.
    "data_dir": os.path.join(APP_DIR, "maaltijd_data"),
    # Excel met de leden. Leeg betekent: "leden.xlsx" in de opslagmap.
    "members_file": "",
    # True: het programma start schermvullend. Zet op false in maaltijd_config.json
    # (of start met "python maaltijd.py --venster") voor een gewoon venster.
    "start_maximized": True,
}


class StoreError(Exception):
    """
    Fout met een melding die we rechtstreeks aan de gebruiker mogen tonen.

    Gebruik dit overal waar iets misgaat dat de gebruiker zelf kan oplossen
    (bestand open in Excel, map niet bereikbaar, verkeerd ledenbestand, ...).
    Schrijf de tekst dus in gewoon Nederlands, zonder technische termen.
    Methodes met @guarded tonen deze fout automatisch in een venstertje.
    """


def load_config():
    """Leest de instellingen; ontbrekende of kapotte instellingen worden standaardwaarden."""
    config = dict(DEFAULT_CONFIG)  # kopie, zodat we de standaardwaarden niet aanpassen
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            config.update(json.load(f))
    except (OSError, ValueError):
        # Geen of een kapot bestand: gewoon met de standaardwaarden verder.
        # Bij de volgende keer "Opslaan" in Instellingen wordt het opnieuw geschreven.
        pass
    return config


def save_config(config):
    """Schrijft de instellingen naar maaltijd_config.json naast dit script."""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
    except OSError as e:
        raise StoreError(f"De instellingen konden niet worden opgeslagen: {e}")


# ---------------------------------------------------------------------------
# 2. Excel (xlsx) lezen en schrijven zonder externe libraries
#    Een xlsx is een zip met XML-bestanden; voor onze tabellen is weinig nodig.
#
#    Belangrijkste bestanden in zo'n zip:
#      [Content_Types].xml          welk soort bestand elk onderdeel is
#      _rels/.rels                  verwijst naar de werkmap
#      xl/workbook.xml              lijst met werkbladen (namen)
#      xl/_rels/workbook.xml.rels   koppelt bladnamen aan bladbestanden
#      xl/styles.xml                opmaak (vet, valuta)
#      xl/worksheets/sheet1.xml     de cellen van blad 1
#      xl/sharedStrings.xml         (alleen bij Excel zelf) gedeelde teksten
#
#    Wil je ooit meer opmaak (kleuren, randen)? Dan is het makkelijker om de
#    library openpyxl te gebruiken; dit zelfgebouwde stuk is bewust minimaal.
# ---------------------------------------------------------------------------

# De XML-namespace van Excel. ElementTree zet die voor elke tagnaam, dus een
# cel heet intern "{http://...}c". Met _M + "c" schrijven we dat kort.
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_M = "{%s}" % MAIN_NS

# Stijlnummers voor het schrijven. Ze verwijzen naar de volgorde van de <xf>
# regels in <cellXfs> in STYLES_XML hieronder; pas die twee dus altijd samen aan.
STYLE_NORMAL, STYLE_BOLD, STYLE_MONEY, STYLE_BOLD_MONEY = 0, 1, 2, 3


def _col_index(ref):
    """'C5' -> 2 (kolomnummer vanaf 0). Werkt ook voor 'AA1' -> 26."""
    letters = re.match(r"[A-Z]+", ref).group()
    n = 0
    for ch in letters:
        # Kolomletters zijn een soort 26-tallig getal waarbij A=1 en Z=26.
        n = n * 26 + ord(ch) - 64
    return n - 1


def _col_letter(index):
    """2 -> 'C' (kolomnummer vanaf 0). Het omgekeerde van _col_index."""
    letters = ""
    index += 1
    while index:
        index, rest = divmod(index - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def read_xlsx_rows(path):
    """
    Geeft het eerste werkblad als lijst van rijen (lijsten van tekst).

    Alles komt terug als tekst (ook getallen), met spaties eromheen weggehaald.
    Lege cellen tussen gevulde cellen worden "". Helemaal lege rijen vallen weg.
    Formules worden niet uitgerekend: we lezen de laatst opgeslagen waarde.
    """
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()

            # Gedeelde tekstjes: cellen verwijzen hier naar met een nummer.
            shared = []
            if "xl/sharedStrings.xml" in names:
                root = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for si in root.findall(_M + "si"):
                    shared.append("".join(t.text or "" for t in si.iter(_M + "t")))

            # We nemen het eerste werkblad. In vrijwel elk bestand is dat sheet1.xml.
            sheets = sorted(n for n in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n))
            if not sheets:
                raise StoreError("Dit Excel bestand bevat geen werkblad.")
            sheet = "xl/worksheets/sheet1.xml" if "xl/worksheets/sheet1.xml" in sheets else sheets[0]
            root = ET.fromstring(z.read(sheet))
    except zipfile.BadZipFile:
        raise StoreError("Dit is geen geldig .xlsx bestand (opslaan als 'Excel-werkmap').")
    except OSError as e:
        raise StoreError(f"Het Excel bestand kon niet worden gelezen: {e}")

    rows = []
    for row in root.iter(_M + "row"):
        cells = {}  # kolomnummer -> tekst; lege cellen staan niet in de XML
        for c in row.findall(_M + "c"):
            # Het attribuut t zegt wat voor cel het is:
            #   "s"         = gedeelde tekst: <v> is een nummer in `shared`
            #   "inlineStr" = tekst staat direct in de cel (zo schrijven wij zelf)
            #   "str"/geen  = formule-uitkomst of getal: <v> is de waarde zelf
            kind = c.get("t")
            if kind == "inlineStr":
                value = "".join(t.text or "" for t in c.iter(_M + "t"))
            else:
                v = c.find(_M + "v")
                value = v.text if v is not None and v.text else ""
                if kind == "s" and value:
                    value = shared[int(value)]
            # Een getal als 123 staat soms als "123.0"; maak er weer "123" van.
            if re.fullmatch(r"\d+\.0+", value):
                value = value.split(".")[0]
            cells[_col_index(c.get("r"))] = value.strip()
        if cells:
            # Van {0: "a", 2: "c"} maken we ["a", "", "c"], zodat kolomnummers kloppen.
            rows.append([cells.get(i, "") for i in range(max(cells) + 1)])
    return rows


# Tekens die niet in XML mogen staan (stuurcodes). We halen ze stilletjes weg
# voordat we tekst wegschrijven, anders weigert Excel het bestand te openen.
_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# Minimale opmaak voor Excel. De vier <xf> regels in <cellXfs> zijn, op volgorde:
#   0 = normaal, 1 = vet, 2 = getal met 2 decimalen, 3 = vet getal met 2 decimalen
# numFmtId 4 is een ingebouwd Excel-formaat "#,##0.00"; Excel toont het in de
# taal van de gebruiker, dus in Nederland als 1.234,50.
# fills en borders zijn verplicht van Excel, ook al gebruiken we ze niet.
STYLES_XML = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="{MAIN_NS}">'
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="4">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="4" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>'
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>"
)


def _cell_xml(ref, cell):
    """
    Maakt de XML voor een cel. `ref` is de celnaam (bijv. "B3").
    Een cel is een waarde, of een tuple (waarde, stijlnummer) voor opmaak:
        "Jan"                   -> tekst
        12                      -> getal
        ("Totaal", STYLE_BOLD)  -> vetgedrukte tekst
    """
    value, style = cell if isinstance(cell, tuple) else (cell, STYLE_NORMAL)
    if value is None or value == "":
        return ""  # lege cellen laten we gewoon weg
    s = f' s="{style}"' if style else ""
    if isinstance(value, (int, float)):
        # Getallen als echt getal, zodat je er in Excel mee kunt rekenen.
        return f'<c r="{ref}"{s}><v>{value}</v></c>'
    # Tekst schrijven we "inline": dan is geen apart sharedStrings.xml nodig.
    text = escape(_ILLEGAL_XML.sub("", str(value)))
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def _sheet_xml(rows):
    """Maakt de XML voor een heel werkblad uit een lijst van rijen."""
    # Kolombreedte op basis van de langste tekst in de kolom (minimaal 8, maximaal 50).
    widths = {}
    for row in rows:
        for i, cell in enumerate(row):
            value = cell[0] if isinstance(cell, tuple) else cell
            widths[i] = max(widths.get(i, 8), min(len(str(value)) + 2, 50))
    cols = "".join(
        f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>' for i, w in sorted(widths.items())
    )
    body = ""
    for r, row in enumerate(rows, start=1):
        cells = "".join(_cell_xml(f"{_col_letter(c)}{r}", cell) for c, cell in enumerate(row))
        body += f'<row r="{r}">{cells}</row>'
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{MAIN_NS}">'
        f"<cols>{cols}</cols><sheetData>{body}</sheetData></worksheet>"
    )


def write_xlsx(path, sheets):
    """
    Schrijft een werkmap. `sheets` is een lijst van (bladnaam, rijen).
    Er wordt eerst naar een tijdelijk bestand geschreven, zodat een crash of een
    geopend Excel bestand nooit een half bestand achterlaat.

    Voorbeeld:
        write_xlsx("test.xlsx", [("Blad1", [["Naam", "Aantal"], ["Jan", 2]])])
    """
    # Vaste namespaces en bestandstypes uit de Office Open XML standaard.
    # Deze teksten moeten exact zo blijven, anders opent Excel het bestand niet.
    ct = "http://schemas.openxmlformats.org/package/2006/content-types"
    rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    doc = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ssml = "application/vnd.openxmlformats-officedocument.spreadsheetml"
    head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="{ssml}.worksheet+xml"/>'
        for i in range(1, len(sheets) + 1)
    )
    content_types = (
        f'{head}<Types xmlns="{ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        f'<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/xl/workbook.xml" ContentType="{ssml}.sheet.main+xml"/>'
        f'<Override PartName="/xl/styles.xml" ContentType="{ssml}.styles+xml"/>{overrides}</Types>'
    )
    root_rels = (
        f'{head}<Relationships xmlns="{rel}"><Relationship Id="rId1" '
        f'Type="{doc}/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    )
    def sheet_name(name):
        # Bladnamen mogen maximaal 31 tekens zijn en geen []:*?/\ bevatten.
        return escape(re.sub(r"[\[\]:*?/\\]", "", name)[:31])

    sheet_tags = "".join(
        f'<sheet name="{sheet_name(name)}" sheetId="{i}" r:id="rId{i}"/>'
        for i, (name, _) in enumerate(sheets, start=1)
    )
    workbook = (
        f'{head}<workbook xmlns="{MAIN_NS}" xmlns:r="{doc}"><sheets>{sheet_tags}</sheets></workbook>'
    )
    wb_rels = "".join(
        f'<Relationship Id="rId{i}" Type="{doc}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, len(sheets) + 1)
    ) + f'<Relationship Id="rId{len(sheets) + 1}" Type="{doc}/styles" Target="styles.xml"/>'
    wb_rels = f'{head}<Relationships xmlns="{rel}">{wb_rels}</Relationships>'

    # Eerst alles in een .tmp bestand, daarna in een keer omwisselen (os.replace).
    # Staat het doelbestand open in Excel, dan faalt alleen die laatste stap met
    # PermissionError en blijft het oude bestand heel.
    tmp = path + ".tmp"
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", content_types)
            z.writestr("_rels/.rels", root_rels)
            z.writestr("xl/workbook.xml", workbook)
            z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
            z.writestr("xl/styles.xml", STYLES_XML)
            for i, (_, rows) in enumerate(sheets, start=1):
                z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(rows))
        os.replace(tmp, path)
    except PermissionError:
        raise StoreError(
            f"Het bestand '{os.path.basename(path)}' is geopend in Excel.\n"
            "Sluit het bestand en probeer het opnieuw."
        )
    except OSError as e:
        raise StoreError(f"Het Excel bestand kon niet worden opgeslagen: {e}")
    finally:
        # Na succes bestaat tmp niet meer; na een fout ruimen we het hier op.
        if os.path.exists(tmp):
            os.remove(tmp)


# ---------------------------------------------------------------------------
# 3. Ledenlijst en zoeken
# ---------------------------------------------------------------------------


class Member:
    """
    Een lid: lidnummer en volledige naam.

    Het lidnummer is altijd tekst, ook als het er als getal uitziet. Zo blijven
    voorloopnullen ("007") en lidnummers met letters ("A12") gewoon werken.
    Het lidnummer is de sleutel waarmee alles in maaltijd.json en soos.json
    wordt opgeslagen.
    """

    def __init__(self, nr, naam):
        self.nr = nr
        self.naam = naam

    def label(self):
        """Tekst zoals die in de zoeklijst verschijnt."""
        return f"{self.nr}  {self.naam}"


def load_members(path):
    """
    Leest de leden uit Excel. De eerste rij mag kopjes hebben; we herkennen kolommen
    als 'lidnummer'/'nummer' en 'naam' (of 'voornaam' + 'achternaam'). Zonder kopjes
    geldt: kolom A is het lidnummer en kolom B de naam.
    """
    if not os.path.isfile(path):
        raise StoreError(f"Het ledenbestand is niet gevonden:\n{path}\n\nStel het in via Instellingen.")
    rows = read_xlsx_rows(path)
    if not rows:
        raise StoreError("Het ledenbestand is leeg.")

    # Stap 1: kijk of de eerste rij kopjes bevat en zo ja, welke kolom wat is.
    # De volgorde van de elif's is belangrijk: "voornaam" bevat ook "naam",
    # dus die moeten we eerder herkennen dan de algemene "naam".
    nr_col = name_col = first_col = last_col = None
    header = [h.lower() for h in rows[0]]
    for i, h in enumerate(header):
        if nr_col is None and ("lid" in h or "nummer" in h or h in ("nr", "no", "id")):
            nr_col = i
        elif "voornaam" in h:
            first_col = i
        elif "achternaam" in h:
            last_col = i
        elif "naam" in h and name_col is None:
            name_col = i

    # Stap 2: alleen als we zowel een nummer- als een naamkolom vonden, is rij 1
    # een kopjesrij. Anders gaan we uit van: kolom A = lidnummer, kolom B = naam.
    has_header = nr_col is not None and (name_col is not None or first_col is not None)
    if has_header:
        rows = rows[1:]
    else:
        nr_col, name_col, first_col, last_col = 0, 1, None, None

    # Stap 3: elke rij omzetten naar een Member.
    members, seen = [], set()
    for row in rows:
        def get(i):
            # Veilig een cel ophalen: korte rijen of een ontbrekende kolom geven "".
            return row[i].strip() if i is not None and i < len(row) else ""

        nr = get(nr_col)
        naam = get(name_col) if name_col is not None else " ".join(x for x in (get(first_col), get(last_col)) if x)
        if not nr or not naam or nr in seen:
            continue  # lege regels en dubbele lidnummers slaan we over
        seen.add(nr)
        members.append(Member(nr, naam))
    if not members:
        raise StoreError("Er zijn geen leden gevonden. Controleer de kolommen in het Excel bestand.")
    return members


def find_members(members, text):
    """
    Zoekt leden op lidnummer of naam. Een exact lidnummer wint altijd. Anders
    moeten alle getypte woorden in de naam voorkomen (hoofdletters maken niet uit),
    of het lidnummer moet beginnen met de tekst.

    Voorbeelden (met leden 1001 Anna de Vries en 1010 Jelle Bos):
        "1001"      -> alleen Anna (exact lidnummer)
        "100"       -> Anna en iedereen wiens lidnummer met 100 begint
        "anna"      -> Anna
        "de vries"  -> Anna (beide woorden komen voor)
    """
    t = text.strip().lower()
    if not t:
        return []
    # Een exact lidnummer geeft precies een treffer, zodat Enter direct werkt.
    exact = [m for m in members if m.nr.lower() == t]
    if exact:
        return exact
    words = t.split()
    found = [m for m in members if all(w in m.naam.lower() for w in words) or m.nr.lower().startswith(t)]
    return sorted(found, key=lambda m: m.naam.lower())


# ---------------------------------------------------------------------------
# 4. Geld en datum hulpjes
# ---------------------------------------------------------------------------


# Datums bewaren en tonen we overal als tekst "dd-mm-jjjj" (bijv. 30-10-2026).
# Dat leest prettig voor Nederlanders en mag in een Windows-bestandsnaam staan
# (een "/" mag daar niet in).


def today_str():
    """De datum van vandaag als 'dd-mm-jjjj'."""
    return datetime.date.today().strftime("%d-%m-%Y")


def parse_datum(text):
    """
    Controleert een datum als dd-mm-jjjj. Geeft de nette tekst terug of ValueError.
    Ook "5/1/2026" mag: dat wordt "05-01-2026". Een niet bestaande datum zoals
    31-02-2026 geeft ook een ValueError.
    """
    return datetime.datetime.strptime(text.strip().replace("/", "-"), "%d-%m-%Y").strftime("%d-%m-%Y")


# Hoogste prijs (in euro) die we accepteren. Een drankje van meer dan dit is vrijwel
# zeker een typefout (bijv. 150 in plaats van 1,50). Hoger nodig? Pas dit getal aan.
MAX_PRIJS = 10000


def parse_prijs(text):
    """
    '1,50' of '1.50' -> 150 (centen). Geld rekenen we in centen.

    Gooit ValueError bij onzin. De tekst van die fout is een nette melding in
    het Nederlands die direct aan de gebruiker getoond mag worden, dus:
        except ValueError as e:
            self.message(str(e), error=True)

    We gebruiken Decimal in plaats van float, omdat float niet exact rekent
    (0.1 + 0.2 is daar 0.30000000000000004). Negatieve prijzen en prijzen boven
    MAX_PRIJS worden geweigerd, als vangnet tegen typefouten.
    """
    # Euroteken en spaties weg, komma wordt punt: "€ 1,50" -> "1.50".
    t = text.strip().replace("€", "").replace(" ", "").replace(",", ".")
    try:
        d = Decimal(t)
    except InvalidOperation:
        raise ValueError("Prijs niet herkend. Typ bijvoorbeeld 1,50.")
    # is_finite eerst: "nan" en "inf" zijn geldige Decimals, maar geen prijs.
    if not d.is_finite():
        raise ValueError("Prijs niet herkend. Typ bijvoorbeeld 1,50.")
    if d < 0:
        raise ValueError("Een prijs kan niet negatief zijn.")
    if d > MAX_PRIJS:
        # :, zet een komma als duizendtalscheiding; in het Nederlands is dat een punt.
        maximum = f"{MAX_PRIJS:,}".replace(",", ".")
        raise ValueError(f"Die prijs is te hoog (maximaal {maximum} euro). Controleer de komma, bijvoorbeeld 1,50.")
    # Naar hele centen, netjes afgerond (1,505 -> 151).
    return int((d * 100).quantize(Decimal(1), rounding=ROUND_HALF_UP))


def euro(cents):
    """150 -> '1,50' en 123456 -> '1.234,56' (Nederlandse notatie, zonder euroteken)."""
    euros, rest = divmod(abs(cents), 100)
    # Python zet duizendtallen met een komma; die wordt hier een punt.
    return f"{'-' if cents < 0 else ''}{euros:,}".replace(",", ".") + f",{rest:02d}"


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

    timeout     : zoveel seconden wachten we maximaal op het slot.
    stale_after : een slot dat ouder is dan dit, is van een gecrasht programma
                  en wordt opgeruimd. Een normale schrijfactie duurt milliseconden.
    """

    def __init__(self, path, timeout=10.0, stale_after=30.0):
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


# ---------------------------------------------------------------------------
# 6. App: gedeelde logica voor alle schermen
# ---------------------------------------------------------------------------


class App:
    """
    Het hart van het programma: het hoofdvenster, de instellingen, de ledenlijst,
    en alle logica die niet bij een enkele pagina hoort (export, archief).

    Elke pagina krijgt een verwijzing naar de App (self.app) en haalt daar de
    gegevens vandaan. Volgorde bij het starten (zie main()):
        App()  ->  setup_style()  ->  App.build()  ->  mainloop()
    """

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Maaltijd en soos")
        self.root.minsize(900, 600)
        # Elke fout in een knop of timer komt hier terecht in plaats van stil in de console.
        self.root.report_callback_exception = self._unexpected_error
        self.config = load_config()
        self._apply_window_size()
        self.members = []  # lijst van Member, uit het ledenbestand
        self.members_error = ""  # leeg = ledenlijst in orde, anders de foutmelding
        self.navbar = None
        self.pages = {}  # naam -> pagina; alle pagina's leven in hetzelfde venster
        self.reload_members()

    def _apply_window_size(self):
        """
        Schermvullend (gemaximaliseerd, met titelbalk) tenzij anders ingesteld.
        De pagina's rekken mee met elke venstergrootte, ook als je later verkleint.
        """
        if self.config.get("start_maximized", True) and "--venster" not in sys.argv:
            try:
                self.root.state("zoomed")  # Windows
            except tk.TclError:
                try:
                    self.root.attributes("-zoomed", True)  # Linux
                except tk.TclError:
                    w, h = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
                    self.root.geometry(f"{w}x{h}+0+0")  # vangnet voor andere systemen
        else:
            self.root.geometry("1150x740")

    def build(self):
        """Bouwt de navbar en alle pagina's, en toont de eerste pagina."""
        self.navbar = Navbar(self)
        self.navbar.pack(side="top", fill="x")
        self.pages = {
            "maaltijd": MaaltijdPage(self, self.root),
            "soos": SoosPage(self, self.root),
            "drankjes": DrinksPage(self, self.root),
            "totalen": TotalsPage(self, self.root),
            "instellingen": SettingsPage(self, self.root),
        }
        # Zonder bruikbare ledenlijst starten we bij de instellingen.
        self.show_page("instellingen" if self.members_error else "maaltijd")

    # -- onverwachte fouten: nooit stil crashen, wel een duidelijke melding --
    def _unexpected_error(self, exc, val, tb):
        """
        Vangnet voor fouten die nergens anders zijn opgevangen. Omdat alles direct
        op schijf staat, gaat er bij zo'n fout geen werk verloren. Wil je bij het
        ontwikkelen de volledige foutmelding zien, zet hier tijdelijk:
            import traceback; traceback.print_exception(exc, val, tb)
        """
        if isinstance(val, StoreError):
            messagebox.showerror("Let op", str(val))
        else:
            messagebox.showerror("Onverwachte fout", f"Er ging iets mis: {val}\n\nDe gegevens zijn veilig, probeer het nog eens.")

    # -- instellingen en bestanden --
    @property
    def data_dir(self):
        """De opslagmap uit de instellingen."""
        return self.config["data_dir"]

    @property
    def members_path(self):
        """Het ledenbestand; standaard leden.xlsx in de opslagmap."""
        return self.config["members_file"] or os.path.join(self.data_dir, "leden.xlsx")

    # De stores worden bij elke aanroep opnieuw gemaakt (dat kost niets). Zo
    # wijzen ze na het wijzigen van de opslagmap meteen naar de nieuwe map.
    def meal_store(self):
        """Opslag van de maaltijd (maaltijd.json)."""
        return JsonStore(os.path.join(self.data_dir, "maaltijd.json"), meal_default, auto_datum)

    def soos_store(self):
        """Opslag van de soos en de drankjes (soos.json)."""
        return JsonStore(os.path.join(self.data_dir, "soos.json"), soos_default, auto_datum)

    def reload_members(self):
        """Leest het ledenbestand opnieuw in. Bij een fout blijven de oude leden staan."""
        try:
            self.members = load_members(self.members_path)
            self.members_error = ""
        except StoreError as e:
            self.members_error = str(e)
        return not self.members_error

    # -- export --
    # De exports bouwen een lijst van rijen op (zie write_xlsx) en schrijven die
    # weg in de opslagmap. Een bestaand bestand met dezelfde naam wordt
    # overschreven: zo kun je tijdens de avond gerust vaker exporteren.

    def _export_path(self, prefix, datum):
        """Pad van een exportbestand, zonder tekens die Windows niet toestaat in de naam."""
        os.makedirs(self.data_dir, exist_ok=True)
        safe = re.sub(r'[<>:"/\\|?*]', "-", datum)
        return os.path.join(self.data_dir, f"{prefix} {safe}.xlsx")

    def export_meal(self):
        """Schrijft 'Maaltijd <datum>.xlsx' met de aanmeldlijst. Geeft het pad terug."""
        meal = self.meal_store().read()
        datum = meal["datum"]
        # Indeling: rij 1 titel, rij 2 leeg, rij 3 kopjes, dan een rij per lid,
        # een lege rij en tot slot de totaalregel.
        rows = [[(f"Maaltijd {datum}", STYLE_BOLD)], [], [(h, STYLE_BOLD) for h in ("Lidnummer", "Naam", "Aantal", "Aangemeld om")]]
        total = 0
        for nr, e in meal["aanmeldingen"].items():
            # Lidnummers als getal als dat kan (dan sorteert Excel goed), maar
            # "007" of "A12" blijven tekst zodat er niets verloren gaat.
            rows.append([int(nr) if nr.isdigit() and not nr.startswith("0") else nr, e["naam"], e["aantal"], e.get("tijd", "")])
            total += e["aantal"]
        rows += [[], [("Totaal", STYLE_BOLD), "", (total, STYLE_BOLD)]]
        path = self._export_path("Maaltijd", datum)
        write_xlsx(path, [("Maaltijd", rows)])
        return path

    def export_soos(self):
        """Schrijft 'Soos <datum>.xlsx' met de rekeningen, een kolom per drankje."""
        soos = self.soos_store().read()
        datum = soos["datum"]
        drinks = soos["drankjes"]
        # Kopjes: Lidnummer, Naam, dan een kolom per drankje (met prijs), dan Totaal.
        head = ["Lidnummer", "Naam"] + [f"{d['naam']} ({euro(d['cent'])})" for d in drinks] + ["Totaal (euro)"]
        rows = [[(f"Soos {datum}", STYLE_BOLD)], [], [(h, STYLE_BOLD) for h in head]]
        column_sums = [0] * len(drinks)  # totaal aantal per drankje, voor de onderste regel
        grand = 0  # totaalbedrag van de hele soos, in centen
        for nr, e in soos["verbruik"].items():
            counts = [e["drank"].get(str(d["id"]), 0) for d in drinks]
            if not any(counts):
                continue  # leden die alles terug naar 0 hebben geboekt, slaan we over
            cents = member_total_cents(soos, nr)
            grand += cents
            column_sums = [a + b for a, b in zip(column_sums, counts)]
            rows.append(
                [int(nr) if nr.isdigit() and not nr.startswith("0") else nr, e["naam"]]
                + [c or "" for c in counts]  # 0 tonen we als lege cel, dat leest rustiger
                + [(cents / 100, STYLE_MONEY)]  # Excel wil euro's, wij rekenen in centen
            )
        rows += [[], [("Totaal", STYLE_BOLD), ""] + [(c, STYLE_BOLD) for c in column_sums] + [(grand / 100, STYLE_BOLD_MONEY)]]
        path = self._export_path("Soos", datum)
        write_xlsx(path, [("Soos", rows)])
        return path

    def _archive(self, store, prefix, datum):
        """
        Kopieert een gegevensbestand naar de map 'archief' en geeft de archiefmap terug.
        De tijd (uur, minuut, seconde) in de mapnaam zorgt dat twee archieven op
        dezelfde dag elkaar niet overschrijven. Zo'n JSON-kopie kun je later altijd
        terugzetten door hem terug te kopieren naar de opslagmap.
        """
        stamp = datetime.datetime.now().strftime("%H%M%S")
        folder = os.path.join(self.data_dir, "archief", f"{prefix} {datum}_{stamp}")
        os.makedirs(folder, exist_ok=True)
        if os.path.exists(store.path):
            shutil.copy2(store.path, folder)
        return folder

    def archive_and_reset_meal(self):
        """Bewaart de aanmeldlijst in het archief en begint een lege maaltijd. De soos blijft ongemoeid."""
        store = self.meal_store()
        folder = self._archive(store, "Maaltijd", store.read()["datum"])
        store.update(lambda data: data.update(meal_default()))
        return folder

    def archive_and_reset_soos(self):
        """Bewaart de soos in het archief en begint leeg. De drankjes en de maaltijd blijven ongemoeid."""
        store = self.soos_store()
        folder = self._archive(store, "Soos", store.read()["datum"])

        def reset(data):
            data["verbruik"] = {}  # de drankjes zelf blijven bestaan
            data["datum"] = today_str()
            data["datum_handmatig"] = False

        store.update(reset)
        return folder

    # -- pagina's --
    def show_page(self, name):
        """Toont een pagina en verbergt de andere. Zonder ledenlijst kan alleen Instellingen."""
        # Maaltijd en Soos hebben de ledenlijst nodig om iemand te kunnen kiezen.
        if name in ("maaltijd", "soos") and self.members_error:
            messagebox.showerror("Ledenlijst", self.members_error)
            name = "instellingen"
        # Alle pagina's bestaan altijd; we laten er steeds een zien met pack().
        # Verborgen pagina's blijven op de achtergrond verversen, dat is goedkoop.
        for page in self.pages.values():
            page.pack_forget()
        self.pages[name].pack(fill="both", expand=True)
        self.pages[name].on_show()
        self.navbar.highlight(name)


def guarded(method):
    """
    Toont StoreError-meldingen als vriendelijk venstertje in plaats van te crashen.

    Zet @guarded boven elke methode die door een knop wordt aangeroepen en iets
    leest of opslaat. Binnen zo'n methode mag je dus gewoon `raise StoreError(...)`
    doen; de gebruiker krijgt dan de melding en het programma loopt door.
    Werkt op elke klasse die een tkinter-widget is (nodig voor parent=self).
    """

    @functools.wraps(method)
    def wrapper(self, *args, **kwargs):
        try:
            return method(self, *args, **kwargs)
        except StoreError as e:
            messagebox.showerror("Let op", str(e), parent=self)

    return wrapper


# ---------------------------------------------------------------------------
# 7. Schermen
#
#    Bouwstenen:
#      setup_style  lettergroottes en knopstijlen voor het hele programma
#      MemberPicker zoekveld + lijst om een lid te kiezen (Maaltijd en Soos)
#      BasePage     basis voor elke pagina: statusregel en automatisch verversen
#      DatumField   datumveld met "Datum opslaan" en "Vandaag"
#
#    Pagina's (in App.pages):
#      "maaltijd"     MaaltijdPage   aanmelden voor de maaltijd
#      "soos"         SoosPage       drankjes boeken per lid (knop Soos)
#      "drankjes"     DrinksPage     drankjes en prijzen beheren (Instellingen > Drankjes en prijzen)
#      "totalen"      TotalsPage     overzicht van de avond (Instellingen > Totalen van de avond)
#      "instellingen" SettingsPage   opslagmap en ledenbestand (Instellingen > Opslag en ledenlijst)
#
#    Navbar: de balk bovenin die tussen de pagina's wisselt.
#
#    Lay-out: tkinter kent pack (stapelen) en grid (tabel). Binnen een frame
#    gebruik je er maar een van; mengen in hetzelfde frame laat het programma
#    vastlopen.
# ---------------------------------------------------------------------------


def setup_style(root):
    """
    Grotere letters, zodat het ook aan de bar goed leesbaar is.
    "Big.TButton" is een eigen knopstijl voor de belangrijkste knoppen; gebruik
    hem met ttk.Button(..., style="Big.TButton").
    """
    import tkinter.font as tkfont

    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
        tkfont.nametofont(name).configure(size=12)
    style = ttk.Style(root)
    style.configure("Treeview", rowheight=30, font=("Segoe UI", 12))
    style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
    style.configure("Big.TButton", font=("Segoe UI", 13, "bold"), padding=8)
    # Voor de + en - knoppen in een drankrij: ruim genoeg om te raken, maar laag,
    # zodat er veel rijen onder elkaar passen.
    style.configure("Row.TButton", font=("Segoe UI", 12, "bold"), padding=(2, 0))


class MemberPicker(ttk.Frame):
    """
    Zoekveld voor leden. Typ een lidnummer of naam; de lijst eronder toont de treffers.
    Enter kiest de enige treffer (of de geselecteerde in de lijst).

    on_pick(member)            : wordt aangeroepen als er een lid gekozen is
    on_message(tekst, error=)  : om een melding in de statusregel te tonen
    single_click               : False = dubbelklikken op een naam kiest hem;
                                 True = een enkele klik kiest hem meteen (sneller)

    Bediening met het toetsenbord: typen, pijltje omlaag om in de lijst te
    komen, pijltjes om te kiezen, Enter om te bevestigen.
    """

    def __init__(self, master, app, on_pick, on_message, single_click=False):
        super().__init__(master)
        self.app, self.on_pick, self.on_message = app, on_pick, on_message
        self.matches = []  # de leden die nu in de lijst staan, in dezelfde volgorde
        self.var = tk.StringVar()
        # Bij elke toetsaanslag de treffers opnieuw zoeken.
        self.var.trace_add("write", self._refresh)

        ttk.Label(self, text="Lidnummer of naam:").pack(anchor="w")
        self.entry = ttk.Entry(self, textvariable=self.var, font=("Segoe UI", 15))
        self.entry.pack(fill="x", pady=(2, 6))
        self.entry.bind("<Return>", self._enter)
        self.entry.bind("<Down>", self._to_list)

        # exportselection=False: anders verdwijnt de selectie zodra je ergens anders klikt.
        self.list = tk.Listbox(self, height=8, exportselection=False, font=("Segoe UI", 12))
        self.list.pack(fill="both", expand=True)
        self.list.bind("<Return>", self._enter)
        if single_click:
            # Bij het LOSLATEN van de muisknop kiezen, niet bij pijltjestoetsen:
            # anders zou je met pijltje omlaag al een lid kiezen terwijl je nog zoekt.
            self.list.bind("<ButtonRelease-1>", self._click)
        else:
            self.list.bind("<Double-Button-1>", self._enter)

    def focus(self):
        """Zet de cursor in het zoekveld."""
        self.entry.focus_set()

    def _refresh(self, *_):
        """Vult de lijst met treffers voor wat er nu getypt staat."""
        # Maximaal 50 treffers: meer is niet te overzien en maakt typen traag.
        self.matches = find_members(self.app.members, self.var.get())[:50]
        self.list.delete(0, "end")
        for m in self.matches:
            self.list.insert("end", m.label())

    def _to_list(self, _event):
        """Pijltje omlaag in het zoekveld: naar de eerste treffer in de lijst."""
        if self.matches:
            self.list.focus_set()
            self.list.selection_clear(0, "end")
            self.list.selection_set(0)

    def _enter(self, _event=None):
        """
        Enter, dubbelklik of de knop Aanmelden. Kiest in deze volgorde:
          1. het lid dat in de lijst geselecteerd is
          2. de enige treffer, als er precies een is
        Anders volgt een melding, zodat er nooit per ongeluk de verkeerde wordt gekozen.
        """
        if not self.var.get().strip():
            return
        selected = self.list.curselection()
        if selected:
            self._pick(self.matches[selected[0]])
        elif len(self.matches) == 1:
            self._pick(self.matches[0])
        elif not self.matches:
            self.on_message(f"Niemand gevonden voor '{self.var.get().strip()}'. Controleer het lidnummer of de naam.", error=True)
        else:
            self.on_message("Meerdere leden gevonden. Klik op de juiste in de lijst.", error=True)

    def _click(self, event):
        """Enkele klik op een naam in de lijst: kies dat lid meteen."""
        if not self.matches:
            return
        index = self.list.nearest(event.y)
        box = self.list.bbox(index)
        # nearest() geeft ook de laatste regel als je onder de lijst klikt;
        # alleen een klik die echt op een naam valt, telt.
        if box and box[1] <= event.y <= box[1] + box[3]:
            self._pick(self.matches[index])

    def _pick(self, member):
        self.var.set("")  # eerst leegmaken, zodat een volgende zoekopdracht schoon begint
        self.entry.focus_set()
        self.on_pick(member)


class BasePage(ttk.Frame):
    """
    Gemeenschappelijk voor elke pagina: statusregel en automatisch verversen.

    Een kindklasse moet:
      * refresh() maken: gegevens lezen en het scherm bijwerken
      * aan het eind van __init__ self.refresh() en self.start_polling() aanroepen
    En mag:
      * focus_main() maken: waar de cursor staat als de pagina in beeld komt
      * self.message(tekst, error=True/False) gebruiken voor de statusregel

    Het truukje met self._sig ("signature"): refresh() maakt een JSON-tekst van
    de gegevens en vergelijkt die met de vorige keer. Is er niets veranderd, dan
    stopt refresh() meteen. Zo flikkert het scherm niet en blijft een selectie
    in een lijst staan. Wil je het scherm hoe dan ook opnieuw tekenen (bijv. na
    het kiezen van een ander lid), zet dan eerst self._sig = None.
    """

    POLL_MS = 2000  # hoe vaak (in milliseconden) elke pagina de bestanden opnieuw leest

    def __init__(self, app, parent):
        super().__init__(parent)
        self.app = app
        self._sig = None  # zie de uitleg hierboven
        self._job = None  # id van de geplande volgende _poll, om te kunnen stoppen
        self.bind("<Destroy>", self._on_destroy)
        # Statusregel onderaan: groen voor "gelukt", rood voor "let op".
        # Dit is een gewone tk.Label (geen ttk), omdat we de kleur willen wisselen.
        self.status = tk.StringVar()
        self.status_label = tk.Label(self, textvariable=self.status, anchor="w", font=("Segoe UI", 13, "bold"), padx=10, pady=6)
        self.status_label.pack(side="bottom", fill="x")

    def on_show(self):
        """Wordt aangeroepen als de pagina in beeld komt: direct actuele gegevens tonen."""
        self._sig = None
        try:
            self.refresh()
        except StoreError as e:
            self.message(str(e), error=True)
        self.focus_main()

    def focus_main(self):
        """Kindklassen zetten hier de cursor in het juiste veld."""

    def message(self, text, error=False):
        """Toont een melding in de statusregel (rood als error=True)."""
        self.status.set(text)
        self.status_label.configure(bg="#fde2e2" if error else "#e1f3e1", fg="#8a1f1f" if error else "#1d5e1d")

    def start_polling(self):
        """Start het automatisch verversen. Een keer aanroepen per pagina."""
        self._poll()

    def _poll(self):
        """Leest elke paar seconden de gegevens opnieuw: zo zie je wijzigingen van het andere scherm."""
        try:
            self.refresh()
        except StoreError as e:
            # Hier geen venstertje: dat zou elke 2 seconden terugkomen.
            self.message(str(e), error=True)
        except tk.TclError:
            return  # het venster wordt afgesloten; stoppen met verversen
        # after() plant de volgende keer in, zonder het programma te blokkeren.
        self._job = self.after(self.POLL_MS, self._poll)

    def _on_destroy(self, event):
        """Stopt de geplande verversing als de pagina wordt opgeruimd."""
        if event.widget is self and self._job:
            self.after_cancel(self._job)

    def refresh(self):  # wordt door de kindklassen ingevuld
        raise NotImplementedError


class DatumField(ttk.Frame):
    """
    Datumveld met knoppen. De datum is standaard automatisch die van vandaag. Sla je een
    eigen datum op, dan blijft die staan tot je op "Vandaag" klikt of een nieuwe
    maaltijd of soos begint. Het automatische deel zit in auto_datum().

    label     : tekst voor het veld, bijv. "Datum maaltijd:"
    get_store : functie die de JsonStore teruggeeft (bijv. app.meal_store)
    message   : functie om meldingen te tonen (de message van de pagina)
    on_change : wordt aangeroepen na opslaan of terugzetten, zodat de pagina ververst
    """

    def __init__(self, master, label, get_store, message, on_change):
        super().__init__(master)
        self.get_store, self.message, self.on_change = get_store, message, on_change
        self.var = tk.StringVar()  # wat er in het veld staat
        self.info = tk.StringVar()  # "(automatisch: vandaag)" of "(zelf ingesteld)"
        ttk.Label(self, text=label).pack(side="left")
        self.entry = ttk.Entry(self, textvariable=self.var, width=12, font=("Segoe UI", 13))
        self.entry.pack(side="left", padx=6)
        self.entry.bind("<Return>", lambda e: self.save())
        ttk.Button(self, text="Datum opslaan", command=self.save).pack(side="left")
        ttk.Button(self, text="Vandaag", command=self.reset).pack(side="left", padx=6)
        ttk.Label(self, textvariable=self.info, foreground="#555").pack(side="left", padx=4)

    def show(self, data):
        """Toont de datum uit de gegevens, maar overschrijft niets terwijl iemand typt."""
        if self.focus_get() is not self.entry:
            self.var.set(data["datum"])
        self.info.set("(zelf ingesteld)" if data.get("datum_handmatig") else "(automatisch: vandaag)")

    @guarded
    def save(self):
        """Slaat de getypte datum op als handmatige datum."""
        try:
            datum = parse_datum(self.var.get())
        except ValueError:
            self.message("Datum niet herkend. Gebruik dd-mm-jjjj, bijvoorbeeld 30-10-2026.", error=True)
            return
        self.get_store().update(set_manual_datum(datum))
        self.var.set(datum)
        self.message(f"Datum ingesteld op {datum}.")
        self.on_change()

    @guarded
    def reset(self):
        """Zet de datum terug op automatisch (vandaag). auto_datum vult hem bij de volgende leesbeurt in."""
        self.get_store().update(lambda data: data.__setitem__("datum_handmatig", False))
        self.var.set(today_str())
        self.message(f"De datum volgt weer automatisch de dag van vandaag ({today_str()}).")
        self.on_change()


class MaaltijdPage(BasePage):
    """
    Pagina "Maaltijd": leden aanmelden voor de maaltijd.

    Links het zoekveld (MemberPicker), rechts de lijst van aangemelde leden met
    knoppen om het aantal aan te passen of iemand te verwijderen. Bovenin de
    datum, de export en "Nieuwe maaltijd".
    Een lid kan meer dan eens aangemeld worden (bijvoorbeeld met gasten); dat
    staat als "aantal" in de lijst.
    """

    def __init__(self, app, parent):
        super().__init__(app, parent)
        self.message("Typ een lidnummer of naam en klik op de naam om iemand aan te melden.")

        # Bovenbalk: datum en knoppen
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)
        self.datum = DatumField(top, "Datum maaltijd:", app.meal_store, self.message, self._datum_changed)
        self.datum.pack(side="left")
        ttk.Button(top, text="Nieuwe maaltijd", command=self.new_meal).pack(side="right")
        ttk.Button(top, text="Exporteer lijst naar Excel", command=self.export).pack(side="right", padx=8)

        # Middenstuk in twee kolommen. weight 2 en 3: links 40%, rechts 60% van de
        # breedte. uniform="k" houdt die verhouding vast, ook bij verkleinen.
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=10, pady=4)
        body.columnconfigure(0, weight=2, uniform="k")
        body.columnconfigure(1, weight=3, uniform="k")
        body.rowconfigure(0, weight=1)

        # Links: aanmelden
        left = ttk.LabelFrame(body, text=" Aanmelden ", padding=8)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.picker = MemberPicker(left, app, self.sign_up, self.message, single_click=True)
        self.picker.pack(fill="both", expand=True)
        ttk.Button(left, text="Aanmelden", style="Big.TButton", command=self.picker._enter).pack(fill="x", pady=(8, 0))

        # Rechts: de aanmeldlijst. Elke rij heeft als id (iid) het lidnummer,
        # zodat we de geselecteerde rij direct aan een lid kunnen koppelen.
        right = ttk.LabelFrame(body, text=" Aangemeld ", padding=8)
        right.grid(row=0, column=1, sticky="nsew")
        self.tree = ttk.Treeview(right, columns=("nr", "naam", "aantal"), show="headings", selectmode="browse")
        for col, text, width in (("nr", "Lidnr", 80), ("naam", "Naam", 260), ("aantal", "Aantal", 70)):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor="w" if col == "naam" else "center")
        self.tree.pack(fill="both", expand=True)
        buttons = ttk.Frame(right)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="+1", command=lambda: self.change_count(+1)).pack(side="left")
        ttk.Button(buttons, text="-1", command=lambda: self.change_count(-1)).pack(side="left", padx=6)
        ttk.Button(buttons, text="Verwijderen", command=self.remove).pack(side="left")
        self.total_var = tk.StringVar()
        ttk.Label(buttons, textvariable=self.total_var, font=("Segoe UI", 13, "bold")).pack(side="right")

        self.refresh()
        self.start_polling()

    def focus_main(self):
        self.picker.focus()

    # -- tonen --
    def refresh(self):
        """Leest maaltijd.json en tekent de lijst opnieuw als er iets veranderd is."""
        data = self.app.meal_store().read()
        sig = json.dumps(data, sort_keys=True)
        if sig == self._sig:
            return  # niets veranderd
        self._sig = sig
        self.datum.show(data)
        # Selectie onthouden, lijst leegmaken en opnieuw vullen, selectie terugzetten.
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        total = 0
        for nr, e in data["aanmeldingen"].items():
            self.tree.insert("", "end", iid=nr, values=(nr, e["naam"], e["aantal"]))
            total += e["aantal"]
        if selected and self.tree.exists(selected[0]):
            self.tree.selection_set(selected[0])
        self.total_var.set(f"Totaal: {total} personen")

    def _selected_nr(self):
        """Lidnummer van de geselecteerde rij, of None (met een melding) als er niets gekozen is."""
        selected = self.tree.selection()
        if not selected:
            self.message("Klik eerst op een naam in de lijst.", error=True)
            return None
        return selected[0]

    # -- acties --
    @guarded
    def sign_up(self, member):
        """
        Meldt een lid aan (wordt aangeroepen door de MemberPicker).
        Staat het lid er al op, dan vragen we eerst of het echt nog een keer moet;
        zo voorkomen we dubbele aanmeldingen door een keer te vaak op Enter drukken.
        """
        current = self.app.meal_store().read()["aanmeldingen"].get(member.nr)
        if current and not messagebox.askyesno(
            "Al aangemeld",
            f"{member.naam} staat al op de lijst ({current['aantal']}x).\n\nNog een keer aanmelden (bijvoorbeeld voor een gast)?",
            parent=self,
        ):
            self.message(f"{member.naam} is niet nog een keer aangemeld.")
            return

        def add(data):
            # setdefault: bestaat het lid nog niet, dan maken we een nieuwe regel
            # met aantal 0 (en de tijd van nu); daarna tellen we er altijd 1 bij.
            entry = data["aanmeldingen"].setdefault(
                member.nr, {"naam": member.naam, "aantal": 0, "tijd": datetime.datetime.now().strftime("%H:%M")}
            )
            entry["aantal"] += 1

        data = self.app.meal_store().update(add)
        self.message(f"{member.naam} is aangemeld ({data['aanmeldingen'][member.nr]['aantal']}x).")
        self.refresh()

    @guarded
    def change_count(self, delta):
        """Knoppen +1 en -1: het aantal van het geselecteerde lid aanpassen."""
        nr = self._selected_nr()
        if nr is None:
            return
        naam = self.tree.item(nr, "values")[1]
        if delta < 0 and int(self.tree.item(nr, "values")[2]) <= 1:
            self.remove()  # een keer minder dan 1 betekent verwijderen, met bevestiging
            return

        def change(data):
            # Controle "nr in ...": een ander scherm kan het lid net verwijderd hebben.
            if nr in data["aanmeldingen"]:
                data["aanmeldingen"][nr]["aantal"] += delta

        self.app.meal_store().update(change)
        self.message(f"Aantal van {naam} aangepast.")
        self.refresh()

    @guarded
    def remove(self):
        """Knop Verwijderen: het geselecteerde lid helemaal van de lijst halen (na bevestiging)."""
        nr = self._selected_nr()
        if nr is None:
            return
        naam = self.tree.item(nr, "values")[1]
        if not messagebox.askyesno("Verwijderen", f"{naam} helemaal van de lijst verwijderen?", parent=self):
            return
        self.app.meal_store().update(lambda data: data["aanmeldingen"].pop(nr, None))
        self.message(f"{naam} is verwijderd.")
        self.refresh()

    def _datum_changed(self):
        """Na het wijzigen van de datum: cursor terug in het zoekveld en scherm bijwerken."""
        self.picker.focus()
        self.refresh()

    @guarded
    def export(self):
        """Knop Exporteer: schrijft 'Maaltijd <datum>.xlsx' in de opslagmap."""
        path = self.app.export_meal()
        self.message("Export gelukt.")
        messagebox.showinfo("Export gelukt", f"Het bestand is opgeslagen als:\n{path}", parent=self)

    @guarded
    def new_meal(self):
        """
        Knop Nieuwe maaltijd: exporteren, archiveren en leegmaken, in die volgorde.
        Gaat de export mis (bijv. bestand open in Excel), dan stopt alles via een
        StoreError voordat er iets gewist wordt.
        """
        if not messagebox.askyesno(
            "Nieuwe maaltijd",
            "De maaltijdlijst wordt eerst geexporteerd naar Excel en bewaard in het archief. "
            "Daarna begint de lijst leeg. De soos blijft zoals die is.\n\nDoorgaan?",
            parent=self,
        ):
            return
        path = self.app.export_meal()  # mislukt de export, dan wordt er niets gewist
        self.app.archive_and_reset_meal()
        self.refresh()
        self.message("Nieuwe maaltijd gestart.")
        messagebox.showinfo("Klaar", f"De vorige maaltijd is geexporteerd:\n{path}", parent=self)


class SoosPage(BasePage):
    """
    Pagina "Soos": drankjes boeken per lid.

    Werkwijze aan de bar: 1) klik links op een lid (een klik is genoeg),
    2) druk bij het drankje op + of -. Elk drankje is een eenvoudige rij:
        Bier   1,50   [-] [+]   aantal
    De rijen worden automatisch over 1 tot 4 kolommen verdeeld, zodat ze zonder
    scrollen in beeld passen (zie _layout_rows).
    Hier staat bewust alleen de rekening van het gekozen lid, niet die van
    iedereen (privacy). Het overzicht staat op TotalsPage, het beheer van de
    drankjes op DrinksPage.
    """

    MAX_COLUMNS = 4  # nooit meer kolommen dan dit
    BUTTON_GAP = 30  # ruimte (pixels) tussen de min- en de plusknop, tegen misklikken
    ROW_HEIGHT = 36  # aanname voor de hoogte van een rij (pixels), als meten niet lukt

    def __init__(self, app, parent):
        super().__init__(app, parent)
        self.current = None  # het lid (Member) waarvoor we nu boeken, of None
        self.rows = {}  # drank-id (tekst) -> de widgets van die drankrij
        self._rows_sig = None  # welke drankjes de rijen nu tonen (zie _build_rows)
        self.message("Kies eerst een lid: typ een lidnummer of naam en klik op de naam.")

        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)
        self.datum = DatumField(top, "Datum soos:", app.soos_store, self.message, self._datum_changed)
        self.datum.pack(side="left")
        ttk.Button(top, text="Nieuwe soos", command=self.new_soos).pack(side="right")
        ttk.Button(top, text="Exporteer lijst naar Excel", command=self.export).pack(side="right", padx=8)

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=10, pady=4)
        body.columnconfigure(0, weight=1, uniform="k")
        body.columnconfigure(1, weight=3, uniform="k")
        body.rowconfigure(0, weight=1)

        # Links: lid kiezen
        left = ttk.LabelFrame(body, text=" 1. Klik op een lid ", padding=8)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.picker = MemberPicker(left, app, self.pick_member, self.message, single_click=True)
        self.picker.pack(fill="both", expand=True)

        # Rechts: de drankjes van het gekozen lid. Elk drankje is een rij met een
        # - en een + knop (zie _build_rows). Dat kost bij de bar maar een klik per drankje.
        middle = ttk.LabelFrame(body, text=" 2. Druk bij het drankje op + of - ", padding=8)
        middle.grid(row=0, column=1, sticky="nsew", padx=(0, 8))
        self.member_var = tk.StringVar(value="Nog geen lid gekozen")
        ttk.Label(middle, textvariable=self.member_var, font=("Segoe UI", 14, "bold")).pack(anchor="w")
        self.bill_var = tk.StringVar()
        ttk.Label(middle, textvariable=self.bill_var, font=("Segoe UI", 12)).pack(anchor="w", pady=(0, 6))

        # Een gewoon frame kan niet scrollen. Daarom: een Canvas met daarin een frame
        # (rows_frame) met alle drankrijen, plus een scrollbalk. Zijn er weinig
        # drankjes, dan is de scrollbalk niet nodig en zie je hem niet bewegen.
        holder = ttk.Frame(middle)
        holder.pack(fill="both", expand=True)
        background = ttk.Style(self).lookup("TFrame", "background")
        self.canvas = tk.Canvas(holder, highlightthickness=0, bg=background)
        scrollbar = ttk.Scrollbar(holder, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.rows_frame = ttk.Frame(self.canvas)
        self.rows_window = self.canvas.create_window((0, 0), window=self.rows_frame, anchor="nw")
        self._layout_key = None  # (kolommen, rijen per kolom) van de huidige indeling
        # Houd het scrollgebied gelijk aan de inhoud en de inhoud zo breed als het canvas.
        self.rows_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        # Verandert het canvas van grootte (venster groter of kleiner), dan delen we de
        # rijen opnieuw in over de beschikbare breedte en hoogte.
        self.canvas.bind("<Configure>", self._on_canvas_resize)
        # Scrollen met het muiswiel, alleen zolang de muis boven de drankjes hangt.
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_wheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

        self.refresh()
        self.start_polling()

    def focus_main(self):
        self.picker.focus()

    # -- tonen --
    def refresh(self):
        """Leest soos.json en toont de drankjes en de rekening van het gekozen lid."""
        soos = self.app.soos_store().read()
        # Het gekozen lid hoort bij de "signature": een ander lid kiezen moet het
        # scherm ook opnieuw tekenen, ook al zijn de gegevens zelf gelijk.
        sig = json.dumps([soos, self.current.nr if self.current else None], sort_keys=True)
        if sig == self._sig:
            return
        self._sig = sig
        self.datum.show(soos)

        # Rijen opnieuw bouwen als de drankjes veranderd zijn, anders alleen bijwerken.
        drinks_sig = json.dumps([(d["id"], d["naam"], d["cent"]) for d in soos["drankjes"]])
        if drinks_sig != self._rows_sig:
            self._build_rows(soos["drankjes"])
            self._rows_sig = drinks_sig
        counts = soos["verbruik"].get(self.current.nr, {}).get("drank", {}) if self.current else {}
        for d in soos["drankjes"]:
            row = self.rows[str(d["id"])]
            n = counts.get(str(d["id"]), 0)
            # Het aantal is zwart en vet als er iets geboekt is, anders grijs.
            row["count"].configure(text=str(n), foreground="#000" if n else "#999")
            # Zonder gekozen lid zijn de knoppen uit: dan kun je niets op niemand boeken.
            # Min is ook uit bij 0, want lager dan nul kan niet.
            row["plus"].state(["!disabled"] if self.current else ["disabled"])
            row["minus"].state(["!disabled"] if self.current and n else ["disabled"])
        if self.current:
            self.member_var.set(f"{self.current.naam} ({self.current.nr})")
            self.bill_var.set(f"Rekening: EUR {euro(member_total_cents(soos, self.current.nr))}")
        else:
            self.member_var.set("Nog geen lid gekozen")
            self.bill_var.set("")

    def _on_wheel(self, event):
        """Muiswiel scrollt de drankjes, maar alleen als ze niet allemaal in beeld passen."""
        if self.rows_frame.winfo_height() > self.canvas.winfo_height():
            self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def _on_canvas_resize(self, event):
        """Het canvas heeft een nieuwe maat: inhoud even breed maken en rijen opnieuw indelen."""
        self.canvas.itemconfigure(self.rows_window, width=event.width)
        self._layout_rows()

    def _build_rows(self, drinks):
        """
        Bouwt voor elk drankje een eenvoudige rij:

            Bier   1,50   [-] [+]   aantal

        Dit doen we alleen als de drankjes zelf veranderen (nieuw, andere prijs,
        verwijderd), niet bij elke boeking. Daarna past refresh() alleen nog de
        aantallen aan, zodat de knoppen niet flikkeren terwijl je klikt.
        De plek van elke rij bepaalt _layout_rows().
        """
        for child in self.rows_frame.winfo_children():
            child.destroy()
        self.rows = {}
        self._layout_key = None  # er zijn nieuwe rijen, dus de indeling moet opnieuw
        if not drinks:
            ttk.Label(
                self.rows_frame,
                text="Er zijn nog geen drankjes.\nVoeg ze toe via Instellingen > Drankjes en prijzen.",
                justify="left",
            ).grid(row=0, column=0, sticky="w", pady=10)
            return
        # Alle namen even breed (in tekens), zodat prijs en knoppen in elke rij op
        # dezelfde plek staan, ook als de namen verschillend lang zijn.
        name_width = max(12, max(len(d["naam"]) for d in drinks) + 1)
        for d in drinks:
            drink_id = str(d["id"])  # tekst, want zo staat het in soos.json
            row = ttk.Frame(self.rows_frame)
            ttk.Label(row, text=d["naam"], width=name_width, font=("Segoe UI", 13)).grid(row=0, column=0, sticky="w", padx=(6, 6))
            ttk.Label(row, text=euro(d["cent"]), width=7, anchor="e", font=("Segoe UI", 12), foreground="#555").grid(row=0, column=1, padx=4)
            # i=drink_id in de lambda: anders zouden alle knoppen het laatste drankje boeken.
            # Volgorde: eerst min, dan plus. Tussen de twee knoppen zit bewust een ruime
            # opening (BUTTON_GAP pixels), zodat je niet per ongeluk de verkeerde raakt.
            gap = self.BUTTON_GAP // 2
            minus = ttk.Button(row, text="-", width=3, style="Row.TButton", command=lambda i=drink_id: self.change(i, -1))
            minus.grid(row=0, column=2, padx=(6, gap), pady=1)
            plus = ttk.Button(row, text="+", width=3, style="Row.TButton", command=lambda i=drink_id: self.change(i, +1))
            plus.grid(row=0, column=3, padx=(gap, 2), pady=1)
            count = ttk.Label(row, text="0", width=3, anchor="center", font=("Segoe UI", 13, "bold"))
            count.grid(row=0, column=4, padx=(4, 6))
            self.rows[drink_id] = {"naam": d["naam"], "frame": row, "count": count, "minus": minus, "plus": plus}
        self._layout_rows()

    def _layout_rows(self):
        """
        Verdeelt de rijen over kolommen, zo dat alles zonder scrollen in beeld past.

        Hoeveel rijen er onder elkaar passen volgt uit de hoogte van het canvas. Passen
        niet alle drankjes onder elkaar, dan komt er een kolom bij, tot MAX_COLUMNS en
        zolang elke kolom breed genoeg blijft voor de breedste rij. Past het dan nog
        niet (heel veel drankjes of een heel klein venster), dan blijft scrollen mogelijk.

        De rijen staan op volgorde van boven naar beneden, dan de volgende kolom, zoals
        een menukaart. Er wordt alleen opnieuw geplaatst als de indeling echt verandert.
        """
        count = len(self.rows)
        if not count:
            return
        height, width = self.canvas.winfo_height(), self.canvas.winfo_width()
        if height <= 1 or width <= 1:
            return  # nog niet getekend; de eerste <Configure> roept ons opnieuw aan
        # Meten wat een rij echt nodig heeft. update_idletasks laat tkinter eerst de maten
        # van net gebouwde rijen uitrekenen; zonder dat zijn ze nog 1 pixel groot.
        self.rows_frame.update_idletasks()
        frames = [entry["frame"] for entry in self.rows.values()]
        row_height = frames[0].winfo_reqheight() or self.ROW_HEIGHT
        # De breedste rij bepaalt hoe smal een kolom mag worden (12 = ruimte om de rij heen).
        column_width = max(frame.winfo_reqwidth() for frame in frames) + 12
        rows_fit = max(1, height // row_height)
        columns_needed = -(-count // rows_fit)  # naar boven afgerond delen
        columns_possible = max(1, width // column_width)
        columns = max(1, min(columns_needed, columns_possible, self.MAX_COLUMNS))
        per_column = -(-count // columns)
        if (columns, per_column) == self._layout_key:
            return
        self._layout_key = (columns, per_column)
        # Kolommen die we gebruiken delen de breedte gelijk; ongebruikte krijgen geen ruimte.
        for column in range(self.MAX_COLUMNS):
            if column < columns:
                self.rows_frame.columnconfigure(column, weight=1, uniform="drank")
            else:
                self.rows_frame.columnconfigure(column, weight=0, uniform="")
        for index, entry in enumerate(self.rows.values()):
            # sticky="w": de rij houdt zijn eigen breedte (links uitgelijnd) en rekt niet over
            # het hele scherm uit. Anders staat de naam ver van de knoppen af.
            entry["frame"].grid(row=index % per_column, column=index // per_column, sticky="w", padx=6)

    def _force_refresh(self):
        """Scherm hoe dan ook opnieuw tekenen (zie de uitleg over _sig in BasePage)."""
        self._sig = None
        self.refresh()

    # -- lid kiezen --
    def pick_member(self, member):
        """Wordt aangeroepen door de MemberPicker: dit lid wordt het 'huidige' lid."""
        self.current = member
        self._force_refresh()
        self.message(f"{member.naam} gekozen. Druk bij een drankje op + of -.")

    # -- plus en min --
    @guarded
    def change(self, drink_id, delta):
        """
        Plus (+1) of min (-1) voor een drankje van het huidige lid.
        drink_id is het id van het drankje als tekst; de knoppen in _build_rows geven dat mee.
        """
        if self.current is None:
            self.message("Klik eerst links op een lid.", error=True)
            return
        member = self.current
        # `result` is een doosje om het nieuwe aantal uit de functie apply() naar
        # buiten te krijgen; apply() draait binnen JsonStore.update onder het slot.
        result = {}

        def apply(data):
            entry = data["verbruik"].setdefault(member.nr, {"naam": member.naam, "drank": {}})
            now = entry["drank"].get(drink_id, 0)
            result["new"] = max(0, now + delta)  # nooit onder nul
            if result["new"]:
                entry["drank"][drink_id] = result["new"]
            else:
                entry["drank"].pop(drink_id, None)

        self.app.soos_store().update(apply)
        name = self.rows[drink_id]["naam"]
        if delta < 0 and result["new"] == 0:
            self.message(f"{name} voor {member.naam} staat nu op 0.")
        else:
            self.message(f"{member.naam}: {result['new']}x {name}.")
        self._force_refresh()

    def _datum_changed(self):
        """Na het wijzigen van de datum: cursor terug in het zoekveld en scherm bijwerken."""
        self.picker.focus()
        self._force_refresh()

    @guarded
    def new_soos(self):
        """
        Knop Nieuwe soos: exporteren, archiveren en de rekeningen leegmaken.
        De drankjes en prijzen blijven staan, de maaltijd wordt niet aangeraakt.
        """
        if not messagebox.askyesno(
            "Nieuwe soos",
            "De soos-rekeningen worden eerst geexporteerd naar Excel en bewaard in het archief. "
            "Daarna beginnen ze leeg (de drankjes blijven staan). De maaltijdlijst blijft zoals die is.\n\nDoorgaan?",
            parent=self,
        ):
            return
        path = self.app.export_soos()  # mislukt de export, dan wordt er niets gewist
        self.app.archive_and_reset_soos()
        self.current = None
        self._force_refresh()
        self.message("Nieuwe soos gestart.")
        messagebox.showinfo("Klaar", f"De vorige soos is geexporteerd:\n{path}", parent=self)

    @guarded
    def export(self):
        """Knop Exporteer: schrijft 'Soos <datum>.xlsx' in de opslagmap."""
        path = self.app.export_soos()
        self.message("Export gelukt.")
        messagebox.showinfo("Export gelukt", f"Het bestand is opgeslagen als:\n{path}", parent=self)


class DrinksPage(BasePage):
    """
    Drankjes toevoegen, verwijderen en prijzen aanpassen. Staat op een eigen pagina,
    zodat niemand aan de bar per ongeluk (of expres) een prijs wijzigt.
    """

    def __init__(self, app, parent):
        super().__init__(app, parent)
        self.message("Selecteer een drankje om de prijs te wijzigen of het te verwijderen.")
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Drankjes en prijzen", font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(0, 10))

        self.tree = ttk.Treeview(frame, columns=("naam", "prijs"), show="headings", selectmode="browse")
        self.tree.heading("naam", text="Drankje")
        self.tree.heading("prijs", text="Prijs (euro)")
        self.tree.column("naam", width=300)
        self.tree.column("prijs", width=120, anchor="e")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # Nieuw drankje
        add = ttk.LabelFrame(frame, text=" Nieuw drankje ", padding=8)
        add.pack(fill="x", pady=(10, 0))
        self.new_name, self.new_price = tk.StringVar(), tk.StringVar()
        ttk.Label(add, text="Naam").pack(side="left")
        name_entry = ttk.Entry(add, textvariable=self.new_name, width=24)
        name_entry.pack(side="left", padx=6)
        ttk.Label(add, text="Prijs").pack(side="left")
        price_entry = ttk.Entry(add, textvariable=self.new_price, width=8)
        price_entry.pack(side="left", padx=6)
        name_entry.bind("<Return>", lambda e: price_entry.focus_set())
        price_entry.bind("<Return>", lambda e: self.add_drink())
        ttk.Button(add, text="Toevoegen", command=self.add_drink).pack(side="left")
        self.name_entry = name_entry

        # Geselecteerd drankje
        edit = ttk.LabelFrame(frame, text=" Geselecteerd drankje ", padding=8)
        edit.pack(fill="x", pady=(10, 0))
        self.edit_price = tk.StringVar()
        ttk.Label(edit, text="Nieuwe prijs").pack(side="left")
        edit_entry = ttk.Entry(edit, textvariable=self.edit_price, width=8)
        edit_entry.pack(side="left", padx=6)
        edit_entry.bind("<Return>", lambda e: self.change_price())
        ttk.Button(edit, text="Prijs wijzigen", command=self.change_price).pack(side="left")
        ttk.Button(edit, text="Drankje verwijderen", command=self.delete_drink).pack(side="right")

        self.refresh()
        self.start_polling()

    def focus_main(self):
        self.name_entry.focus_set()

    def refresh(self):
        """Toont de drankjes en prijzen. Kijkt alleen naar de drankjes, niet naar
        het verbruik, zodat boekingen aan de bar deze pagina niet laten verversen."""
        soos = self.app.soos_store().read()
        sig = json.dumps(soos["drankjes"], sort_keys=True)
        if sig == self._sig:
            return
        self._sig = sig
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        for d in soos["drankjes"]:
            self.tree.insert("", "end", iid=str(d["id"]), values=(d["naam"], euro(d["cent"])))
        if selected and self.tree.exists(selected[0]):
            self.tree.selection_set(selected[0])

    def _on_select(self, _event):
        """Vult de prijs van het gekozen drankje alvast in om te wijzigen."""
        selected = self.tree.selection()
        if selected:
            self.edit_price.set(self.tree.item(selected[0], "values")[1])

    def _selected(self):
        """(drank-id, naam) van het geselecteerde drankje, of (None, None) met een melding."""
        selected = self.tree.selection()
        if not selected:
            self.message("Klik eerst op een drankje in de lijst.", error=True)
            return None, None
        return selected[0], self.tree.item(selected[0], "values")[0]

    def _force_refresh(self):
        """Scherm hoe dan ook opnieuw tekenen (zie de uitleg over _sig in BasePage)."""
        self._sig = None
        self.refresh()

    @guarded
    def add_drink(self):
        """Knop Toevoegen: nieuw drankje met naam en prijs. Dubbele namen worden geweigerd."""
        naam = self.new_name.get().strip()
        if not naam:
            self.message("Vul de naam van het drankje in.", error=True)
            return
        try:
            cent = parse_prijs(self.new_price.get())
        except ValueError as e:
            self.message(str(e), error=True)
            return
        result = {}

        def add(data):
            # Controle binnen het slot: zo kan een ander scherm niet tegelijk
            # hetzelfde drankje toevoegen.
            if any(d["naam"].lower() == naam.lower() for d in data["drankjes"]):
                result["duplicate"] = True
                return
            # Ids worden nooit hergebruikt, ook niet na verwijderen. Zo kan een
            # oud archief nooit naar het verkeerde drankje verwijzen.
            data["drankjes"].append({"id": data["volgende_id"], "naam": naam, "cent": cent})
            data["volgende_id"] += 1

        self.app.soos_store().update(add)
        if result.get("duplicate"):
            self.message(f"Het drankje '{naam}' bestaat al.", error=True)
            return
        self.new_name.set("")
        self.new_price.set("")
        self.name_entry.focus_set()
        self.message(f"Drankje '{naam}' (EUR {euro(cent)}) toegevoegd.")
        self._force_refresh()

    @guarded
    def change_price(self):
        """
        Knop Prijs wijzigen. Is het drankje al geboekt, dan vragen we eerst
        bevestiging, want de rekeningen worden met de nieuwe prijs berekend
        (zie member_total_cents).
        """
        drink_id, name = self._selected()
        if drink_id is None:
            return
        try:
            cent = parse_prijs(self.edit_price.get())
        except ValueError as e:
            self.message(str(e), error=True)
            return
        soos = self.app.soos_store().read()
        if any(e["drank"].get(drink_id) for e in soos["verbruik"].values()) and not messagebox.askyesno(
            "Prijs wijzigen",
            f"'{name}' is deze avond al geboekt. De nieuwe prijs geldt dan ook voor de al geboekte drankjes.\n\nDoorgaan?",
            parent=self,
        ):
            return

        def change(data):
            for d in data["drankjes"]:
                if str(d["id"]) == drink_id:
                    d["cent"] = cent

        self.app.soos_store().update(change)
        self.message(f"Prijs van '{name}' is nu EUR {euro(cent)}.")
        self._force_refresh()

    @guarded
    def delete_drink(self):
        """Knop Drankje verwijderen. Kan alleen als het drankje nergens geboekt staat."""
        drink_id, name = self._selected()
        if drink_id is None:
            return
        if not messagebox.askyesno("Drankje verwijderen", f"Drankje '{name}' verwijderen?", parent=self):
            return
        result = {}

        def delete(data):
            # Een drankje dat al geboekt is blijft staan, anders klopt de rekening niet meer.
            if any(e["drank"].get(drink_id) for e in data["verbruik"].values()):
                result["in_use"] = True
                return
            data["drankjes"] = [d for d in data["drankjes"] if str(d["id"]) != drink_id]

        self.app.soos_store().update(delete)
        if result.get("in_use"):
            self.message(f"'{name}' is al geboekt en kan niet verwijderd worden. Gebruik 'Nieuwe soos' of boek het terug naar 0.", error=True)
        else:
            self.message(f"Drankje '{name}' verwijderd.")
        self._force_refresh()


class TotalsPage(BasePage):
    """
    Totalen van de avond: wie hoeveel soos heeft en hoeveel mensen er eten.
    Staat bewust op een eigen pagina, zodat de rekeningen van anderen niet
    zichtbaar zijn voor wie aan de bar of bij de aanmelding zit.
    """

    def __init__(self, app, parent):
        super().__init__(app, parent)
        self.message("Totalen van de avond.")
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)
        self.title_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.title_var, font=("Segoe UI", 18, "bold")).pack(anchor="w")
        self.meal_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.meal_var, font=("Segoe UI", 13)).pack(anchor="w", pady=(4, 10))

        self.tree = ttk.Treeview(frame, columns=("nr", "naam", "totaal"), show="headings", selectmode="none")
        for col, text, width, anchor in (("nr", "Lidnr", 90, "center"), ("naam", "Naam", 300, "w"), ("totaal", "Soos (euro)", 120, "e")):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor=anchor)
        self.tree.pack(fill="both", expand=True)
        self.grand_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.grand_var, font=("Segoe UI", 14, "bold")).pack(anchor="e", pady=(8, 0))

        # Geen self.refresh() hier: start_polling() roept refresh() direct aan.
        self.start_polling()

    def refresh(self):
        """Leest beide bestanden en toont het aantal eters en alle soosrekeningen."""
        meal = self.app.meal_store().read()
        soos = self.app.soos_store().read()
        sig = json.dumps([meal, soos], sort_keys=True)
        if sig == self._sig:
            return
        self._sig = sig
        self.title_var.set("Totalen van de avond")
        persons = sum(e["aantal"] for e in meal["aanmeldingen"].values())
        self.meal_var.set(f"Maaltijd {meal['datum']}: {persons} personen aangemeld ({len(meal['aanmeldingen'])} leden)")
        self.tree.delete(*self.tree.get_children())
        grand = 0
        for nr, e in soos["verbruik"].items():
            cents = member_total_cents(soos, nr)
            if cents or any(e["drank"].values()):
                self.tree.insert("", "end", iid=nr, values=(nr, e["naam"], euro(cents)))
                grand += cents
        self.grand_var.set(f"Totaal soos {soos['datum']}: EUR {euro(grand)}")


class SettingsPage(BasePage):
    """
    Pagina "Instellingen": opslagmap en ledenbestand instellen.

    De instellingen worden pas bewaard na een geslaagde controle: de map moet
    beschrijfbaar zijn en het ledenbestand moet leden bevatten. Bestaande
    gegevens worden bij een andere opslagmap niet meeverhuisd; kopieer die zo
    nodig zelf (maaltijd.json, soos.json en eventueel de map archief).
    """

    def __init__(self, app, parent):
        super().__init__(app, parent)
        self.dir_var = tk.StringVar()
        self.file_var = tk.StringVar()
        self.info_var = tk.StringVar()

        frame = ttk.Frame(self, padding=20)
        frame.pack(pady=20)
        ttk.Label(frame, text="Instellingen", font=("Segoe UI", 18, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 14))
        ttk.Label(frame, text="Opslagmap (gegevens, exports, archief):").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.dir_var, width=70).grid(row=2, column=0, pady=(2, 10))
        ttk.Button(frame, text="Bladeren...", command=self.browse_dir).grid(row=2, column=1, padx=6, pady=(2, 10))
        ttk.Label(frame, text="Ledenbestand (Excel .xlsx):").grid(row=3, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.file_var, width=70).grid(row=4, column=0, pady=(2, 10))
        ttk.Button(frame, text="Bladeren...", command=self.browse_file).grid(row=4, column=1, padx=6, pady=(2, 10))
        ttk.Label(
            frame,
            text="Tip: kies voor de opslagmap een gedeelde map als meerdere computers\nsamen willen werken. Bestaande gegevens worden niet verplaatst.",
            foreground="#555",
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(0, 10))
        buttons = ttk.Frame(frame)
        buttons.grid(row=6, column=0, columnspan=2, sticky="w")
        ttk.Button(buttons, text="Opslaan", style="Big.TButton", command=self.save).pack(side="left")
        ttk.Button(buttons, text="Ledenlijst opnieuw inlezen", command=self.reload).pack(side="left", padx=10)
        ttk.Label(frame, textvariable=self.info_var, justify="left").grid(row=7, column=0, columnspan=2, sticky="w", pady=(16, 0))

    def refresh(self):
        pass  # niets te verversen; de velden worden gevuld in on_show

    def on_show(self):
        """Toont altijd de opgeslagen waarden, ook als er eerder iets half is getypt."""
        self.dir_var.set(self.app.data_dir)
        self.file_var.set(self.app.members_path)
        self.update_info()
        if self.app.members_error:
            self.message(self.app.members_error.split("\n")[0], error=True)
        else:
            self.message("Wijzig de instellingen en klik op Opslaan.")

    def update_info(self):
        """Toont onderaan hoeveel leden er zijn ingelezen, en werkt de navbar bij."""
        if self.app.members_error:
            self.info_var.set("Let op: " + self.app.members_error)
        else:
            self.info_var.set(f"{len(self.app.members)} leden ingelezen uit:\n{self.app.members_path}")
        self.app.navbar.update_info()

    def browse_dir(self):
        """Knop Bladeren bij de opslagmap: Windows-venster om een map te kiezen."""
        chosen = filedialog.askdirectory(parent=self, initialdir=self.dir_var.get() or APP_DIR)
        if chosen:
            self.dir_var.set(os.path.normpath(chosen))

    def browse_file(self):
        """Knop Bladeren bij het ledenbestand: Windows-venster om een .xlsx te kiezen."""
        chosen = filedialog.askopenfilename(parent=self, filetypes=[("Excel", "*.xlsx")], initialdir=os.path.dirname(self.file_var.get()) or APP_DIR)
        if chosen:
            self.file_var.set(os.path.normpath(chosen))

    @guarded
    def save(self):
        """Knop Opslaan: eerst alles controleren, pas daarna echt opslaan."""
        data_dir = self.dir_var.get().strip()
        members_file = self.file_var.get().strip()
        if not data_dir:
            raise StoreError("Vul een opslagmap in.")
        # Proefbestand schrijven en weer weghalen: de enige betrouwbare manier om
        # te weten of we in de map mogen schrijven (ook op een netwerkschijf).
        try:
            os.makedirs(data_dir, exist_ok=True)
            probe = os.path.join(data_dir, ".schrijftest")
            with open(probe, "w") as f:
                f.write("ok")
            os.remove(probe)
        except OSError:
            raise StoreError("In deze map kan niet worden geschreven. Kies een andere map.")
        load_members(members_file)  # gooit een duidelijke fout als het bestand niet klopt

        self.app.config["data_dir"] = data_dir
        # Leeg laten als het het standaard ledenbestand in de opslagmap is.
        default_members = os.path.join(data_dir, "leden.xlsx")
        self.app.config["members_file"] = "" if os.path.normpath(members_file) == os.path.normpath(default_members) else members_file
        save_config(self.app.config)
        self.app.reload_members()
        self.update_info()
        self.message("Instellingen opgeslagen.")

    @guarded
    def reload(self):
        """Knop Ledenlijst opnieuw inlezen: handig na het bijwerken van het Excel-bestand."""
        if self.app.reload_members():
            self.message(f"{len(self.app.members)} leden ingelezen.")
        else:
            self.message("De ledenlijst kon niet worden ingelezen.", error=True)
        self.update_info()


class Navbar(tk.Frame):
    """
    Balk bovenin: Maaltijd, Soos en Instellingen.
    Maaltijd en Soos zijn gewone knoppen voor het werk aan de bar en bij de
    aanmelding. Instellingen is een uitklapmenu met alles wat beheer is
    (drankjes en prijzen, totalen, opslag en ledenlijst). Zo staan de
    beheerpagina's niet zichtbaar naast de dagelijkse knoppen.
    De knop van het onderdeel waar je zit, kleurt blauw.
    """

    BG, ACTIVE = "#1f2d3d", "#2f7fd1"  # donkerblauwe balk, lichtblauw voor de huidige knop
    # Logo's naast maaltijd.py. De navbar-versie is vooraf netjes verkleind tot
    # 40 pixels hoog. Ontbreekt die, dan verkleint tkinter het grote logo zelf
    # (dat ziet er wat rafeliger uit, omdat tkinter alleen pixels kan weglaten).
    LOGO_FILE = "ichthus-logo-gray-navbar.png"
    LOGO_FILE_LARGE = "ichthus-logo-gray.png"
    LOGO_BG = "#f4f5f7"  # licht vlak achter het grijze logo, voor genoeg contrast
    # Welke pagina bij welke knop hoort (voor het blauw kleuren).
    GROUP = {"maaltijd": "maaltijd", "soos": "soos", "drankjes": "instellingen", "totalen": "instellingen", "instellingen": "instellingen"}
    # De pagina's onder het Instellingen-menu: (paginanaam, tekst in het menu).
    # None betekent een scheidingslijn in het menu.
    SETTINGS_MENU = (
        ("drankjes", "Drankjes en prijzen"),
        ("totalen", "Totalen van de avond"),
        None,
        ("instellingen", "Opslag en ledenlijst"),
    )

    def __init__(self, app):
        super().__init__(app.root, bg=self.BG)
        self.app = app
        self._add_logo()
        self.buttons = {}
        self._add_button("maaltijd", "Maaltijd")
        self._add_button("soos", "Soos")
        self._add_settings_dropdown()
        self.info = tk.Label(self, bg=self.BG, fg="#cbd5e1", padx=16)
        self.info.pack(side="right")
        self.update_info()

    def _add_logo(self):
        """
        Logo links in de balk. Het logo is grijs (#808080) op een transparante
        achtergrond; op de donkerblauwe balk is dat slecht leesbaar, daarom
        staat het op een eigen licht vlak (LOGO_BG).

        Volgorde van proberen:
          1. LOGO_FILE (al op maat, 216x40)
          2. LOGO_FILE_LARGE, 3x verkleind met subsample(3) (593x110 -> 198x37).
             tkinter kan alleen met hele getallen verkleinen.
          3. De naam als tekst, zodat het programma altijd blijft werken.
        Een ander logo? Maak een PNG van ongeveer 40 pixels hoog en zet de naam
        in LOGO_FILE.
        """
        try:
            image = tk.PhotoImage(file=os.path.join(APP_DIR, self.LOGO_FILE))
        except tk.TclError:
            try:
                image = tk.PhotoImage(file=os.path.join(APP_DIR, self.LOGO_FILE_LARGE)).subsample(3)
            except tk.TclError:
                tk.Label(self, text="Ichthus Utrecht", bg=self.BG, fg="white", font=("Segoe UI", 14, "bold"), padx=16).pack(side="left")
                return
        # Bewaar een verwijzing in self.logo: zonder die verwijzing ruimt Python
        # het plaatje op en blijft er een leeg vlak over.
        self.logo = image
        tk.Label(self, image=self.logo, bg=self.LOGO_BG, padx=14, pady=4).pack(side="left", fill="y")

    def _style(self):
        """Gedeelde opmaak voor alle navbar-knoppen (platte knoppen zonder rand)."""
        return dict(
            relief="flat", bd=0, cursor="hand2", font=("Segoe UI", 13, "bold"), padx=20, pady=10,
            bg=self.BG, fg="white", activebackground=self.ACTIVE, activeforeground="white",
        )

    def _add_button(self, key, label):
        """Gewone navbar-knop die pagina `key` opent."""
        button = tk.Button(self, text=label, command=lambda: self.app.show_page(key), **self._style())
        button.pack(side="left")
        self.buttons[key] = button

    def _add_settings_dropdown(self):
        """Instellingen is een knop die een uitklapmenu opent (items staan in SETTINGS_MENU)."""
        # ▾ is een klein pijltje naar beneden, als hint dat er een menu is.
        button = tk.Menubutton(self, text="Instellingen  ▾", **self._style())
        # tearoff=0: zonder het ouderwetse streepje om het menu los te trekken.
        menu = tk.Menu(button, tearoff=0, font=("Segoe UI", 12))
        for item in self.SETTINGS_MENU:
            if item is None:
                menu.add_separator()
                continue
            page, label = item
            # k=page in de lambda: anders zouden alle items naar de laatste pagina wijzen.
            menu.add_command(label=label, command=lambda k=page: self.app.show_page(k))
        button.configure(menu=menu)
        button.pack(side="left")
        self.buttons["instellingen"] = button

    def highlight(self, name):
        """Kleurt de knop van de huidige pagina."""
        current = self.GROUP.get(name)
        for key, button in self.buttons.items():
            button.configure(bg=self.ACTIVE if key == current else self.BG)

    def update_info(self):
        """Rechts in de balk: aantal leden, of in rood dat er geen ledenlijst is."""
        if self.app.members_error:
            self.info.configure(text="Geen ledenlijst", fg="#ff9b9b")
        else:
            self.info.configure(text=f"{len(self.app.members)} leden", fg="#cbd5e1")


def main():
    """Startpunt van het programma."""
    app = App()  # venster en instellingen
    setup_style(app.root)  # lettergroottes en stijlen (moet voor build)
    app.build()  # navbar en pagina's
    app.root.mainloop()  # wacht op klikken en toetsen tot het venster dicht gaat


if __name__ == "__main__":
    main()
