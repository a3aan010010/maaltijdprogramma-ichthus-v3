"""
2. Excel (xlsx) lezen en schrijven, zonder externe libraries.
"""

import os
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from .instellingen import StoreError


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
