"""
3. De ledenlijst inlezen en leden zoeken op lidnummer of naam.
"""

import os

from .excel import read_xlsx_rows
from .instellingen import StoreError


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
