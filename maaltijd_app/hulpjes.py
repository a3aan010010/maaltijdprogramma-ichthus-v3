"""
4. Geld en datum: omzetten tussen tekst en centen of dd-mm-jjjj.
"""

import datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation


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
