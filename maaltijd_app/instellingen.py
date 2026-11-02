"""
1. Instellingen: de config, de vaste paden naast maaltijd.py en StoreError.
"""

import json
import os


# ---------------------------------------------------------------------------
# 1. Instellingen
# ---------------------------------------------------------------------------

# De map waar maaltijd.py staat (een map boven deze map maaltijd_app). De
# instellingen, het foutlog en de logo's staan altijd hier, zodat het programma
# ze terugvindt, ook als de opslagmap zelf is verplaatst.
APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(APP_DIR, "maaltijd_config.json")
# Onverwachte fouten worden met alle details hierin bijgeschreven (zie App.log_error).
LOG_PATH = os.path.join(APP_DIR, "maaltijd_fout.log")

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
