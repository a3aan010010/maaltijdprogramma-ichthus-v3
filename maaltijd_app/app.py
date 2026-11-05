"""
6. App: het hoofdvenster en de logica die niet bij een enkele pagina hoort (export, archief).
"""

import ctypes
import datetime
import os
import shutil
import sys
import tkinter as tk
import traceback
from tkinter import messagebox

from . import __version__
from .excel import STYLE_BOLD, STYLE_BOLD_MONEY, STYLE_MONEY, write_xlsx
from .hulpjes import euro, today_str
from .instellingen import APP_DIR, LOG_PATH, StoreError, load_config
from .leden import load_members
from .navbar import Navbar
from .opslag import JsonStore, auto_datum, excel_nr, export_path, meal_default, member_total_cents, soos_default
from .paginas import DrinksPage, MaaltijdPage, SettingsPage, SoosPage, TotalsPage


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

    # Het Ichthus-visje, naast maaltijd.py. Een .ico bevat het visje in alle maten
    # (16 tot 256 pixels); Windows kiest zelf de scherpste voor titelbalk en taakbalk.
    ICON_FILE = "ichthus-visje.ico"

    def __init__(self):
        self._own_taskbar_button()  # moet voor het venster bestaat
        self.root = tk.Tk()
        self.root.title(f"Maaltijd en soos {__version__}")
        self._apply_icon()
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

    @staticmethod
    def _own_taskbar_button():
        """
        Het programma draait als python.exe, dus Windows toont in de taakbalk anders
        het Python-icoon. Met een eigen AppUserModelID krijgt het venster een eigen
        taakbalkknop, met het visje erop.
        """
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("IchthusUtrecht.MaaltijdEnSoos")
        except (AttributeError, OSError):
            pass  # geen Windows: niets te doen

    def _apply_icon(self):
        """
        Het visje linksboven in het venster in plaats van de tkinter-veer.
        default=: geldt ook voor de vraagvensters (Ja/Nee, Nieuwe soos).
        Ontbreekt het bestand, of is het geen Windows, dan blijft het standaardicoon.
        """
        try:
            self.root.iconbitmap(default=os.path.join(APP_DIR, self.ICON_FILE))
        except tk.TclError:
            pass

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
        op schijf staat, gaat er bij zo'n fout geen werk verloren. De volledige
        foutmelding staat daarna in maaltijd_fout.log (zie log_error).
        """
        if isinstance(val, StoreError):
            messagebox.showerror("Let op", str(val))
        else:
            self.log_error(exc, val, tb)
            messagebox.showerror(
                "Onverwachte fout",
                f"Er ging iets mis: {val}\n\nDe gegevens zijn veilig, probeer het nog eens.\n"
                f"De details staan in {os.path.basename(LOG_PATH)}.",
            )

    @staticmethod
    def log_error(exc, val, tb):
        """Schrijft een fout met datum, tijd en alle details onderaan maaltijd_fout.log."""
        try:
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write(f"--- {datetime.datetime.now():%d-%m-%Y %H:%M:%S}\n")
                f.write("".join(traceback.format_exception(exc, val, tb)))
        except OSError:
            pass  # loggen mag zelf nooit een nieuwe fout geven

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
    # weg in de opslagmap. De vorige export van dezelfde lijst wordt vervangen
    # (zo kun je tijdens de avond gerust vaker exporteren), die van een andere
    # lijst nooit (zie export_path in opslag.py).

    def _export_path(self, prefix, data):
        """Pad van het exportbestand voor deze lijst."""
        try:
            os.makedirs(self.data_dir, exist_ok=True)
        except OSError as e:
            raise StoreError(f"De opslagmap kon niet worden aangemaakt: {e}")
        return export_path(self.data_dir, prefix, data["datum"], data.get("export_bestand"))

    @staticmethod
    def _remember_export(store, path):
        """Onthoudt in de gegevens welk bestand de export van deze lijst is."""
        store.update(lambda data: data.__setitem__("export_bestand", os.path.basename(path)))

    def export_meal(self):
        """Schrijft 'Maaltijd <datum>.xlsx' met de aanmeldlijst. Geeft het pad terug."""
        store = self.meal_store()
        meal = store.read()
        datum = meal["datum"]
        # Indeling: rij 1 titel, rij 2 leeg, rij 3 kopjes, dan een rij per lid,
        # een lege rij en tot slot de totaalregel.
        rows = [[(f"Maaltijd {datum}", STYLE_BOLD)], [], [(h, STYLE_BOLD) for h in ("Lidnummer", "Naam", "Aantal", "Aangemeld om")]]
        total = 0
        for nr, e in meal["aanmeldingen"].items():
            rows.append([excel_nr(nr), e["naam"], e["aantal"], e.get("tijd", "")])
            total += e["aantal"]
        rows += [[], [("Totaal", STYLE_BOLD), "", (total, STYLE_BOLD)]]
        path = self._export_path("Maaltijd", meal)
        write_xlsx(path, [("Maaltijd", rows)])
        self._remember_export(store, path)
        return path

    def export_soos(self):
        """Schrijft 'Soos <datum>.xlsx' met de rekeningen, een kolom per drankje."""
        store = self.soos_store()
        soos = store.read()
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
                [excel_nr(nr), e["naam"]]
                + [c or "" for c in counts]  # 0 tonen we als lege cel, dat leest rustiger
                + [(cents / 100, STYLE_MONEY)]  # Excel wil euro's, wij rekenen in centen
            )
        rows += [[], [("Totaal", STYLE_BOLD), ""] + [(c, STYLE_BOLD) for c in column_sums] + [(grand / 100, STYLE_BOLD_MONEY)]]
        path = self._export_path("Soos", soos)
        write_xlsx(path, [("Soos", rows)])
        self._remember_export(store, path)
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
        try:
            os.makedirs(folder, exist_ok=True)
            if os.path.exists(store.path):
                shutil.copy2(store.path, folder)
        except OSError as e:
            # StoreError stopt "Nieuwe ..." voordat er iets gewist wordt.
            raise StoreError(f"De kopie voor het archief kon niet worden gemaakt: {e}")
        return folder

    def archive_and_reset_meal(self):
        """Bewaart de aanmeldlijst in het archief en begint een lege maaltijd. De soos blijft ongemoeid."""
        store = self.meal_store()
        folder = self._archive(store, "Maaltijd", store.read()["datum"])

        def reset(data):
            data.clear()  # ook export_bestand weg: de volgende export hoort bij een nieuwe lijst
            data.update(meal_default())

        store.update(reset)
        return folder

    def archive_and_reset_soos(self):
        """Bewaart de soos in het archief en begint leeg. De drankjes en de maaltijd blijven ongemoeid."""
        store = self.soos_store()
        folder = self._archive(store, "Soos", store.read()["datum"])

        def reset(data):
            data["verbruik"] = {}  # de drankjes zelf blijven bestaan
            data["datum"] = today_str()
            data["datum_handmatig"] = False
            data.pop("export_bestand", None)  # de volgende export hoort bij een nieuwe soos

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
        # Verborgen pagina's lezen niets; on_show() haalt de nieuwste gegevens op.
        for page in self.pages.values():
            page.pack_forget()
        self.pages[name].pack(fill="both", expand=True)
        self.pages[name].on_show()
        self.navbar.highlight(name)
