"""
7a. Bouwstenen voor de schermen: stijl, MemberPicker, BasePage en DatumField.
"""

import os
import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, ttk

from .dialogen import NewListDialog, guarded
from .hulpjes import parse_datum, today_str
from .instellingen import LOG_PATH, StoreError
from .leden import find_members
from .opslag import set_auto_datum, set_manual_datum


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
        self.entry.bind("<Return>", self.submit)
        self.entry.bind("<Down>", self._to_list)

        # exportselection=False: anders verdwijnt de selectie zodra je ergens anders klikt.
        self.list = tk.Listbox(self, height=8, exportselection=False, font=("Segoe UI", 12))
        self.list.pack(fill="both", expand=True)
        self.list.bind("<Return>", self.submit)
        if single_click:
            # Bij het LOSLATEN van de muisknop kiezen, niet bij pijltjestoetsen:
            # anders zou je met pijltje omlaag al een lid kiezen terwijl je nog zoekt.
            self.list.bind("<ButtonRelease-1>", self._click)
        else:
            self.list.bind("<Double-Button-1>", self.submit)

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

    def submit(self, _event=None):
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
      * aan het eind van __init__ self.start_polling() aanroepen, maar NIET zelf
        self.refresh(): de eerste keer lezen gebeurt in on_show, als de pagina in
        beeld komt. Daar wordt een StoreError (bijvoorbeeld een beschadigd bestand)
        netjes als melding getoond; in __init__ zou die het opstarten afbreken.
    En mag:
      * focus_main() maken: waar de cursor staat als de pagina in beeld komt
      * self.message(tekst, error=True/False) gebruiken voor de statusregel

    Het truukje met self._sig ("signature"): refresh() maakt een JSON-tekst van
    de gegevens en vergelijkt die met de vorige keer. Is er niets veranderd, dan
    stopt refresh() meteen. Zo flikkert het scherm niet en blijft een selectie
    in een lijst staan. Wil je het scherm hoe dan ook opnieuw tekenen (bijv. na
    het kiezen van een ander lid), roep dan self._force_refresh() aan.
    """

    POLL_MS = 2000  # hoe vaak (in milliseconden) de zichtbare pagina de bestanden opnieuw leest

    def __init__(self, app, parent):
        super().__init__(parent)
        self.app = app
        self._sig = None  # zie de uitleg hierboven
        self._job = None  # id van de geplande volgende _poll, om te kunnen stoppen
        self._last_poll_error = None  # zodat dezelfde fout niet elke 2 seconden in het log komt
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
        """
        Leest elke paar seconden de gegevens opnieuw, zodat bijvoorbeeld de datum na
        middernacht vanzelf meeloopt. Alleen de pagina die in beeld is leest echt;
        een verborgen pagina wordt bijgewerkt zodra hij in beeld komt (on_show).
        """
        try:
            if self.winfo_ismapped():
                self.refresh()
            self._last_poll_error = None
        except StoreError as e:
            # Hier geen venstertje: dat zou elke 2 seconden terugkomen.
            self.message(str(e), error=True)
        except tk.TclError:
            return  # het venster wordt afgesloten; stoppen met verversen
        except Exception as e:
            # Onverwachte fout: wel melden en loggen, maar gewoon door blijven verversen.
            # Zonder deze regel zou de pagina na een fout nooit meer bijwerken.
            if repr(e) != self._last_poll_error:
                self._last_poll_error = repr(e)
                self.app.log_error(*sys.exc_info())
            self.message(f"Er ging iets mis bij het verversen. Zie {os.path.basename(LOG_PATH)}.", error=True)
        # after() plant de volgende keer in, zonder het programma te blokkeren.
        self._job = self.after(self.POLL_MS, self._poll)

    def _on_destroy(self, event):
        """Stopt de geplande verversing als de pagina wordt opgeruimd."""
        if event.widget is self and self._job:
            self.after_cancel(self._job)

    def refresh(self):  # wordt door de kindklassen ingevuld
        raise NotImplementedError

    def _force_refresh(self):
        """Scherm hoe dan ook opnieuw tekenen (zie de uitleg over _sig hierboven)."""
        self._sig = None
        self.refresh()

    def show_export(self, path):
        """Meldt een geslaagde export, in de statusregel en in een venstertje."""
        self.message("Export gelukt.")
        messagebox.showinfo("Export gelukt", f"Het bestand is opgeslagen als:\n{path}", parent=self)

    def start_new_list(self, title, text, export, archive_and_reset, done_text):
        """
        Gedeelde werkwijze van "Nieuwe maaltijd" en "Nieuwe soos". Geeft True als er
        echt opnieuw begonnen is, False als er geannuleerd is.

        Eerst het bevestigingsvenster met het vinkje "Excel aanmaken". Staat dat aan,
        dan eerst export(); mislukt die (bijv. bestand open in Excel), dan stopt alles
        via een StoreError voordat er iets gewist wordt. Daarna archive_and_reset():
        een stille back-up in het archief (zie App._archive) en de lijst leegmaken.
        """
        dialog = NewListDialog(self.app.root, title, text, f"{title} starten")
        if dialog.result is None:
            return False
        path = export() if dialog.result else None
        archive_and_reset()
        self.message(done_text if path is None else f"{done_text} Excel opgeslagen als {os.path.basename(path)}.")
        return True


class DatumField(ttk.Frame):
    """
    Datumveld met knoppen. De datum is standaard automatisch: vandaag zolang de lijst
    leeg is, en daarna de dag waarop de lijst begon (ook na middernacht). Sla je een
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
        if data.get("datum_handmatig"):
            self.info.set("(zelf ingesteld)")
        elif data["datum"] == today_str():
            self.info.set("(automatisch: vandaag)")
        else:
            self.info.set("(automatisch: de dag waarop deze lijst begon)")

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
        """Zet de datum op vandaag en weer op automatisch (zie set_auto_datum)."""
        self.get_store().update(set_auto_datum())
        self.var.set(today_str())
        self.message(f"De datum is weer automatisch en staat op vandaag ({today_str()}).")
        self.on_change()
