"""
Vraagvensters (ja/nee, Nieuwe maaltijd/soos) en de decorator @guarded.
"""

import functools
import tkinter as tk
from tkinter import messagebox, ttk

from .instellingen import StoreError


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


class ModalDialog(tk.Toplevel):
    """
    Basis voor een eigen vraagvenster: titel, tekst, midden boven het hoofdvenster,
    en zolang het open is kan er niets anders worden aangeklikt.

    Een kindklasse roept eerst super().__init__(parent, titel, tekst) aan, zet
    daarna haar eigen knoppen in self.frame en roept tot slot self.run() aan.
    Escape of het kruisje sluit het venster zonder iets te kiezen.
    """

    def __init__(self, parent, title, text):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.frame = ttk.Frame(self, padding=18)
        self.frame.pack()
        ttk.Label(self.frame, text=text, justify="left", wraplength=420).pack(anchor="w")
        self.bind("<Escape>", lambda e: self.destroy())

    def run(self, focus=None):
        """Zet het venster op zijn plek, geeft `focus` de focus en wacht tot het dicht is."""
        parent = self.master
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.grab_set()  # zolang dit venster open is, kan er niets anders worden aangeklikt
        if focus is not None:
            focus.focus_set()
        self.wait_window()  # wacht hier tot het venster dicht is


class YesNoDialog(ModalDialog):
    """
    Ja/nee-vraag met Nederlandse knoppen. De standaard tkinter-vraag (messagebox.askyesno)
    krijgt zijn knopteksten van Windows en toont dus "Yes" en "No" als Windows op Engels
    staat; daarom hebben we een eigen venster.

    De knop Nee heeft de focus, zodat Enter uit gewoonte nooit iets wist of aanmeldt.
    Escape of het kruisje betekent ook Nee. Na afloop staat in `result` True (ja) of False (nee).
    Gebruik liever de functie ask_yes_no hieronder.
    """

    def __init__(self, parent, title, text):
        super().__init__(parent, title, text)
        self.result = False
        buttons = ttk.Frame(self.frame)
        buttons.pack(anchor="e", pady=(16, 0))
        self.no_button = ttk.Button(buttons, text="Nee", style="Big.TButton", command=self.destroy)
        self.no_button.pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="Ja", style="Big.TButton", command=self._yes).pack(side="right")
        # Enter drukt op de knop die nu de focus heeft (Tab wisselt tussen Ja en Nee).
        self.bind("<Return>", lambda e: self.focus_get().invoke() if isinstance(self.focus_get(), ttk.Button) else None)
        # Krijgt het venster zelf de focus (bijv. doordat Windows het naar voren haalt),
        # zet die dan terug op Nee, zodat Enter nooit per ongeluk Ja betekent.
        self.bind("<FocusIn>", lambda e: self.no_button.focus_set() if e.widget is self else None)
        self.run(focus=self.no_button)

    def _yes(self):
        self.result = True
        self.destroy()


def ask_yes_no(title, message, parent):
    """Stelt een ja/nee-vraag met Nederlandse knoppen. True = Ja, False = Nee."""
    return YesNoDialog(parent.winfo_toplevel(), title, message).result


class NewListDialog(ModalDialog):
    """
    Bevestigingsvenster voor "Nieuwe maaltijd" en "Nieuwe soos", met een vinkje
    "Excel aanmaken" dat standaard aan staat.

    Na afloop staat in `result`:
        None   als er is geannuleerd
        True   als er doorgegaan wordt met een Excel (vinkje aan)
        False  als er doorgegaan wordt zonder Excel (vinkje uit)
    Gebruik: NewListDialog(parent, titel, tekst, knoptekst).result
    """

    def __init__(self, parent, title, text, ok_text):
        super().__init__(parent, title, text)
        self.result = None
        # Standaard aan: er wordt eerst een Excel gemaakt (en gaat er niets leeg als dat
        # mislukt). Haal je het vinkje weg, dan blijft het bij de stille back-up.
        self.make_excel = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.frame, text="Excel aanmaken", variable=self.make_excel).pack(anchor="w", pady=(14, 14))
        buttons = ttk.Frame(self.frame)
        buttons.pack(anchor="e")
        ttk.Button(buttons, text="Annuleren", command=self.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(buttons, text=ok_text, style="Big.TButton", command=self._ok).pack(side="right")
        self.run()

    def _ok(self):
        self.result = self.make_excel.get()
        self.destroy()
