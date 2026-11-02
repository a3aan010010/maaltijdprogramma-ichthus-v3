"""
7c. De navbar: de balk bovenin die tussen de pagina's wisselt.
"""

import os
import tkinter as tk

from .instellingen import APP_DIR


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
