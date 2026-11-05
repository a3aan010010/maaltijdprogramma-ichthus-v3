"""
7b. De pagina's: Maaltijd, Soos, Drankjes en prijzen, Totalen en Instellingen.
"""

import datetime
import json
import os
import tkinter as tk
from tkinter import filedialog, ttk

from .bouwstenen import BasePage, DatumField, MemberPicker
from .dialogen import ask_yes_no, guarded
from .hulpjes import euro, parse_prijs
from .instellingen import APP_DIR, StoreError, save_config
from .leden import load_members
from .opslag import member_total_cents


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
        ttk.Button(left, text="Aanmelden", style="Big.TButton", command=self.picker.submit).pack(fill="x", pady=(8, 0))

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

        self.start_polling()  # de eerste keer lezen gebeurt in on_show (zie BasePage)

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
        if current and not ask_yes_no(
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
        if not ask_yes_no("Verwijderen", f"{naam} helemaal van de lijst verwijderen?", parent=self):
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
        self.show_export(self.app.export_meal())

    @guarded
    def new_meal(self):
        """
        Knop Nieuwe maaltijd: eventueel eerst 'Maaltijd <datum>.xlsx', dan een stille
        back-up van maaltijd.json in het archief en de lijst leegmaken (zie start_new_list).
        """
        if self.start_new_list(
            "Nieuwe maaltijd",
            "De maaltijdlijst wordt leeggemaakt. Een kopie blijft bewaard in het archief. "
            "De soos blijft zoals die is.",
            self.app.export_meal,
            self.app.archive_and_reset_meal,
            "Nieuwe maaltijd gestart. Een kopie van de vorige staat in het archief.",
        ):
            self.refresh()


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

        self.start_polling()  # de eerste keer lezen gebeurt in on_show (zie BasePage)

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
        Knop Nieuwe soos: eventueel eerst 'Soos <datum>.xlsx', dan een stille back-up van
        soos.json in het archief en de rekeningen leegmaken (zie start_new_list).
        De drankjes en prijzen blijven staan, de maaltijd wordt niet aangeraakt.
        """
        if self.start_new_list(
            "Nieuwe soos",
            "De soos-rekeningen worden leeggemaakt. Een kopie blijft bewaard in het archief. "
            "De drankjes en de maaltijdlijst blijven zoals ze zijn.",
            self.app.export_soos,
            self.app.archive_and_reset_soos,
            "Nieuwe soos gestart. Een kopie van de vorige staat in het archief.",
        ):
            self.current = None
            self._force_refresh()

    @guarded
    def export(self):
        """Knop Exporteer: schrijft 'Soos <datum>.xlsx' in de opslagmap."""
        self.show_export(self.app.export_soos())


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

        self.start_polling()  # de eerste keer lezen gebeurt in on_show (zie BasePage)

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
        if any(e["drank"].get(drink_id) for e in soos["verbruik"].values()) and not ask_yes_no(
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
        if not ask_yes_no("Drankje verwijderen", f"Drankje '{name}' verwijderen?", parent=self):
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

        self.start_polling()  # de eerste keer lezen gebeurt in on_show (zie BasePage)

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
