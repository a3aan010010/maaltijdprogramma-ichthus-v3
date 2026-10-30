# Maaltijd en soos

Programma voor Ichthus Utrecht om aanmeldingen voor de maaltijd en het drankverbruik in de soos bij te houden, en beide naar Excel te exporteren.

| | |
|---|---|
| **Gemaakt door** | Adriaan Abrahamse |
| **Datum** | 30 oktober 2026 |
| **Versie** | 1.0 |
| **Taal** | Python 3, alleen standaard libraries |
| **Bestand** | `maaltijd.py` (alles zit in dit ene bestand) |

Gebouwd met hulp van Claude Code (Anthropic).

---

## Inhoud

1. [Wat kan het?](#wat-kan-het)
2. [Installeren en starten](#installeren-en-starten)
3. [Eerste keer instellen](#eerste-keer-instellen)
4. [Gebruik](#gebruik)
5. [Bestanden en opslag](#bestanden-en-opslag)
6. [Met meerdere computers tegelijk](#met-meerdere-computers-tegelijk)
7. [Problemen oplossen](#problemen-oplossen)
8. [Privacy](#privacy)
9. [Voor ontwikkelaars](#voor-ontwikkelaars)
10. [Wijzigingen](#wijzigingen)

---

## Wat kan het?

- **Maaltijd:** leden melden zich aan met hun lidnummer of naam. Iemand kan vaker aangemeld worden (bijvoorbeeld met gasten) en weer verwijderd worden.
- **Soos:** per lid drankjes boeken met Erbij en Eraf.
- **Drankjes:** zelf drankjes toevoegen, prijzen wijzigen en drankjes verwijderen, op een aparte pagina.
- **Totalen:** overzicht van de avond (aantal eters en alle soosrekeningen), op een aparte pagina zodat niet iedereen aan de bar alles ziet.
- **Excel-export:** twee losse bestanden, `Maaltijd <datum>.xlsx` en `Soos <datum>.xlsx`.
- **Archief:** bij "Nieuwe maaltijd" of "Nieuwe soos" wordt alles eerst geëxporteerd en bewaard.
- **Datum:** standaard automatisch vandaag, en los aan te passen.
- **Tegelijk werken:** maaltijd en soos kunnen tegelijk gebruikt worden, ook op meerdere computers.

## Installeren en starten

1. Installeer Python 3.8 of nieuwer van [python.org](https://www.python.org/downloads/). Vink bij de installatie op Windows **"Add Python to PATH"** aan. Verder is er niets nodig; ook Excel hoeft niet op de computer te staan.
2. Zet `maaltijd.py` in een eigen map.
3. Start het programma:
   - dubbelklik op `maaltijd.py`, of
   - open een opdrachtprompt in die map en typ `python maaltijd.py`.

Het programma opent schermvullend. Voor een gewoon venster:

```
python maaltijd.py --venster
```

Of zet in `maaltijd_config.json` de regel `"start_maximized": false`.

## Eerste keer instellen

Het programma heeft een Excel-bestand met de leden nodig.

**Standaardplek:** `maaltijd_data\leden.xlsx` naast `maaltijd.py`. Staat het daar, dan werkt alles meteen. Een ander bestand of een andere map kies je via **Instellingen > Opslag en ledenlijst**.

**Indeling van het ledenbestand:** het eerste werkblad wordt gelezen. De eerste rij mag kopjes hebben; deze worden herkend:

| Kolom | Herkende kopjes |
|---|---|
| Lidnummer | `Lidnummer`, `Nummer`, `Nr`, `Id` (alles met "lid" of "nummer") |
| Naam | `Naam`, of los `Voornaam` en `Achternaam` |

Zonder kopjes geldt: **kolom A = lidnummer, kolom B = naam**.

Lege regels en dubbele lidnummers worden overgeslagen. Lidnummers met letters of voorloopnullen (`A12`, `007`) werken ook. Sla het bestand op als gewone Excel-werkmap (`.xlsx`).

Na een wijziging in het ledenbestand: **Instellingen > Opslag en ledenlijst > Ledenlijst opnieuw inlezen**.

## Gebruik

### Navigatie

Bovenin staat de balk met **Maaltijd**, **Soos** en **Instellingen**. Instellingen is een uitklapmenu met de beheerpagina's: **Drankjes en prijzen**, **Totalen van de avond** en **Opslag en ledenlijst**. Rechts in de balk staat hoeveel leden er zijn ingelezen.

### Maaltijd

1. Typ een lidnummer of (een deel van) een naam. De lijst toont de treffers.
2. **Klik één keer** op de naam in de lijst. Dat meldt die persoon meteen aan. Er is ook maar één treffer? Dan werkt **Enter** (of de knop **Aanmelden**) ook.
3. Staat iemand er al op, dan vraagt het programma eerst of het nog een keer moet (bijvoorbeeld voor een gast). Per ongeluk aangemeld? Selecteer de naam rechts en klik **-1** of **Verwijderen**.

In de lijst rechts: selecteer een naam en gebruik **+1**, **-1** of **Verwijderen**.

- **Exporteer lijst naar Excel:** maakt of vervangt `Maaltijd <datum>.xlsx`.
- **Nieuwe maaltijd:** exporteert, bewaart een kopie in het archief en maakt de lijst leeg. De soos blijft staan.

### Soos

1. Typ links een lidnummer of naam en **klik één keer** op de naam in de lijst.
2. Druk bij het drankje op **+** of **-**. Elk drankje is een eenvoudige rij:

   ```
   Bier    1,50    [-]  [+]    2
   ```

   (naam, prijs, min, plus, aantal). Het aantal is vet en zwart als dit lid er iets van geboekt heeft.

De rijen worden automatisch over 1 tot 4 kolommen verdeeld, zodat ze zonder scrollen in beeld passen. 30 drankjes passen in een normaal venster in twee kolommen; in een kleiner venster of bij nog meer drankjes kun je scrollen met het muiswiel of de scrollbalk. Wil je de volgorde van de knoppen wisselen (bijvoorbeeld eerst plus), of de ruimte ertussen aanpassen, dan kan dat in `SoosPage._build_rows` en `SoosPage.BUTTON_GAP`.

Zolang er geen lid gekozen is, zijn de knoppen grijs. Min is ook grijs bij 0, want lager kan niet.

Je ziet alleen de rekening van het gekozen lid.

- **Exporteer lijst naar Excel:** maakt of vervangt `Soos <datum>.xlsx`.
- **Nieuwe soos:** exporteert, archiveert en maakt de rekeningen leeg. De drankjes en de maaltijd blijven staan.

### Instellingen > Drankjes en prijzen

- **Nieuw drankje:** naam en prijs (bijvoorbeeld `1,50`) en dan **Toevoegen**.
- **Prijs wijzigen:** klik een drankje aan, pas de prijs aan en klik **Prijs wijzigen**. Is het drankje die avond al geboekt, dan geldt de nieuwe prijs ook voor die boekingen; het programma vraagt daarom eerst om bevestiging.
- **Drankje verwijderen:** kan alleen als het nog nergens geboekt staat.

### Instellingen > Totalen van de avond

Toont het aantal aangemelde eters en alle soosrekeningen met het totaalbedrag.

### Datum

Maaltijd en soos hebben elk een eigen datumveld.

- Standaard is de datum **automatisch vandaag**.
- Typ een andere datum (`dd-mm-jjjj`) en klik **Datum opslaan** om hem vast te zetten.
- **Vandaag** zet hem terug op automatisch. "Nieuwe maaltijd" en "Nieuwe soos" doen dat ook.

Tip: loopt een avond door na middernacht, sla de datum dan eenmalig op, anders krijgt een export na middernacht de datum van de nieuwe dag.

## Bestanden en opslag

Naast `maaltijd.py`:

| Bestand | Inhoud |
|---|---|
| `maaltijd_config.json` | De instellingen (opslagmap, ledenbestand, schermvullend). |
| `ichthus-logo-gray-navbar.png` | Het logo in de navbar (216x40, al op maat). |
| `ichthus-logo-gray.png` | Het grote logo. Wordt gebruikt als de navbar-versie ontbreekt. |

Kopieer je het programma naar een andere computer, neem de logobestanden dan mee. Zonder logo werkt alles ook; dan staat er "Ichthus Utrecht" als tekst in de navbar.

In de opslagmap (standaard `maaltijd_data`):

| Bestand of map | Inhoud |
|---|---|
| `leden.xlsx` | Het ledenbestand (als je geen ander kiest). |
| `maaltijd.json` | De huidige maaltijd. |
| `soos.json` | De drankjes en de huidige soosrekeningen. |
| `Maaltijd <datum>.xlsx` | Export van de maaltijd. |
| `Soos <datum>.xlsx` | Export van de soos. |
| `archief\` | Een kopie van de gegevens bij elke "Nieuwe maaltijd" en "Nieuwe soos". |
| `*.lock` | Tijdelijk slotbestand tijdens het opslaan. Verdwijnt vanzelf. |

Elke klik wordt meteen opgeslagen. Het programma afsluiten of een stroomstoring kost dus geen gegevens.

**Een archief terugzetten:** sluit het programma, kopieer `maaltijd.json` of `soos.json` uit de archiefmap terug naar de opslagmap en start het programma weer.

## Met meerdere computers tegelijk

1. Zet de opslagmap op een gedeelde schijf of netwerkmap.
2. Start op elke computer het programma en kies bij **Instellingen > Opslag en ledenlijst** dezelfde opslagmap.

Elke wijziging wordt onder een slot opgeslagen en de schermen verversen zich elke 2 seconden, dus niemand overschrijft het werk van een ander.

Synchronisatiediensten zoals OneDrive of Dropbox zijn minder geschikt: die kunnen bij gelijktijdig werken kopieën met "conflict" in de naam maken. Een echte gedeelde netwerkmap werkt het best.

## Problemen oplossen

| Melding of probleem | Oplossing |
|---|---|
| "Het ledenbestand is niet gevonden" | Kies het juiste bestand bij Instellingen > Opslag en ledenlijst. |
| "Er zijn geen leden gevonden" | Controleer de kolommen (zie [Eerste keer instellen](#eerste-keer-instellen)). |
| "Dit is geen geldig .xlsx bestand" | Open het bestand in Excel en kies Opslaan als > Excel-werkmap (`.xlsx`). Oude `.xls` en `.csv` werken niet. |
| "Het bestand ... is geopend in Excel" | Sluit het exportbestand in Excel en klik nog eens op exporteren. |
| "De gegevens zijn bezig bij een ander scherm" | Even wachten en opnieuw proberen. Blijft het, controleer dan of de netwerkmap bereikbaar is. |
| Iemand staat niet in de zoeklijst | Staat die persoon in het ledenbestand? Klik daarna op Instellingen > Opslag en ledenlijst > Ledenlijst opnieuw inlezen. |
| Een bestand `maaltijd.json.kapot-...` | Het gegevensbestand was beschadigd en is apart gezet; het programma is schoon verdergegaan. Kijk in het archief voor de laatste goede versie. |
| Dubbelklikken op `maaltijd.py` doet niets | Start via de opdrachtprompt met `python maaltijd.py` om de foutmelding te zien. Meestal is Python niet (goed) geïnstalleerd. |

## Privacy

Het ledenbestand, de JSON-bestanden, de exports en het archief bevatten persoonsgegevens van leden. Daarom:

- Bewaar de opslagmap alleen op een plek waar bevoegde mensen bij kunnen.
- Zet de map `maaltijd_data` niet in een openbare Git-repository of gedeelde map.
- Gebruik voor testen een verzonnen ledenlijst.
- Ruim oude exports en archieven op als ze niet meer nodig zijn.

De soosrekeningen van alle leden staan bewust alleen op de pagina Totalen, niet op het scherm aan de bar.

## Voor ontwikkelaars

Alles staat in `maaltijd.py`, met uitgebreide uitleg in de code zelf. Lees eerst het blok bovenaan het bestand.

### Opbouw van het bestand

| Sectie | Inhoud |
|---|---|
| 1. Instellingen | `DEFAULT_CONFIG`, `load_config`, `save_config`, `StoreError` |
| 2. Excel | `read_xlsx_rows` en `write_xlsx`: zelfgebouwd met `zipfile` en XML |
| 3. Leden | `Member`, `load_members`, `find_members` |
| 4. Hulpjes | datum (`today_str`, `parse_datum`) en geld (`parse_prijs`, `euro`) |
| 5. Opslag | `FileLock`, `JsonStore`, het gegevensformaat, `auto_datum` |
| 6. App | hoofdvenster, export, archief, `show_page`, de decorator `@guarded` |
| 7. Schermen | `MemberPicker`, `BasePage`, `DatumField`, de pagina's en `Navbar` |

### Belangrijkste afspraken

- **Gegevens aanpassen gaat altijd via `JsonStore.update()`.** Die pakt een slot, leest de nieuwste versie, past aan en schrijft atomair weg. Nooit zelf lezen en wegschrijven, anders kunnen twee schermen elkaars werk overschrijven.
- **Geld in hele centen (int).** Nooit kommagetallen opslaan.
- **Datums als tekst `dd-mm-jjjj`.**
- **Lidnummers zijn altijd tekst.**
- **Fouten voor de gebruiker zijn `StoreError`** met een duidelijke Nederlandse tekst. Methodes achter een knop krijgen `@guarded`, dan verschijnt de fout als venstertje.
- **Niets in het geheugen dat niet ook op schijf staat.** Pagina's lezen de bestanden elke 2 seconden opnieuw (`BasePage._poll`).

### Een nieuwe pagina toevoegen

1. Maak een klasse die erft van `BasePage` en schrijf een `refresh()` die de gegevens leest en het scherm bijwerkt.
2. Roep aan het eind van `__init__` `self.refresh()` en `self.start_polling()` aan.
3. Voeg de pagina toe in `App.build()` aan `self.pages`.
4. Maak een knop in `Navbar.__init__` (met `_add_button`) of een item in `Navbar.SETTINGS_MENU`. Zet hem ook in `Navbar.GROUP`, zodat de juiste knop blauw kleurt.

### Een nieuwe instelling toevoegen

Zet hem met een standaardwaarde in `DEFAULT_CONFIG`. Bestaande `maaltijd_config.json`-bestanden blijven dan gewoon werken.

### Een veld toevoegen aan maaltijd.json of soos.json

Voeg het toe in `meal_default()` of `soos_default()` en werk de beschrijving van het gegevensformaat bij (sectie 5). Bestaande bestanden missen het veld nog; gebruik `data.get("veld", standaard)` of vul het aan in een upgrade-functie zoals `auto_datum`.

### Testen

Er zijn geen externe testtools nodig. Een snelle controle of het bestand nog klopt:

```
python -m py_compile maaltijd.py
```

Test daarna met een **verzonnen** ledenlijst en een aparte opslagmap (kies die bij Instellingen > Opslag en ledenlijst), zodat je de echte gegevens niet aanraakt. Loop minimaal langs: aanmelden, +1/-1/verwijderen, drankje toevoegen en boeken, prijs wijzigen, beide exports openen in Excel, Nieuwe maaltijd en Nieuwe soos.

## Wijzigingen

| Versie | Datum | Wijziging |
|---|---|---|
| 1.0 | 30-10-2026 | Eerste versie: maaltijd, soos, drankjes, totalen, instellingen, Excel-export en archief. |
