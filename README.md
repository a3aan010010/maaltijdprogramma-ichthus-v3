# Maaltijd en soos

Programma voor Ichthus Utrecht om aanmeldingen voor de maaltijd en het drankverbruik in de soos bij te houden, en beide naar Excel te exporteren.

| | |
|---|---|
| **Gemaakt door** | Adriaan Abrahamse |
| **Datum** | 5 november 2026 |
| **Versie** | 1.2 |
| **Taal** | Python 3, alleen standaard libraries |
| **Starten** | `maaltijd.py` (de onderdelen staan in de map `maaltijd_app`) |
| **Licentie** | Apache 2.0 (zie `LICENSE`) |

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
- **Archief:** bij "Nieuwe maaltijd" of "Nieuwe soos" wordt vanzelf (zonder melding) een kopie van de gegevens bewaard. Daarnaast wordt standaard een Excel gemaakt; dat kun je in het bevestigingsvenster uitzetten met het vinkje "Excel aanmaken".
- **Datum:** standaard automatisch: vandaag, en vanaf de eerste aanmelding of boeking ligt hij vast (ook als de avond na middernacht doorloopt). Los aan te passen.
- **Tegelijk werken:** maaltijd en soos kunnen tegelijk gebruikt worden, ook op meerdere computers.

## Installeren en starten

1. Installeer Python 3.8 of nieuwer van [python.org](https://www.python.org/downloads/). Vink bij de installatie op Windows **"Add Python to PATH"** aan. Verder is er niets nodig; ook Excel hoeft niet op de computer te staan.
2. Zet `maaltijd.py`, de map `maaltijd_app`, de logobestanden, `ichthus-visje.ico` en `soos_standaard.json` samen in een eigen map. `maaltijd_app` moet altijd naast `maaltijd.py` staan, anders start het programma niet.
3. Start het programma:
   - dubbelklik op `maaltijd.py`, of
   - open een opdrachtprompt in die map en typ `python maaltijd.py`.

Het programma opent schermvullend. De versie staat in de titelbalk. Voor een gewoon venster:

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

- **Exporteer lijst naar Excel:** maakt `Maaltijd <datum>.xlsx`, of vervangt de vorige export van deze lijst. Een export van een andere lijst met dezelfde datum wordt nooit overschreven; dan komt er `Maaltijd <datum> (2).xlsx` naast.
- **Nieuwe maaltijd:** bewaart stil een kopie in het archief en maakt de lijst leeg. De soos blijft staan. In het bevestigingsvenster staat het vinkje **Excel aanmaken** standaard aan: dan wordt eerst `Maaltijd <datum>.xlsx` gemaakt (lukt dat niet, dan wordt er niets gewist). Zet je het vinkje uit, dan blijft het bij de back-up.

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

- **Exporteer lijst naar Excel:** maakt `Soos <datum>.xlsx`, of vervangt de vorige export van deze soos. Net als bij de maaltijd wordt de export van een andere soos nooit overschreven (dan `Soos <datum> (2).xlsx`).
- **Nieuwe soos:** bewaart stil een kopie in het archief en maakt de rekeningen leeg. De drankjes en de maaltijd blijven staan. Ook hier staat het vinkje **Excel aanmaken** standaard aan: dan wordt eerst `Soos <datum>.xlsx` gemaakt. Zet je het uit, dan blijft het bij de back-up.

### Instellingen > Drankjes en prijzen

- **Nieuw drankje:** naam en prijs (bijvoorbeeld `1,50`) en dan **Toevoegen**.
- **Prijs wijzigen:** klik een drankje aan, pas de prijs aan en klik **Prijs wijzigen**. Is het drankje die avond al geboekt, dan geldt de nieuwe prijs ook voor die boekingen; het programma vraagt daarom eerst om bevestiging.
- **Drankje verwijderen:** kan alleen als het nog nergens geboekt staat.

### Instellingen > Totalen van de avond

Toont het aantal aangemelde eters en alle soosrekeningen met het totaalbedrag.

### Datum

Maaltijd en soos hebben elk een eigen datumveld.

- Standaard is de datum **automatisch**: zolang de lijst leeg is, is het vandaag. Bij de eerste aanmelding of boeking ligt de datum vast. Loopt de avond door na middernacht, dan houden de lijst, de export en het archief dus gewoon de datum van de avond.
- Staat er na een paar dagen nog een oude datum ("automatisch: de dag waarop deze lijst begon"), dan is "Nieuwe maaltijd" of "Nieuwe soos" vergeten.
- Typ een andere datum (`dd-mm-jjjj`) en klik **Datum opslaan** om hem zelf in te stellen.
- **Vandaag** zet de datum op vandaag en weer op automatisch. "Nieuwe maaltijd" en "Nieuwe soos" maken de lijst leeg; daarna volgt de datum weer vandaag.

## Bestanden en opslag

Naast `maaltijd.py`:

| Bestand | Inhoud |
|---|---|
| `maaltijd_config.json` | De instellingen (opslagmap, ledenbestand, schermvullend). |
| `maaltijd_fout.log` | Details van onverwachte fouten. Ontstaat pas bij de eerste fout; handig om mee te sturen als je hulp vraagt. |
| `ichthus-logo-gray-navbar.png` | Het logo in de navbar (216x40, al op maat). |
| `ichthus-logo-gray.png` | Het grote logo. Wordt gebruikt als de navbar-versie ontbreekt. |
| `ichthus-visje.ico` | Het visje linksboven in het venster en op de taakbalkknop, in alle maten van 16 tot 256 pixels. |
| `soos_standaard.json` | De vaste drankjes en prijzen waarmee een nieuwe installatie begint. Wordt alleen gelezen zolang er nog geen `soos.json` in de opslagmap staat; daarna past het programma alleen `soos.json` aan. Hierin staan nooit boekingen. |

Kopieer je het programma naar een andere computer, neem dan `maaltijd.py`, de map `maaltijd_app`, de logobestanden, `ichthus-visje.ico` en `soos_standaard.json` mee. Zonder logo werkt alles ook; dan staat er "Ichthus Utrecht" als tekst in de navbar, en zonder visje het standaardicoon van tkinter.

In de opslagmap (standaard `maaltijd_data`):

| Bestand of map | Inhoud |
|---|---|
| `leden.xlsx` | Het ledenbestand (als je geen ander kiest). |
| `maaltijd.json` | De huidige maaltijd. |
| `soos.json` | De drankjes en de huidige soosrekeningen. |
| `maaltijd.json.reserve`, `soos.json.reserve` | Reservekopie met dezelfde inhoud. Wordt gebruikt als het echte bestand beschadigd raakt. |
| `*.kapot-...` | Een beschadigd gegevensbestand, bewaard na herstel uit de reservekopie. Mag weg als alles klopt. |
| `Maaltijd <datum>.xlsx` | Export van de maaltijd (`(2)`, `(3)` als er meer lijsten met die datum zijn). |
| `Soos <datum>.xlsx` | Export van de soos (idem). |
| `archief\` | Een kopie van de gegevens bij elke "Nieuwe maaltijd" en "Nieuwe soos". |
| `*.lock` | Tijdelijk slotbestand tijdens het opslaan. Verdwijnt vanzelf. |

Elke klik wordt meteen opgeslagen, met een reservekopie ernaast. Het programma afsluiten of een stroomstoring kost dus geen gegevens.

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
| "... was beschadigd en is hersteld uit de reservekopie" | Het programma heeft de reservekopie teruggezet. Controleer de laatste boekingen en doe je laatste klik opnieuw. Het beschadigde bestand staat ernaast als `*.kapot-...`. |
| "... is beschadigd en er is geen bruikbare reservekopie" | Er wordt niets opgeslagen tot dit is opgelost. Sluit het programma en zet het bestand terug uit `archief\` (zie [Bestanden en opslag](#bestanden-en-opslag)), of verwijder het om leeg te beginnen. |
| "De gegevens konden niet worden opgeslagen: ... Access is denied" | Een ander programma (virusscanner, back-up, of het programma een tweede keer geopend) hield het bestand langer dan een seconde vast. Sluit een eventueel tweede venster en doe de klik opnieuw. |
| "Onverwachte fout" | De gegevens zijn veilig. De details staan in `maaltijd_fout.log` naast `maaltijd.py`. |
| Dubbelklikken op `maaltijd.py` doet niets | Start via de opdrachtprompt met `python maaltijd.py` om de foutmelding te zien. Meestal is Python niet (goed) geïnstalleerd. |

## Privacy

Het ledenbestand, de JSON-bestanden, de exports en het archief bevatten persoonsgegevens van leden. Daarom:

- Bewaar de opslagmap alleen op een plek waar bevoegde mensen bij kunnen.
- Zet de map `maaltijd_data` niet in een openbare Git-repository of gedeelde map.
- Gebruik voor testen een verzonnen ledenlijst.
- Ruim oude exports en archieven op als ze niet meer nodig zijn.

De soosrekeningen van alle leden staan bewust alleen op de pagina Totalen, niet op het scherm aan de bar.

## Voor ontwikkelaars

`maaltijd.py` start alleen het programma. Alle onderdelen staan in de map `maaltijd_app`, met uitgebreide uitleg in de code zelf. Lees eerst het blok bovenaan `maaltijd.py` en daarna `maaltijd_app/__init__.py`.

### Opbouw

| Bestand in `maaltijd_app` | Inhoud |
|---|---|
| `__init__.py` | Het versienummer (`__version__`) en een overzicht van de bestanden |
| `instellingen.py` | `APP_DIR`, `DEFAULT_CONFIG`, `load_config`, `save_config`, `StoreError` |
| `excel.py` | `read_xlsx_rows` en `write_xlsx`: zelfgebouwd met `zipfile` en XML |
| `leden.py` | `Member`, `load_members`, `find_members` |
| `hulpjes.py` | datum (`today_str`, `parse_datum`) en geld (`parse_prijs`, `euro`) |
| `opslag.py` | `FileLock`, `JsonStore`, het gegevensformaat, `auto_datum` |
| `app.py` | hoofdvenster, export, archief, `show_page`, foutlog |
| `dialogen.py` | `ModalDialog`, `YesNoDialog`, `NewListDialog`, de decorator `@guarded` |
| `bouwstenen.py` | `setup_style`, `MemberPicker`, `BasePage`, `DatumField` |
| `paginas.py` | de pagina's: Maaltijd, Soos, Drankjes, Totalen, Instellingen |
| `navbar.py` | de balk bovenin |

De bestanden van `instellingen.py` tot en met `opslag.py` gebruiken geen tkinter en nooit iets uit de schermbestanden. Houd dat zo: dan blijven ze zonder scherm te testen.

### Een nieuwe versie uitbrengen

1. Zet het nieuwe nummer in `maaltijd_app/__init__.py` (`__version__`).
2. Werk het blok bovenaan `maaltijd.py`, de tabel bovenaan deze README, de tabel [Wijzigingen](#wijzigingen) en `SECURITY.md` bij.
3. Commit en zet er een tag op, bijvoorbeeld `git tag -a v1.2 -m "Versie 1.2"`, en push met `git push --follow-tags`.

### Belangrijkste afspraken

- **Gegevens aanpassen gaat altijd via `JsonStore.update()`.** Die pakt een slot, leest de nieuwste versie, past aan en schrijft atomair weg. Nooit zelf lezen en wegschrijven, anders kunnen twee schermen elkaars werk overschrijven.
- **Geld in hele centen (int).** Nooit kommagetallen opslaan.
- **Datums als tekst `dd-mm-jjjj`.**
- **Lidnummers zijn altijd tekst.**
- **Fouten voor de gebruiker zijn `StoreError`** met een duidelijke Nederlandse tekst. Methodes achter een knop krijgen `@guarded`, dan verschijnt de fout als venstertje.
- **Niets in het geheugen dat niet ook op schijf staat.** De pagina die in beeld is leest de bestanden elke 2 seconden opnieuw (`BasePage._poll`); een verborgen pagina leest pas weer als hij in beeld komt (`on_show`).
- **Eigen vraagvensters erven van `ModalDialog`.** "Nieuwe ..."-knoppen gebruiken `BasePage.start_new_list`, exportknoppen `BasePage.show_export`.

### Een nieuwe pagina toevoegen

1. Maak in `paginas.py` een klasse die erft van `BasePage` en schrijf een `refresh()` die de gegevens leest en het scherm bijwerkt.
2. Roep aan het eind van `__init__` `self.start_polling()` aan. Roep daar niet zelf `self.refresh()` aan: de eerste keer lezen gebeurt in `on_show`, waar een melding over bijvoorbeeld een beschadigd bestand netjes wordt getoond in plaats van het opstarten af te breken.
3. Voeg de pagina toe in `App.build()` (`app.py`) aan `self.pages`.
4. Maak een knop in `Navbar.__init__` (met `_add_button`) of een item in `Navbar.SETTINGS_MENU`. Zet hem ook in `Navbar.GROUP`, zodat de juiste knop blauw kleurt.

### Een nieuwe instelling toevoegen

Zet hem met een standaardwaarde in `DEFAULT_CONFIG`. Bestaande `maaltijd_config.json`-bestanden blijven dan gewoon werken.

### Een veld toevoegen aan maaltijd.json of soos.json

Voeg het toe in `meal_default()` of `soos_default()` en werk de beschrijving van het gegevensformaat bij (in `opslag.py`). Bestaande bestanden missen het veld nog; gebruik `data.get("veld", standaard)` of vul het aan in een upgrade-functie zoals `auto_datum`.

### Testen

Er zijn geen externe testtools nodig. In de map `tests` staan tests voor alles wat geen scherm nodig heeft (geld, datum, Excel, ledenlijst, opslag). Draaien vanuit de map van het programma:

```
python -m unittest discover -s tests -v
```

Gebruik in tests alleen verzonnen namen.

Test daarna ook met de hand, met een **verzonnen** ledenlijst en een aparte opslagmap (kies die bij Instellingen > Opslag en ledenlijst), zodat je de echte gegevens niet aanraakt. Loop minimaal langs: aanmelden, +1/-1/verwijderen, drankje toevoegen en boeken, prijs wijzigen, beide exports openen in Excel, Nieuwe maaltijd en Nieuwe soos.

## Wijzigingen

| Versie | Datum | Wijziging |
|---|---|---|
| 1.2 | 05-11-2026 | Geen gegevensverlies meer bij een avond die na middernacht doorloopt: de datum ligt vast vanaf de eerste aanmelding of boeking. Een export overschrijft nooit meer de export van een andere lijst. Elke opslag krijgt een reservekopie; een beschadigd bestand wordt daaruit hersteld in plaats van stil leeg verder te gaan. Opslaan probeert het opnieuw als een ander programma het bestand even vasthoudt. Het Ichthus-visje als icoon van het venster en de taakbalkknop, in plaats van de tkinter-veer. Een nieuwe installatie begint met de vaste drankjes en prijzen uit `soos_standaard.json`. |
| 1.1 | 02-11-2026 | Code opgesplitst in de map `maaltijd_app`. Tests in `tests/`, verplicht te maken met een pre-commit hook en GitHub Actions. Onverwachte fouten in `maaltijd_fout.log`. Opslaan is bestand tegen stroomstoringen (`fsync`). Het scherm wacht maximaal 3 seconden op een bezet slot (was 10). Een pagina blijft verversen na een fout, en alleen de zichtbare pagina leest de bestanden. Versie in de titelbalk. Minder dubbele code. |
| 1.0 | 30-10-2026 | Eerste versie: maaltijd, soos, drankjes, totalen, instellingen, Excel-export en archief. |
