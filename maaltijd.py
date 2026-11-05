"""
===============================================================================
  Maaltijd- en soosprogramma voor Ichthus Utrecht
===============================================================================

  Gemaakt door : Adriaan Abrahamse
  Datum        : 5 november 2026
  Versie       : 1.2
  Taal         : Python 3 (alleen standaard libraries)
  Licentie     : Apache 2.0 (zie LICENSE)

  Gebouwd met hulp van Claude Code (Anthropic).
  Zie README.md voor installatie, gebruik en uitleg voor ontwikkelaars.

  Wijzigingen
  -----------
  1.2  (05-11-2026)  Datum ligt vast vanaf de eerste aanmelding of boeking (geen
                     verspringende datum na middernacht); export overschrijft nooit
                     een andere lijst; reservekopie bij elke opslag en herstel van
                     een beschadigd bestand; opslaan probeert opnieuw bij een bezet bestand;
                     Ichthus-visje als venstericoon (ichthus-visje.ico); een nieuwe
                     installatie begint met de drankjes uit soos_standaard.json.
  1.1  (02-11-2026)  Opgesplitst in de map maaltijd_app; tests in tests/;
                     foutlog maaltijd_fout.log; fsync bij opslaan; scherm
                     bevriest korter bij een bezet slot; versie in de titel.
  1.0  (30-10-2026)  Eerste versie: maaltijd, soos, drankjes, totalen,
                     instellingen, Excel-export en archief.

===============================================================================

Alleen standaard Python libraries (tkinter, zipfile, json, ...), dus er hoeft niets
geinstalleerd te worden. Starten met:  python maaltijd.py
Voor een gewoon (niet schermvullend) venster:  python maaltijd.py --venster
Dit bestand start alleen het programma; alle onderdelen staan in de map
maaltijd_app, die altijd naast dit bestand moet staan.

Wat doet het?
  * Maaltijd: leden melden zich aan met lidnummer OF naam. Iemand mag vaker
    aangemeld worden (bijvoorbeeld met gasten) en kan weer verwijderd worden.
  * Soos: drankjes eenvoudig toevoegen, per lid met plus en min boeken.
  * Beide kunnen tegelijk draaien, ook op twee computers die dezelfde map delen.
    Elke wijziging wordt direct opgeslagen onder een slot (lock) en de schermen
    verversen zichzelf, dus niemand overschrijft elkaars werk.
  * Export naar Excel in twee losse bestanden: "Maaltijd <datum>.xlsx" (aanmeldlijst)
    en "Soos <datum>.xlsx" (rekeningen). Maaltijd en soos hebben elk hun eigen datum,
    exportknop en "Nieuwe ..." knop en raken elkaar nooit.
  * Datum: standaard automatisch die van vandaag, en vanaf de eerste aanmelding of
    boeking ligt hij vast (ook na middernacht). Los aan te passen; de knop "Vandaag"
    zet hem op vandaag en weer op automatisch.
  * Instellingen: de opslagmap en het ledenbestand (Excel) zijn instelbaar.

Opbouw: zie maaltijd_app/__init__.py voor welk bestand wat doet.

Hoe het in grote lijnen werkt (voor wie hieraan verder bouwt):
  * Er is een venster (App.root). Bovenin staat de Navbar, daaronder steeds een
    "pagina" (een ttk.Frame). App.show_page() wisselt van pagina.
  * Alle gegevens staan in twee JSON-bestanden in de opslagmap: maaltijd.json en
    soos.json. Er wordt NIETS in het geheugen bewaard dat niet ook op schijf staat.
    Elke klik doet: slot pakken, bestand lezen, aanpassen, wegschrijven, slot los
    (zie JsonStore.update in opslag.py). Daardoor kunnen meerdere schermen of
    computers veilig tegelijk werken.
  * De pagina die in beeld is, leest elke 2 seconden het bestand opnieuw
    (BasePage._poll) en tekent zich alleen opnieuw als er echt iets veranderd is.
  * Fouten die de gebruiker moet zien, zijn StoreError. De decorator @guarded
    vangt die op en toont een venstertje in plaats van te crashen. Alle andere
    (onverwachte) fouten komen met de volledige details in maaltijd_fout.log.
  * Geld wordt altijd in hele centen (int) opgeslagen, nooit als kommagetal,
    zodat er geen afrondfouten ontstaan.

Een nieuwe pagina toevoegen:
  1. Maak in paginas.py een klasse die erft van BasePage en implementeer refresh().
  2. Roep aan het eind van __init__ self.refresh() en self.start_polling() aan.
  3. Zet hem in App.build() (app.py) in self.pages en maak een knop of menu-item
     in Navbar (en zet hem in Navbar.GROUP, zodat de juiste knop blauw kleurt).
"""

from maaltijd_app.app import App
from maaltijd_app.bouwstenen import setup_style


def main():
    """Startpunt van het programma."""
    app = App()  # venster en instellingen
    setup_style(app.root)  # lettergroottes en stijlen (moet voor build)
    app.build()  # navbar en pagina's
    app.root.mainloop()  # wacht op klikken en toetsen tot het venster dicht gaat


if __name__ == "__main__":
    main()
