"""
De onderdelen van het maaltijd- en soosprogramma. Starten gaat via maaltijd.py
in de map hierboven; lees daar ook eerst de uitleg bovenaan.

Elk bestand is een onderdeel. Bestanden met een lager nummer gebruiken nooit
iets uit een bestand met een hoger nummer:

  1. instellingen.py  config, vaste paden en StoreError
  2. excel.py         xlsx lezen en schrijven
  3. leden.py         ledenlijst inlezen en zoeken
  4. hulpjes.py       geld en datum
  5. opslag.py        FileLock, JsonStore en het gegevensformaat
  6. app.py           het hoofdvenster, export en archief
     dialogen.py      vraagvensters en @guarded
  7. bouwstenen.py    stijl, MemberPicker, BasePage, DatumField
     paginas.py       de pagina's
     navbar.py        de balk bovenin

1 tot en met 5 gebruiken geen tkinter; die zijn zonder scherm te testen (zie tests/).
"""

__version__ = "1.1"
