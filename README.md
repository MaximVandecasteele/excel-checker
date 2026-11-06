# Excel Checker

Een lokale desktopapp voor het vergelijken van een modeloplossing met `.xlsx`-bestanden van leerlingen. De app maakt voor elk leerlingbestand een afzonderlijk HTML-rapport dat je in een browser kunt openen.

## Starten

Vereist Python 3.10 of hoger.

```bash
python3 -m pip install -r requirements.txt
python3 app.py
```

Kies in de app de modeloplossing, de map met leerlingbestanden en de map voor de rapporten. Alleen `.xlsx`-bestanden op het hoogste niveau van de gekozen leerlingenmap worden verwerkt. Excel-tijdelijke bestanden worden overgeslagen.

Selecteer eerst de modeloplossing; de GUI leest de werkbladen automatisch in. Stel vervolgens per werkblad en beoordelingsonderdeel het maximumaantal punten in. Lege velden worden niet beoordeeld en verschijnen niet in het rapport. Onderdelen zonder gevonden fouten krijgen automatisch de maximumscore. Voor onderdelen met fouten vul je de behaalde punten direct naast de titel in; een leeg veld telt als nul en de maxima kunnen niet worden gewijzigd. De totaalscore over alle werkbladen verschijnt onderaan.

Na het vergelijken gebruik je in dezelfde GUI de knop **Open rapporten en vul scores in**. Dat start een lokale rapportviewer in je browser. Boven- en onderaan elk rapport kun je naar de vorige of volgende toets gaan; de navigatie loopt door van de laatste naar de eerste en omgekeerd. Via **Overzicht** keer je terug naar de klaslijst. Een rapportlink licht groen op zodra alle scores zijn ingevuld. Bovenaan rechts staat de positie in de klas, bijvoorbeeld `2 / 7`. Ingevulde scores worden automatisch opgeslagen in de HTML-rapporten en als reserve in `.excel_checker_scores/` in de rapportmap. Met **Beëindig evaluatie** op het overzicht of in een rapport sluit je de rapportviewer; de knop is niet zichtbaar in afdrukken/PDF's. De server is alleen vanaf deze computer bereikbaar.

Klik op **Exporteer rapporten naar PDF** om de HTML-rapporten in de gekozen rapportmap naar aparte PDF-bestanden om te zetten. Kies de uitvoermap wanneer daarom wordt gevraagd. Dit vereist een eenmalige installatie van Chromium:

```bash
python3 -m playwright install chromium
```

Playwright gebruikt Chromium om rapporten lokaal te renderen; er is geen internetverbinding nodig tijdens het omzetten. De navigatiebediening van de rapportviewer wordt niet in PDF-bestanden afgedrukt.

## Tests uitvoeren

```bash
python3 -m pip install -r requirements.txt
python3 -m unittest -v
```

GitHub Actions voert de tests automatisch uit voor Python 3.10 tot en met 3.13 bij elke push en pull request.

## Wat wordt vergeleken?

De eerste versie vergelijkt celwaarden en formules, celopmaak, werkbladnamen en -volgorde, zichtbaarheid, samengevoegde cellen, rij- en kolominstellingen, tabellen, benoemde bereiken, gegevensvalidatie, voorwaardelijke opmaak en grafieken. Aaneengesloten formule- en opmaakverschillen met hetzelfde patroon in een rij of kolom worden als één melding getoond; identieke opmaakverschillen in meerdere losse bereiken worden ook samengevoegd. Grote aaneengesloten gebieden met celverschillen worden samengevat met een bereik, het aantal afwijkende cellen en enkele voorbeelden. Rij- en kolominstellingen worden eveneens compact per bereik weergegeven. Voorwaardelijke opmaak wordt in het rapport samengevat met bereik, voorwaarde en relevante opmaak; interne Excel-regelnummers worden genegeerd. Grafieken worden op constructie vergeleken, zoals type, titel, reeksen en gegevensbereiken. De celwaarden/caches waarop een grafiek gebaseerd is, worden niet als grafiekverschil gezien. Grafiekverschillen worden kort en leesbaar beschreven; de positie op het werkblad wordt genegeerd.

De vergelijking van celinhoud en opmaak is exact, zonder automatische beoordeling van gelijkwaardige formules. Rijhoogtes en kolombreedtes worden alleen als afwijkend gemeld wanneer ze meer dan 20% afwijken van het model. Rijen of kolommen met dezelfde afmeting in het model moeten onderling ook dezelfde afmeting hebben in het leerlingbestand. Overige rij- en kolominstellingen worden exact vergeleken. Autofilters worden op bereik en criteria gecontroleerd en als één leesbare melding weergegeven. Als het model een sorteerinstelling bevat, wordt de data die aan de model-filtercriteria voldoet op de verwachte sorteervolgorde gecontroleerd; een foutieve sortering geeft één melding. Wanneer de records verder identiek zijn maar anders gerangschikt, worden de daardoor ontstane celverschillen niet afzonderlijk getoond. Kolombreedtes die Excel als één bereik opslaat, worden per kolom vergeleken en verschillen worden per betrokken kolom gerapporteerd. Onbelangrijke Excel-verschillen in de schrijfwijze van getalnotaties (zoals een al dan niet ge-escapete spatie) worden genormaliseerd; inhoudelijk verschillende notaties blijven zichtbaar. Excel-functies die `openpyxl` niet volledig kan inlezen of bewaren, zoals sommige slicers, externe koppelingen of ingesloten objecten, vallen buiten deze eerste versie. Gebruik de app daarom op kopieën als bestanden zulke functies bevatten.

Leerlingbestanden (`.xlsx`/`.xlsm`), PDF's en rapportmappen zijn standaard uitgesloten van Git via `.gitignore`, zodat leerlinggegevens niet per ongeluk worden gepubliceerd. Dit project heeft momenteel nog geen open-sourcelicentie.
