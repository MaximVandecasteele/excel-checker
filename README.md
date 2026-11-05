# Excel Checker

Een lokale desktopapp voor het vergelijken van een modeloplossing met `.xlsx`-bestanden van leerlingen. De app maakt voor elk leerlingbestand een afzonderlijk HTML-rapport dat je in een browser kunt openen.

## Starten

Vereist Python 3.10 of hoger.

```bash
python3 -m pip install -r requirements.txt
python3 app.py
```

Kies in de app de modeloplossing, de map met leerlingbestanden en de map voor de rapporten. Alleen `.xlsx`-bestanden op het hoogste niveau van de gekozen leerlingenmap worden verwerkt. Excel-tijdelijke bestanden worden overgeslagen.

## Tests uitvoeren

```bash
python3 -m pip install -r requirements.txt
python3 -m unittest -v
```

GitHub Actions voert de tests automatisch uit voor Python 3.10 tot en met 3.13 bij elke push en pull request.

## Wat wordt vergeleken?

De eerste versie vergelijkt celwaarden en formules, celopmaak, werkbladnamen en -volgorde, zichtbaarheid, samengevoegde cellen, rij- en kolominstellingen, tabellen, benoemde bereiken, gegevensvalidatie, voorwaardelijke opmaak en grafieken. Aaneengesloten formule- en opmaakverschillen met hetzelfde patroon in een rij of kolom worden als één melding getoond; identieke opmaakverschillen in meerdere losse bereiken worden ook samengevoegd. Grote aaneengesloten gebieden met celverschillen worden samengevat met een bereik, het aantal afwijkende cellen en enkele voorbeelden. Rij- en kolominstellingen worden eveneens compact per bereik weergegeven. Voorwaardelijke opmaak wordt in het rapport samengevat met bereik, voorwaarde en relevante opmaak; interne Excel-regelnummers worden genegeerd. Grafieken worden op constructie vergeleken, zoals type, titel, reeksen en gegevensbereiken. De celwaarden/caches waarop een grafiek gebaseerd is, worden niet als grafiekverschil gezien. Grafiekverschillen worden kort en leesbaar beschreven; de positie op het werkblad wordt genegeerd.

De vergelijking van celinhoud en opmaak is exact, zonder automatische beoordeling van gelijkwaardige formules. Rijhoogtes en kolombreedtes worden alleen als afwijkend gemeld wanneer ze meer dan 20% afwijken van het model. Rijen of kolommen met dezelfde afmeting in het model moeten onderling ook dezelfde afmeting hebben in het leerlingbestand. Overige rij- en kolominstellingen worden exact vergeleken. Autofilters worden op bereik en criteria gecontroleerd en als één leesbare melding weergegeven. Als het model een sorteerinstelling bevat, wordt de data die aan de model-filtercriteria voldoet op de verwachte sorteervolgorde gecontroleerd; een foutieve sortering geeft één melding. Wanneer de records verder identiek zijn maar anders gerangschikt, worden de daardoor ontstane celverschillen niet afzonderlijk getoond. Kolombreedtes die Excel als één bereik opslaat, worden per kolom vergeleken en verschillen worden per betrokken kolom gerapporteerd. Onbelangrijke Excel-verschillen in de schrijfwijze van getalnotaties (zoals een al dan niet ge-escapete spatie) worden genormaliseerd; inhoudelijk verschillende notaties blijven zichtbaar. Er wordt nog geen score toegekend. Excel-functies die `openpyxl` niet volledig kan inlezen of bewaren, zoals sommige slicers, externe koppelingen of ingesloten objecten, vallen buiten deze eerste versie. Gebruik de app daarom op kopieën als bestanden zulke functies bevatten.

Leerlingbestanden (`.xlsx`/`.xlsm`) en rapportmappen zijn standaard uitgesloten van Git via `.gitignore`, zodat leerlinggegevens niet per ongeluk worden gepubliceerd. Dit project heeft momenteel nog geen open-sourcelicentie.
