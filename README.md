# Excel and Word Checker

Een lokale desktopapp voor het vergelijken van een modeloplossing met `.xlsx`-bestanden van leerlingen. De app maakt voor elk leerlingbestand een afzonderlijk HTML-rapport dat je in een browser kunt openen.

## Opbouw

De mappen `Excel/` en `Word/` bevatten zelfstandige apps en documentvergelijkers. `Shared/` bevat de herbruikbare documentworkflow, rapportviewer en PDF-export. De documentworkflow krijgt een documentvergelijker en rapportgenerator als functies mee, zodat beide apps dezelfde algemene verwerking kunnen gebruiken met eigen documenttype-specifieke code.

## Starten

Vereist Python 3.10 of hoger.

```bash
python3 -m pip install -r Excel/requirements.txt
python3 -m Excel.app
```

Voor Word:

```bash
python3 -m pip install -r Word/requirements.txt
python3 -m Word.app
```

In de Word-app kies je de modeloplossing, een map met leerlingbestanden en de rapportmap. De app verwerkt `.docx`-bestanden in de map en submappen. De puntenverdeling begint met **Document opschonen**, voor dubbele spaties, lege alinea's en niet-verwijderde `(1)`/`(2)`-stijlaanwijzingen. Daarna volgt **Stijlen op juiste alinea's**, dat controleert of elke alinea de modelstijl gebruikt; vervolgens worden de stijldefinities per stijl vergeleken in de volgorde Titel, Ondertitel, Kop 1 t/m Kop 9, Standaard en daarna overige gebruikte stijlen. Tot slot volgen de controles voor automatische nummering, inhoudsopgave en pagina-eindes. Stel desgewenst een bestandsnaameinde in (zonder `.docx`) en puntenmaxima per onderdeel. Spaties en hoofdletters in het bestandsnaameinde worden genegeerd. Standaardstijlen zoals `Normal` en `Heading 1` worden getoond als `Standaard` en `Kop 1`; zelf aangemaakte stijlnamen blijven ongewijzigd. De inhoudsopgave wordt als geheel gecontroleerd op een automatische inhoudsopgave en op ontbrekende, extra of gewijzigde vermeldingen inclusief paginanummers. Opmaakstijlen van de inhoudsopgave worden niet beoordeeld. Een onderdeel zonder verschillen krijgt automatisch alle ingestelde punten. In het rapport staan de scorevelden bij het onderdeel waarop ze betrekking hebben. Verschillen in stijldefinities worden per instelling getoond met Nederlandse Word-termen, zoals **Lettertype**, **Grootte** en **Uitlijning**. Dubbele spaties zijn per onderdeel gegroepeerd; klap de melding open om de alinea's en exacte plaatsen te bekijken. Gewijzigde tekens en ingevoerde tekstwaarden zijn eveneens te bekijken.

Kies in de app de modeloplossing, de hoofdmap met leerlingbestanden en de map voor de rapporten. Vul eventueel de gewenste tekenreeks in waarop de bestandsnaam moet eindigen, zonder `.xlsx`; die extensie wordt verondersteld. Alleen `.xlsx`-bestanden die aan dit einde voldoen worden verwerkt. Spaties en hoofdletters in dit einde worden genegeerd, zodat bijvoorbeeld `Word1` ook past bij `Word 1`. Laat het veld leeg om alle `.xlsx`-bestanden in de gekozen map en onderliggende mappen te verwerken. Excel-tijdelijke bestanden worden overgeslagen. De rapportbestandsnamen bevatten de mapstructuur zodat bestanden met dezelfde naam herkenbaar blijven.

Selecteer eerst de modeloplossing; de GUI leest de werkbladen automatisch in. Stel vervolgens per werkblad en beoordelingsonderdeel het maximumaantal punten in. Lege velden worden niet beoordeeld en verschijnen niet in het rapport. Onderdelen zonder gevonden fouten krijgen automatisch de maximumscore. Voor onderdelen met fouten vul je de behaalde punten direct naast de titel in; een leeg veld telt als nul en de maxima kunnen niet worden gewijzigd. De totaalscore over alle werkbladen verschijnt onderaan.

Na het vergelijken gebruik je in dezelfde GUI de knop **Open rapporten en vul scores in**. Dat start een lokale rapportviewer in je browser. Boven- en onderaan elk rapport kun je naar de vorige of volgende toets gaan; de navigatie loopt door van de laatste naar de eerste en omgekeerd. Via **Overzicht** keer je terug naar de klaslijst. Een rapportlink licht groen op zodra alle scores zijn ingevuld. Bovenaan rechts staat de positie in de klas, bijvoorbeeld `2 / 7`. Ingevulde scores worden automatisch opgeslagen in de HTML-rapporten en als reserve in `.excel_checker_scores/` in de rapportmap. Met **Beëindig evaluatie** op het overzicht of in een rapport sluit je de rapportviewer; de knop is niet zichtbaar in afdrukken/PDF's. De server is alleen vanaf deze computer bereikbaar.

Klik op **Exporteer rapporten naar PDF** om de HTML-rapporten in de gekozen rapportmap naar aparte PDF-bestanden om te zetten. Kies de uitvoermap wanneer daarom wordt gevraagd. Dit vereist een eenmalige installatie van Chromium:

```bash
python3 -m playwright install chromium
```

Playwright gebruikt Chromium om rapporten lokaal te renderen; er is geen internetverbinding nodig tijdens het omzetten. De navigatiebediening van de rapportviewer wordt niet in PDF-bestanden afgedrukt.

## Tests uitvoeren

```bash
python3 -m pip install -r Excel/requirements.txt -r Word/requirements.txt
python3 -m unittest -v
```

GitHub Actions voert de tests automatisch uit voor Python 3.10 tot en met 3.13 bij elke push en pull request.

## Wat wordt vergeleken?

De eerste versie vergelijkt celwaarden en formules, celopmaak, werkbladnamen en -volgorde, zichtbaarheid, samengevoegde cellen, rij- en kolominstellingen, tabellen, benoemde bereiken, gegevensvalidatie, voorwaardelijke opmaak en grafieken. Aaneengesloten formule- en opmaakverschillen met hetzelfde patroon in een rij of kolom worden als één melding getoond; identieke opmaakverschillen in meerdere losse bereiken worden ook samengevoegd. Grote aaneengesloten gebieden met celverschillen worden samengevat met een bereik, het aantal afwijkende cellen en enkele voorbeelden. Rij- en kolominstellingen worden eveneens compact per bereik weergegeven. Voorwaardelijke opmaak wordt in het rapport samengevat met bereik, voorwaarde en relevante opmaak; interne Excel-regelnummers worden genegeerd. Grafieken worden op constructie vergeleken, zoals type, titel, reeksen en gegevensbereiken. De celwaarden/caches waarop een grafiek gebaseerd is, worden niet als grafiekverschil gezien. Grafiekverschillen worden kort en leesbaar beschreven; de positie op het werkblad wordt genegeerd.

De vergelijking van celinhoud en opmaak is exact, zonder automatische beoordeling van gelijkwaardige formules. Rijhoogtes en kolombreedtes worden alleen als afwijkend gemeld wanneer ze meer dan 20% afwijken van het model. Rijen of kolommen met dezelfde afmeting in het model moeten onderling ook dezelfde afmeting hebben in het leerlingbestand. Overige rij- en kolominstellingen worden exact vergeleken. Autofilters worden op bereik en criteria gecontroleerd en als één leesbare melding weergegeven. Als het model een sorteerinstelling bevat, wordt de data die aan de model-filtercriteria voldoet op de verwachte sorteervolgorde gecontroleerd; een foutieve sortering geeft één melding. Wanneer de records verder identiek zijn maar anders gerangschikt, worden de daardoor ontstane celverschillen niet afzonderlijk getoond. Kolombreedtes die Excel als één bereik opslaat, worden per kolom vergeleken en verschillen worden per betrokken kolom gerapporteerd. Onbelangrijke Excel-verschillen in de schrijfwijze van getalnotaties (zoals een al dan niet ge-escapete spatie) worden genormaliseerd; inhoudelijk verschillende notaties blijven zichtbaar. Excel-functies die `openpyxl` niet volledig kan inlezen of bewaren, zoals sommige slicers, externe koppelingen of ingesloten objecten, vallen buiten deze eerste versie. Gebruik de app daarom op kopieën als bestanden zulke functies bevatten.

Leerlingbestanden (`.xlsx`/`.xlsm`), PDF's en rapportmappen zijn standaard uitgesloten van Git via `.gitignore`, zodat leerlinggegevens niet per ongeluk worden gepubliceerd. Dit project heeft momenteel nog geen open-sourcelicentie.

De Word-vergelijker controleert alinea-tekst, opschoonpunten (dubbele spaties, lege alinea's en verwijderde titelmarkeringen `(1)`/`(2)`), de toegepaste stijl per alinea en daarna de definities van gebruikte stijlen. Ook automatische nummering, de inhoud en automatische generatie van de inhoudsopgave en pagina-/sectie-eindes worden gecontroleerd. De inhoudsopmaakstijlen worden niet beoordeeld. Tekstvergelijkingen markeren verwijderde en toegevoegde tekens. Alleen `.docx`-bestanden worden ondersteund; wijzigingen aan Word-functies die niet zijn opgenomen in deze controles, zoals afbeeldingen, tabellenopmaak en documentmetadata, worden niet beoordeeld.
