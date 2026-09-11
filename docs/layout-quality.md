# Exportkedja 0.3.0a7: lokal installation

## Paketeringsfel rättat i 0.3.0a7

Versionerna 0.3.0a1 till 0.3.0a6 kunde starta GUI:t men saknade den nya tjänsten i den modulplats som py2apps startkod faktiskt valde. Den återanvända runtime-miljön innehöll äldre kataloger med punktnamn (`doccompare.comparison`, `doccompare.rendering`, `doccompare.parsers`). Startkodens Finder prioriterade dessa framför det nya vanliga paketträdet. Det gav `No module named 'doccompare.comparison.service'` vid jämförelse.

Tidigare tester med inbäddad Python lade det nya paketträdet på sys.path men körde inte py2apps import-hook. De verifierade motorn och filinnehållet, men fångade inte denna startskillnad. Startkontrollerna av GUI:t ska inte tolkas som fungerande jämförelser i dessa äldre appbyggen.

Byggskriptet uppdaterar nu också alla DocCompare-underpaket som anges i startkodens `_path_hooks`. Ett saknat underpaket stoppar bygget. Före arkivering kör den paketerade Python-tolken `verify_local_bundle.py`, som exekverar den faktiska startkoden med dess Finder och kontrollerar att alla berörda importer kommer från aktuella paketerade källor. Bara det sista GUI-startanropet ersätts av kontrollen.

Verifiering: den nya kontrollen återskapade exakt användarens importfel i 0.3.0a6 och godkändes i 0.3.0a7. Totalt 34 ordinarie tester godkändes, inklusive regressionstest som först visar felet med ett äldre flyttat underpaket och därefter importerar den nya tjänsten efter uppdatering. De 15 tidigare Word-scenarierna kördes inte på nytt. En syntetisk jämförelse kördes däremot genom de verkliga start-hookarna, både i det extraherade arkivet och efter installation i Program. Båda gav tre PDF-sidor och +8/-1 ord med godkänd projektionskontroll. Den installerade appen startades och versionsmärkning 0.3.0a7 granskades. Ett komplett GUI-klickflöde är fortfarande inte verifierat via UI-verktyget.

ckglib kartlade GUI-anropet och den nya paketeringsfunktionen. Byggskriptet innehåller kod på modulnivå och granskades manuellt efter parservarningen. Verifieringsskript, regressionstester och versionsfiler är avsiktliga tillägg utanför den smala anropsplanen.

## GUI 0.3.0a4

Svante bad om samma GUI-färger som Meeting Recorder innan installation. Palettens 14 färger hämtades från `meeting-recorder-llt-pr2/meeting_recorder.py`: ljus bakgrund, vita ytor, mörk text och blå accent. Förloppsindikatorn använder samma accent och sidfotens text använder textfärg i stället för kantfärg. Version 0.3.0a3:s PDF-attribution behålls. Ingen jämförelse- eller PDF-logik ändrades.

Det extraherade v4-arkivets signatur och samtliga paketerade Python-källfiler kontrollerades. Den exakta appen startades och huvudfönstret granskades visuellt med versionsmärkning 0.3.0a4. Tk-inställningsdialogen kunde inte öppnas via UI-verktyget; dess färger följer den gemensamma paletten i granskad kod men är inte visuellt verifierade. Inga nya funktionstester behövdes för färgbytet. ckglib granskade GUI-modulen; versionshöjningarna utanför planen är avsiktliga och app.py behöver ingen ändring eftersom gränssnittet är oförändrat. Version 0.3.0a4 installerades lokalt den 11 september 2026; version 0.2.0 sparades som återställningskopia.

## Sammanfattning 0.3.0a2

Efter Svantes granskning har sammanfattningen fått ett kompakt upplägg med inspiration från den tidigare versionen: mindre rubrik, filnamn och datum, färgkodade ändringsantal på egna rader och kort teckenförklaring. Övriga ändringar visas vid behov i en enkel tabell. Filhashar, Word-version och kontrollstatus finns i PDF-metadata. Ingen separat uppskattning av oförändrade ord har återinförts.

Verifiering av denna ändring: 32 ordinarie tester godkända; de 15 testerna mot riktig Word kördes inte på nytt för denna presentationsändring. Prov-PDF:en genererades av den nya appens paketerade rapportmodul. Dess två dokumentsidor har exakt samma innehållsströmmar och sidstorlek som den tidigare godkända prov-PDF:en. Den nya sammanfattningen är en sida. Långa filnamn och en tabell med 40 ändringsrader har också renderats och granskats över två sidor. Slutfilen har öppnats i Förhandsvisning och apparkivets signatur verifierats.

ckglib kartlade `quality_report.render_note` och dess anrop före ändringen. Funktionen behåller sitt gränssnitt, så service, CLI, adapter och integrationstester behövde inga ändringar. Metadataöverföringen i `assemble_pdf` granskades i samma modul. De två versionsfiler som ckglib markerade utanför planen är granskade versionshöjningar till 0.3.0a2. Jämförelsemotorn och begränsningarna för ordinarie release nedan är oförändrade.

Status 2026-09-11: den ombyggda motorn är implementerad och provad mot Microsoft Word 16.112 på macOS. Version 0.3.0a4 är installerad lokalt och källändringarna lämnas för granskning via GitHub. Ingen allmän binär release har skapats.

## Implementerat

- Gemensam tjänst för GUI, CLI och befintlig adapterfasad.
- Native Word-jämförelse med formatändringar, utan att avvisa numreringsrevisioner.
- Unika arbetskopior inom Office-sandboxen, lås mellan DocCompare-jobb och avgränsad stängning av arbetskopior.
- Tillfällig, enhetlig markering av tillägg, borttagningar, flyttar och format. Tolv Word-inställningar återställs efter körningen.
- Word genererar accepterade och avvisade kontrollkopior. Projektionerna jämförs med källornas text, stycke-/rad-/cellgränser, fältinstruktioner, länkmål och refererade bilders innehåll.
- Statistik från samma revisioner som ligger bakom dokumentet. Den gamla separata approximativa ordjämförelsen används inte.
- Svensk bilaga efter dokumentet. PDF-sidornas storlek och innehållsströmmar kontrolleras efter sammanfogning.
- Atomiskt byte av resultatfilen efter lyckad kontroll. Fel ger inget tyst byte av renderingsmotor.
- GUI:s fördröjda felmeddelande behåller undantagstexten även efter att arbetartråden lämnat sin exception-handler.

## Verifiering

Slutkörningen gav **47 godkända tester på 68,77 sekunder**: 32 ordinarie tester och 15 opt-in-tester som styr riktig Word. Nativefallen omfattar oförändrat dokument, ordbyte, tillagd och borttagen tabellrad, ändrat sidhuvud, ändrad sidfot, teckenformat, tillagt och borttaget stycke, tillagd numrerad punkt, lång ersättningstext, ändrad fotnot, bibehållen bild, ändrad bild och bibehållen hyperlänk.

Feltester kontrollerar att tidigare resultat lämnas orört vid fel i Word, projektionskontroll, bilagerendering och sammanfogning. Andra tester kontrollerar låsning, strängescaping, fältresultat, tabbar, befintliga revisioner och GUI:s felcallback.

Det extraherade leveransarkivet kontrollerades med `codesign --verify --deep --strict`. Dess egen inbäddade Python och dess paketerade moduler har producerat den levererade prov-PDF:en. En separat kontroll visade att listan över öppna Word-dokument och samtliga tolv berörda visningsinställningar var desamma före och efter körningen.

Prov-PDF:ens dokumentsidor och bilaga har renderats för visuell kontroll. Även borttagna tabellrader och ändrad listnumrering har granskats visuellt. Den exakta leverans-PDF:en har öppnats i Förhandsvisning.

**GUI-begränsning:** appfönstret startar och ritas korrekt. Hela flödet med filval, jämförelseknapp och automatisk PDF-öppning har inte verifierats genom UI-verktyget; dess klick/tangentbord kunde inte aktivera Tk-filväljaren. Detta bevisar inte ett fel i filväljaren, men räknas inte heller som ett godkänt UI-test. Arbetartrådens felhantering är separat automatiskt testad.

## ckglib-granskning

Före ändring kartlades `produce_pdf`, `_diff_para` och `MacWordAdapter.compare_and_export`. GUI:s och CLI:s dynamiska adapteranrop lästes också manuellt. Efter ändring kartlade `validate-diff service.compare_documents --depth 3 --include-tests` den nya tjänsten och dess anrop från GUI, CLI, adapter och tester; de fem analyserade anropskanterna hade hög säkerhet.

ckglib markerade nya hjälparmoduler, dokumentation, versionsfiler och paketeringsskript utanför den bakåtriktade anropsplanen. Dessa är uttryckliga delar av ändringen och granskades manuellt. Detta är alltså en granskad avvikelse från planens filurval, inte en automatiskt helt grön blast-radius-kontroll. Den gamla OOXML-/HTML-motorn har medvetet lämnats utanför den aktiva kedjan.

## Före ordinarie release

1. Kör hela klickflödet manuellt i provappen på denna Mac och på en ren testinstallation, inklusive Automation-tillstånd och normal licenskontroll.
2. Godkänn ett representativt referensbibliotek från verksamheten: bland annat flernivånumrering, innehållsförteckning, korsreferenser, avsnitt med olika sidhuvuden/sidorientering, tabellkolumner och sammanslagna celler, flyttar och långa avtalspar.
3. Bygg ut projektionskontrollen för numreringssemantik, stil-arv och positionsberoende innehåll. Nuvarande kontroll verifierar inte all formatering eller att ett visst sidhuvud är kopplat till rätt avsnitt.
4. Bestäm uttrycklig policy för dokument som redan innehåller revisioner, kommentarer, formulär och inbäddade objekt. Redan spårade ändringar och OLE/altChunk stoppas i denna version.
5. Gör en ren, låst releasebyggnad och normal signering/notarisering. Den lokala piloten återanvänder beroenderuntime från den installerade DocCompare-appen; den är inte en reproducerbar releasebyggnad.

## Lokal pilot

`scripts/build_local_preview.py` bygger ett separat ZIP-arkiv med **DocCompare Preview.app** från angiven befintlig runtime. Den ändrar inte `/Applications/DocCompare.app`. Signering sker i systemets tillfälliga lokala katalog eftersom FileProvider/iCloud återlägger Finder-metadata på appkataloger under Documents. Arkivet innehåller en lokalt ad hoc-signerad provapp, utan Apple-notarisering. Packa upp piloten i en lokal katalog, exempelvis Hämtade filer. Piloten använder befintlig lokal licens och språkinställning; dokumenten behandlas lokalt av Word.

## Installationskontroll

Den lokala installationen använder namnet DocCompare och det ordinarie bundle-id:t; preview-flaggan har tagits bort och versionsmetadata anger 0.3.0a4. Python-källorna är identiska med det granskade v4-arkivet. Appen signerades lokalt igen och `codesign --verify --deep --strict` godkändes. Den exakta appen i `/Applications/DocCompare.app` startades och huvudfönstrets version och färgpalett kontrollerades. Detta är en installations- och startkontroll; begränsningarna för ett fullständigt GUI-test och en ordinarie release ovan kvarstår.

## Logotyp 0.3.0a5

Sidhuvudet visar nu Liljedahl Advisorys logotypbild i stället för det konfigurerade kontonamnet. Bilden kommer från webbprojektets `public/images/liljedahl-logo.png` och återges oförvrängd mot en mörk yta för kontrast. Kontonamn och språk finns kvar i inställningarna; ändring av namn ersätter inte logotypen. Logotypen paketeras i både wheel och macOS-app.

32 ordinarie tester godkändes (15 Word-tester inte omkörda för denna GUI-ändring). Den installerade appens signatur och källfiler kontrollerades och logotypen granskades i huvudfönstret med versionsmärkning 0.3.0a5. ckglib granskade GUI-ändringen; bildfilen, paketeringsraden och versionshöjningarna utanför anropsplanen är uttryckliga delar av ändringen. Startmodulen behåller sitt gränssnitt och behöver ingen ändring.

## Logotypbakgrund 0.3.0a6

Efter visuell återkoppling har den mörka ytan bakom logotypen tagits bort. Logotypbilden är oförändrad och ligger direkt mot appens ljusa bakgrund. Installerad version 0.3.0a6 har startats och granskats visuellt; paketerade källfiler, logotyp och signatur har kontrollerats. ckglib bekräftar samma GUI-anropskedja; versionsfilerna är avsiktliga ändringar utanför funktionsplanen. Inga funktioner eller PDF-färger ändrades.
