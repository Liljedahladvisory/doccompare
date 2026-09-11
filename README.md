# DocCompare

DocCompare jämför två rena `.docx`-versioner och skapar en PDF med ändringar i löptexten. Den nya exportkedjan använder Microsoft Word för både jämförelse och sidlayout. GUI och CLI anropar samma tjänst.

Denna gren är en **provversion av den ombyggda exportkedjan**. Den behöver provas på representativa verksamhetsdokument innan den ersätter den installerade versionen.

## Användning

Kräver macOS, installerad Microsoft Word och tillåtelse att styra Word via Automation. Källfilerna måste vara `.docx` utan befintliga spårade ändringar. Inbäddade OLE-objekt och altChunk stoppas tills de kan hanteras säkert. PDF-indata stöds inte.

```sh
doccompare tidigare.docx ny.docx -o jamforelse.pdf --author "DocCompare"
doccompare-gui
```

Stäng eventuella modala Word-dialoger innan jämförelsen startas. Word arbetar i kortvariga, unikt namngivna arbetskopior i sin Office-sandbox. Appen köar samtidiga DocCompare-jobb. Den stänger bara arbetskopior som den själv skapat.

- Blå understrykning: tillagd text.
- Röd överstrykning: borttagen text.
- Grön dubbelmarkering: flyttad text, när Word identifierar en flytt.
- Violett: ändrad formatering. Struktur och andra dokumentdelar beskrivs även i bilagan.

Word sätter dokumentsidorna. En separat svensk jämförelsebilaga fogas sist utan att sätta om dem. Längre ändringar kan ge andra rad- och sidbrytningar än i en ren källversion.

## Kvalitetskontroller och avgränsning

1. Granska DOCX-paketen och stoppa redan spårade ändringar eller uttryckligen otillåtna objekt.
2. Jämför arbetskopiorna i Word med formatändringar aktiverade.
3. Spara Words redline och exportera den med ändringar i löptexten.
4. Låt Word acceptera respektive avvisa ändringarna i separata kontrollkopior.
5. Stäm av text, fältinstruktioner, länkmål, bildreferenser och stycke-/tabellgränser mot respektive källa. Beräknade fältresultat och tomma stycken normaliseras i denna kontroll.
6. Läs statistiken från Words faktiska revisioner. Det finns ingen separat approximativ diff som kan ge andra statistikvärden.
7. Validera PDF och kontrollera att sidformat och sidornas innehållsströmmar överlever sammanfogningen oförändrade. Publicera lokalt genom atomiskt filbyte först när alla steg lyckats.

Ett misslyckande lämnar en befintlig resultatfil orörd. Appen växlar inte tyst till HTML- eller egen OOXML-sättning. Den äldre motorn finns kvar för dess tidigare tester men används inte av GUI, CLI eller adapterfasaden.

Kontrollerna är **inte en fullständig visuell eller semantisk bevisning**. Bland annat jämförs inte all style-arv, numrering, tabellgeometri eller placering av sidhuvuden mellan avsnitt automatiskt. Bildreferenser jämförs efter innehåll, men inte all bildgeometri. Word kan normalisera format och kan klassificera en flytt som borttagning plus tillägg. Granska verksamhetskritiska resultat visuellt före extern leverans. För redan spårade dokument behövs ett framtida, uttryckligt val av jämförelsebas.

## Arkitektur

```text
GUI / CLI / kompatibel adapter
  -> comparison.service.compare_documents
     -> revisions.preflight
     -> word_bridge (lås, arbetskopior, Word-automation)
     -> revisions.verify_projections + revision_ledger
     -> quality_report (separat bilaga, PDF-kontroll, sammanfogning)
     -> atomiskt byte av resultatfil
```

## Utveckling och test

Python 3.10+. Installera projektets beroenden i en separat virtuell miljö. WeasyPrint används bara för bilagan och kräver Pango på macOS. Beroendena anges i `pyproject.toml`; inga nya produktionsberoenden tillkom i exportombyggnaden.

```sh
python -m pip install -e .
python -m pip install pytest
python -m pytest tests -q
# Opt-in: syntetiska integrationstester som faktiskt styr installerad Word.
DOCCOMPARE_WORD_TESTS=1 python -m pytest tests/integration -q
```

De automatiska testerna kontrollerar innehåll och felhantering. De ersätter inte visuell granskning av referens-PDF:er. Se `docs/layout-quality.md` för provversionens valideringsstatus och kvarvarande arbete.
