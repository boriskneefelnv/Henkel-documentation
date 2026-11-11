# Henkel Operations Center — kennisbank

Bron: `Henkel_tool_v.html`, één zelfstandig HTML-bestand van 3,3 MB voor Neele-Vat × Henkel export operations.
Alles draait lokaal in de browser; er wordt niets geüpload.

---

## 1. Opbouw van het bestand

| Onderdeel | Omvang | Inhoud |
|---|---|---|
| CDN-libraries | 4 scripts | pdf.js 3.11.174, SheetJS (xlsx) 0.18.5, JSZip 3.10.1, ExcelJS 4.4.0 (allemaal via cdnjs) |
| Inline gebundeld | ~1,1 MB | `@kenjiuno/msgreader` 1.28.0 (Outlook .msg lezen) en `pdf-lib` 1.17.1 (PDF's maken) |
| CSS | ~78 KB | Neele-Vat-huisstijl ("Business Update"): navy, blauw `#266B94`, Henkel-rood `#E1000F`, Calibri/Tahoma |
| HTML | ~33 KB | navigatierail met 2 hoofdmodules, topbar en 12 pagina's |
| Applicatie-JS | ~580 KB | alle logica (plus ingebedde logo's en een blanco DGD-PDF in base64) |
| `#dashPayload` | 1,4 MB base64 | het complete Operations Dashboard (React + SheetJS) als los HTML-document, dat pas bij het eerste bezoek in een iframe (`srcdoc`) wordt geladen |
| Lazy geladen | — | Tesseract.js 5.1.1 (OCR van containernummer-foto's), alleen bij gebruik |

Ingebedde binaire assets: Neele-Vat- en Henkel-logo (sidebar), `KC_LOGO` (Henkel-logo voor het klantrapport), `TO_LOGO_B64` (Neele-Vat-logo, JPEG 500×104, voor de transportopdracht) en `DGD_BLANK_B64` (een leeg Henkel-DGD-stramien).

### Navigatie
- **Hoofdmodule 01 — Operations Tools** (11 modules) en **Hoofdmodule 02 — Operations Dashboard** (1 module).
- `showPage()` wordt door elke later toegevoegde module "ingepakt" (wrapper-patroon), zodat bestaande modules ongewijzigd blijven.
- Sneltoets **Alt+1…9** opent tabDraai, Booking, Klant, Dgd, Maak, To, Opdracht, Bl en Cert. Full set, Ocean boeken en het Dashboard hebben geen sneltoets.
- Kwaliteitslaag: een melding als een CDN-library niet laadt, scroll naar boven bij het wisselen van module en de browsertitel volgt de actieve module.

### Gedeelde bouwstenen
- `pdfToLines()` zet PDF om naar tekstregels en vangt vier praktijkproblemen af: y-drift (ketenen op de vorige y), baseline-jitter (fragmenten ≤5,5 pt samenvoegen als ze niet overlappen), gedraaide zijstroken (worden genegeerd) en gebroken getallen (lijmen bij een gat <1 pt, of <2,5 pt cijfer-op-cijfer).
- `extractMsg()` en `extractEml()` halen onderwerp, bijlagen (als File) en mailtekst uit een mail. `decodeMimeWord()` decodeert RFC 2047.
- `buildEml()` maakt een `.eml` met `X-Unsent: 1`, base64-body en een RFC 2047-onderwerp. Outlook (klassiek) opent dat als concept met aan/cc/bijlagen al ingevuld. Nieuwe Outlook opent .eml alleen-lezen.
- `copyText()` is een klembord-functie met een `execCommand`-fallback.

---

## 2. Module 01 — Draai & Controle

**Doel:** van Henkel-facturen, packing lists en de zending-Excel de draaitabel per HS-code maken voor customs (EXA/COO/EUR1), plus de aanvraagmail.

**Invoer:** factuur-PDF, packing list-PDF, zending-Excel (.xlsx), shipping guide (.docx), of de hele Outlook-mail (.msg/.eml).
Overige bestandstypen (bijvoorbeeld logo's) worden stil overgeslagen. Onleesbare bestanden worden 12 seconden lang gemeld.

### Henkel-documenten lezen (`parseHenkelDoc`)
- **Type:** "PACKING LIST" of "LISTE DE COLISAGE" → PL; "Invoice" of "Facture" → INV. Ondersteunt Engels en Frans.
- **Documentnummer:** `No. xxx From dd.mm.yyyy`, of `N°. … du …`.
- **Decimaalteken:**
  - Eerst detectie op ondubbelzinnige getallen (meerderheid van getallen met zowel punt als komma).
  - Terugval per type: PL = EU-notatie (komma decimaal), factuur = US-notatie (punt decimaal).
  - Wijken detectie en type van elkaar af, dan volgt een waarschuwing.
  - Henkel schrijft gewichten altijd met 3 decimalen, dus "181.620" is 181,620.
- **Brutokop:** "G" in de kop wordt alleen als gram gelezen als de waarde gedeeld door 1000 bij de regelsom past (sommige facturen zeggen "G" terwijl het kg is).
- **Toyota Misr-variant:**
  - "Total Gross Weight =" telt; de variant "including pallet weight" wordt uitgesloten.
  - Sub-total plus Freight charges: de regelsom hoort bij Sub-total.
- **Intercompany (Henkel Jebel Ali):** "Total Article Value" plus "Total To Be Paid" (vaak 0,00). De check "Items total = Total Amount" wordt dan overgeslagen.
- **Factuurregels:**
  - Regex met "Our + Your Mat.No." (moeten gelijk zijn), anders greedy naar het meest rechtse IDH.
  - IDH is 2-7 cijfers; korte IDH's zoals "41" bestaan.
  - Gewichtsregel `GROSS/KG … NET/KG`; regels met `/G` worden omgerekend naar kg.
- **PL-regels:** posregel met qty/unit/gross/tare/net, daaronder IDH en omschrijving, `DELIVERY NOTE: … FROM:`, `CN:`/`KN:` (HS) en `COUNTRY OF ORIGIN`/`PAYS D'ORIGINE`.
- **Vangnetten:**
  - Regelsom ≠ documenttotaal geeft "REGELS MOGELIJK GEMIST".
  - Een factuurregel zonder GROSS/NET wordt gemeld.
  - Een PL-positie zonder IDH wordt gemeld.
- **Pallets:** `Total pallet count = N` en `Total Gross Weight including pallet weight = X` (tolerante notatie).

### Sets en koppeling
- Een set is 1 factuur + 1 PL met hetzelfde documentnummer. Een dubbele upload met identieke inhoud wordt genegeerd; met andere inhoud is het een conflict.
- Een factuurregel wordt aan een PL-regel gekoppeld via `takePl()` met een score: zelfde IDH, delivery-match +4 (mismatch −2), bruto-match +2, aantal-match +1.
- Ontbreekt het gewicht op de factuur (Toyota Misr), dan wordt het PL-gewicht gebruikt.
- `qtyEq()` accepteert exacte gelijkheid, gelijkheid na afronding, en een factor 1000, 100 of 10 (US- tegenover verkorte PL-notatie).
- `normUnit()`: CON is gelijk aan "Cons. unit"; SHP (shipping unit) is iets anders. Bij verschillende eenheden en gelijke gewichten wordt het aantal niet vergeleken.

### Controles per set (FOUT/OK)
1. De VALUE-regels tellen op tot Items total (of tot Sub-total bij losse vracht).
2. Items total = Total Amount (niet bij intercompany).
3. De GW-regels tellen op tot het factuurbruto.
4. Factuurbruto = PL-bruto.
5. De NW-regels tellen op tot het PL-netto.
6. Op de PL: bruto = tarra + netto.
7. Per regel: GW, NW en aantal op factuur en PL zijn gelijk. Regels die maar op één document staan worden gemeld.

Toleranties: gewicht 0,0015, waarde 0,005, gewichtstotalen 0,005, waardetotalen 0,02.
Elke mislukte check krijgt een uitleg in gewone taal: wat er mis is en wat te doen (`explainCheck`).

### #KA-toewijzing (`allocateKa`)
De bron is de zending-Excel: kolommen Reference/Delivery + IDH + "# KA" (en optioneel Qty, Total PL + KA, Bonded goods, Product description). De tool probeert in deze volgorde:
1. Reference + IDH met één factuurregel.
2. Plus het aantal.
3. Een **combinatie van batches** waarvan de aantallen exact optellen (subset-sum; pool van 2 tot 8, kleinste combinatie wint). Voorbeeld: factuur 9020 = Excel 3740 + 5280.
4. Volgorde.
5. Alleen IDH.
6. Vangnet: de resterende #KA naar rato van het nettogewicht verdelen.

Daarna wordt per set+IDH **naar boven afgerond** op hele colli: floor per regel, en het overschot naar de grootste restfracties. In de Eindtabel staat `CEILING(SUMIF(…),1)`.

### Pallets per invoice (`computePallets`)
- Bron: "Total PL + KA" (fractie per regel). Het containertotaal is "Totaal pallets container" naar boven afgerond. "Max pallets" is de capaciteit en wordt alleen als terugval gebruikt.
- Elke set+HS krijgt minimaal 1 pallet; de rest wordt verdeeld volgens de Hamilton-methode naar rato.
- "Kopieer voor mail" geeft een tekstoverzicht.

### T-1 / wederuitvoer
- Kolom "Bonded goods": 1 = T-1, 0 = vrij.
- T-1-regels gaan uit de EXA-draaitabel naar een **apart bestand** `*_wederuitvoer_draai.xlsx`.
- Een regel die zowel 1 als 0 heeft is een conflict. De regel gaat dan veiligheidshalve naar de wederuitvoer en wordt gemeld.

### Palletgewicht op de certificaten (les 14035, Bodo Möller)
- Als "Totaal incl. pallets" op de documenten staat (of handmatig is ingevuld), wordt het palletgewicht naar rato van het bruto over de regels verdeeld. Het afrondingsverschil gaat naar de zwaarste regel.
- Er komen extra kolommen PALLET KG en GW INCL PLT bij.
- Meer dan 9 IDH's **en** palletgewicht maakt een blad **Certificaten**:
  - Maximaal 9 IDH's per CVO/EUR1, gesorteerd op IDH.
  - Op de EUR1 komt het totaalgewicht.
  - Op de EXA komt "loaded on N pallets".

### Excel-uitvoer
- **Met één zending-Excel:** het origineel wordt behouden met alle opmaak (via JSZip/XML-injectie). Het blad wordt hernoemd naar "Inputblad" en verwijzingen worden bijgewerkt. Oude tool-bladen worden vervangen. calcChain wordt verwijderd en `fullCalcOnLoad` gezet.
- **Met meerdere zending-Excels:** de databody's worden onder elkaar samengevoegd.
- **Zonder Excel:** een gereconstrueerd Inputblad uit de factuur- en PL-gegevens (les 14101).
- **Bladen:**
  - **Douane inputblad** (les 14141 Tunis en 14076): kopie van het Inputblad.
    - Product Description wordt pas de factuurnaam na een vinkje plus bevestiging in een popup.
    - De HS-code (21-cijferige string) wordt de HS van de factuur, of anders de eerste 8 cijfers.
    - Gewichten ≥20.000 worden gelezen als gram en omgerekend naar kg.
    - "Total PL + KA" in Henkels puntnotatie (10.11) wordt 10 + 11 = 21 colli.
  - **Tussentabel:** IDH, DESCRIPTION (factuurnaam), #KA, HS, GW, NW, VALUE.
  - **Eindtabel:** draaitabel per HS met SUMIF-formules.
  - **Controle:** totalen, sets, checks, regels, #KA-bron en T-1-splitsing.
  - **Pallets** en **Certificaten**, waar van toepassing.
- **Downloadpoort:** geblokkeerd bij incomplete sets of onleesbare PDF's. Bij een FOUT kan alleen worden gedownload met het vinkje "Toch downloaden". "Maak e-mail" volgt dezelfde poort.

### Mail naar customs
- Aan `customs.cluster1@neelevat.com`, cc `henkel.dc@neelevat.com`.
- Bijlagen: de draai-Excel (plus het T-1-bestand), alle facturen en PL's, en de guide.
- **Herkenning uit het onderwerp:** IC/EC (de laatste die voorkomt), bestemming, SEA, NV-nummer (`NV 60xxxxxx`) en de zendingnaam (`13930 CASABLANCA IC`).
- **Uit de zending-Excel:** de labels "Master" (wordt Ref) en "Container" (wordt de zendingnaam).
- **Mailtekst:**
  - "Svp voor morgen exa/coo/…"
  - Optioneel "(Nkopies)", "Aub draft opmaken", "Ref", "Euro 1100", "Kantoor van uitgang", "Aub onderstaande vermelden op drafts", "Aub in vak 7 eur 1 revised rules" en "Voor COO>".
  - Bij Casablanca komt er een EUR-1-retouradres bij.
- **Excel-regel CVO | EUR1 | ATR | COC | HS:** "A" als het document wordt aangevraagd, anders 0. COC = KvK-legalisatie. HS = aantal unieke HS-codes min 1.
- **Excel-regels documentregistratie** (toegevoegd 10-11-2026): per aangevinkt document één rij, tab-gescheiden, met de knop "Kopieer voor documentregistratie".
  - Kolommen: master (= Ref/NV-nummer) | 5 lege kolommen | documentcode | documentsoort | aantal | datum van vandaag (dd-mm-jjjj). Dat past op het registratieblad A Master · B BP File · C Reference · D IC/EC · E Dest · F Loading date · G Document code · H Description · I Aantal · J Request date.
  - Aantal: EXA altijd 1; CVO, EUR.1, ATR en legalisatie leeg.
  - Aparte HS-regel: code 10001, omschrijving HS, aantal = aantal extra HS-codes (unieke HS-codes min 1, zelfde als de HS-kolom van de A/0-regel). Verschijnt zodra facturen en packing lists zijn ingelezen. De macro `Documentregistratie_Aanvullen` (`macros/`) werkt alleen op tabblad DCMAASVLAKTE en houdt de groene kolommen B–F steeds gelijk aan blad Uit van de planning; andere kolommen raakt hij niet aan.
  - Codes: 71001 EXA, 72001 ATR, 72002 EUR.1, 72200 CVO, 72301 LEGALISEREN (kvk, legalised invoice en gelegaliseerde factuur samen worden één rij).
  - Voor copy, coc en embassy is geen code bekend. Daarvoor komt geen rij; de tool meldt dat.
  - De knop werkt pas als Ref is ingevuld.

### Shipping guide (.docx) lezen (`parseGuideDoc`)
- Documententabel (kop met "document" en "applicable"): per regel de token (exa/eur1/coo/kvk/coc/embassy/atr/legalised invoice), applicable Y/N, originelen, kopieën en conditions.
- Invoice- en PL-condities zijn interne notities en gaan **niet** mee (les 14186 Balchem). Uitzondering: legalisatie van de factuur → token kvk.
- "revised rules" zet vak 7 aan. COO- en EXA-condities worden de COO-suggestie (veldlabels per regel: NTN, LC, Country of Origin, H.S.CODE, …).
- Regels onder "General remarks" en alle korte gebiedende zinnen ("must", "required", "to be mentioned", …), maximaal 12, gaan naar de remarks. De aantallen originelen/kopieën komen als samenvatting bovenaan.

---

## 3. Module 02 — Booking confirmation

- Leest booking-PDF's van verschillende rederijen en maakt een rij van 10 tab-gescheiden kolommen voor Excel: Container type · ETD · ETA · Carrier · Booking · B/L · BC confirm · Vessel · POL · POD.
- 26 rederijen worden herkend: Hapag, CMA CGM, Maersk, MSC, ONE, COSCO, Evergreen, OOCL, Yang Ming, ZIM, Arkas, Grimaldi, Tarros, Turkon, Sealand, Safmarine, NileDutch, HMM, Wan Hai, PIL, Borchard, Unifeeder, Samskip, Messina, Marfret en X-Press.
- Hapag-etappetabel, CMA-route-tabel (alle pagina's) en generieke labels. Alleen pagina 1, omdat de voorwaarden-pagina's valse labels bevatten.
- Datums worden genormaliseerd naar d-m-jjjj. Containermaat (ISO 22GP/42G1/45G1) wordt "40ft" of "2 x 40ft".
- **BC confirm = vandaag.** De datum op de booking wordt alleen als link aangeboden.

---

## 4. Module 03 — Klantcontrole

**Stap 1:** de eigen zending-Excel (WMS) tegenover de Excel van de klant (SAP).
**Stap 2:** de WMS-picklijst (PDF, blok "Controle van de order") tegenover SAP.

- De rol van een bestand wordt herkend aan de kolomkoppen:
  - Eigen versie: Reference/Delivery + IDH + Batch.
  - Klantversie: Turkse SAP-koppen Teslimat/Malzeme/Parti, of de Engelse Delivery/Material/Batch.
- Leest xlsx, oude xls (BIFF) en UTF-16-tabtekst met .xls-extensie (de Henkel-export). Het getalformaat (EU/US) wordt per bestand bepaald.
- Koppeling op **IDH + batch**. Batches als "SVP DOORGEVEN", "?" of zonder cijfers zijn placeholders en worden apart gemeld.
- Controle van aantal (alleen bij dezelfde eenheid), netto en bruto (tolerantie 0,01 of 0,2%). Gram en kg worden als gelijk beschouwd op een factor 1000 na.
- **Interne fout:** past het eigen bruto in geen enkele lezing (gram of kg) bij het eigen anker (netto Qty bij KG, of Total weight KA), dan staat het onder "eerst zelf controleren" en gaat het **niet** naar de klant.
- Overlap onder 34% betekent waarschijnlijk een andere zending.
- **Uitvoer:**
  - Een Engelse klantmail (alleen de afwijkingen; via mailto).
  - Een Excel-rapport "Shipment verification" (ExcelJS, met Henkel-logo, groene of rode banner).

---

## 5. Module 04 — DGD controleren (Multimodal Dangerous Goods Form)

Leest Henkel-DGD-PDF's, groepeert ze op de 600-referentie en controleert:
- **LTD QTY:** maximaal 30 kg bruto en netto per buitenverpakking (IMDG 3.4), en een combinatieverpakking is verplicht (geen losse drums of IBC's).
- **Gewichten:**
  - Bruto ≥ netto.
  - Bruto = netto is fout: het bruto ontbreekt.
  - Netto 0 bij bruto > 0 is fout (14097 Pattex Multispray). Voorstel: de kg-hoeveelheid, of 90%.
  - Bij een KG-regel moet het netto exact de kg-hoeveelheid zijn.
  - Een regel zonder gewicht is fout.
- **Totalen:** bloktotalen = som van de regels. Vak 19 mag **niet lager** zijn dan de DG-som; het mag hoger zijn, want het is het containertotaal.
- **Vlampunt:**
  - Klasse 3 zonder Flp is fout (IMDG 5.4.1.4.3).
  - Aerosols (2.1) zonder Flp worden ook als fout gemeld: MSC keurt dat af (14097/14107).
- **Concentratie** verplicht voor UN 2031 (salpeterzuur), 2014/2015/2984 (waterstofperoxide), 2790 (azijnzuur), 2796 (zwavelzuur) en 1791 (hypochloriet). Les 14142, CMA.
- **Combi-artikel** IDH 2057216 (LOCTITE EA 9466) moet onder UN3082 + UN2735 staan (14147 Steder). Een lege UN-kop zonder regels wordt gemeld.
- **IMDG-segregatie** (tabel 7.2.4): codes 2, 3 en 4 tussen klassen in één container zijn fout. LTD QTY-blokken zijn vrijgesteld.
- **Automatisch opmaken** van onopgemaakte regels ("192 CARTNS"):
  - UN1950 + stuks → AEROSOLS in fiber board boxes (4G).
  - Overige stuks → METAL CANS.
  - KG → IBC 31HA1 bij ≥500 kg per stuk, anders DRUMS 1A1.
- **Correctie-editor:** UN-tekst, LTD QTY verwijderen, verpakking, netto/bruto en een knop om netto en bruto te wisselen. De gecorrigeerde PDF gebruikt de **originele PDF als ondergrond**: het goederengebied wordt wit afgedekt en opnieuw gezet in Courier 9 pt.

---

## 6. Module 05 — DGD maken v2

**Invoer:** de draai-Excel of de mail, plus maximaal **twee packing lists**. Factuur-PDF's worden overgeslagen. Een meegesleepte systeem-DGD levert alleen de kop (consignee, bestemming en referenties) en de vlampunten. Een MSDS levert het UN-nummer, de PSN, klasse, PG, EmS, MP en het vlampunt.

- Alleen regels met een UN-nummer (niet 0 of leeg) gaan mee. Er wordt gegroepeerd per Engelse UN-info-variant (het Nederlandse deel in het veld wordt afgeknipt), en binnen een blok **samengevoegd per IDH**.
- **Bruto:** uit de packing list (gemeten).
  - Bij voorkeur per IDH + delivery note (één PL kan meerdere containers dekken).
  - Anders het IDH-totaal als dat binnen 10% van de Excel ligt.
  - Anders het Excel-gewicht ("Total weight KA").
- **Netto:**
  - KG-regel: de kg-hoeveelheid (rederij-eis).
  - Stuksregel: het gemeten PL-netto (les 14078).
  - Alleen als terugval: 90% van bruto.
- **Verpakkingen:** alleen deze vier vormen:
  - `n DRUMS PACKING 1A1`
  - `n IBCS PACKING 31HA1`
  - `n AEROSOLS IN m FIBER BOARD BOXES (4G)`
  - `n METAL CANS IN m FIBER BOARD BOXES (4G)`

  Een "stuks"-regel van ≥500 kg per collo is een IBC (les 14078, MSC HazCheck). Bij KG wordt het aantal binnenverpakkingen uit de naam gehaald ("12X0,5KG").
- **LTD QTY:** zo nodig worden er dozen bijgemaakt tot elke doos onder 30 kg blijft. Weegt een buitenverpakking meer dan 30 kg (bijvoorbeeld drums), dan wordt LTD QTY verwijderd (les 14092).
- **Duo-artikelen** (meerdere UN's in het UN-info-veld): **de UN-kolom beslist**. Op de DGD komt alleen dat UN-nummer met het volle gewicht, want een extra UN geldt als "DG toegevoegd na boeking" (les 14010, COSCO). Uitzonderingen:
  - Een vastgelegd combi-artikel (`DGD_COMBI_SPLIT`): EA 9466 wordt automatisch gesplitst, hars UN3082 ⅔ en harder UN2735 ⅓ (Steder 14147, Maersk 14177).
  - Een MSDS met 2 stoffen: beide UN's komen op de DGD, gewicht 50/50 (14118), met een waarschuwing over de boeking.
  - Staat het kolomnummer niet tussen de varianten, dan is het een UN-conflict en volgt een verplichte keuzedialoog.
- **Vak 19:** het containertotaal (alle regels, uit dezelfde bron als de DG-regels). Vak 4: `600-master / shipRef`. Vak 5: forwarder-ref (`10xxxxxxxx` uit het onderwerp). Laadhaven Rotterdam of Antwerp.
- **Kwaliteitspoort** (`dgd2QualityGate`) blokkeert de DGD en de CMR bij:
  - Een combi-artikel dat niet is gesplitst, een ontbrekende concentratie, of een UN-conflict.
  - **Een ontbrekend vlampunt bij klasse 3 of 2.1** (eis van Henkel en MSC sinds juli 2026).
  - Netto of bruto leeg, bruto ≤ netto, of het netto van een KG-regel ≠ de hoeveelheid.
  - Colli of stuks die geen heel getal zijn.
  - Een LTD QTY-collo boven 30 kg, of LTD QTY met een enkelvoudige verpakking.
  - Een onbekende klasse, een klasse of PG die niet klopt met de IMDG-lijst, of LTD QTY bij PG I.
  - Meerdere UN-nummers in één regel, bloktotalen die niet kloppen, of geen DG-goederen.
  - Een ontbrekend bruto, vak 19 lager dan de DG-som, of een segregatieconflict.
- **PDF:** gebouwd vanaf het ingebedde blanco Henkel-stramien, met een preview-dialoog en opslaan via Enter.
- **CMR** (vanuit deze module): een popup met incoterm (CIF/FOB/FCA), plaats, volgnummer, laaddatum, vervoerderscode, kenteken, barcode (A + BP-nummer), colli, pallets, cubage, container (ISO 6346-checkdigit en OCR), seal, afzender, vervoerder, papiermodus en kalibratie in mm.
- **Wordt een nieuwe Excel of mail ingeladen,** dan wordt alle zendingsgebonden state gewist. Dat is een les: "DAMM" belandde bijna op een Durban-DGD.

---

## 7. Module 06/07 — CMR maken en Transportopdracht

Beide modules gebruiken hetzelfde scherm en dezelfde ingelezen bestanden (`toModus` is `cmr` of `to`).

- **Booking-lezers:**
  - CMA CGM: tabellen ROUTE / CUT-OFF / DEPOT.
  - Oceanwide: Hapag-stijl "Booking Confirmation".
  - MSC: kolomformulier, waarde naast het label.
  - Mailtekst: Vessel/Carrier/ETS/ETA/VGM, datums zonder jaar, Grimaldi-bootnaam als kapitalenregel.

  Een lezer die alles vindt wint; anders de lezer met de meeste velden.
- **Verplichte booking-velden** voor Van der Kaa: bootnaam, closing (CLD = VGM-cutoff), uithaalreferentie en ETD.
- **Zending-Excel:** Master, BP, container ("Plts 21 // 40HC" of "LOOSE LOAD // 20 DV"), IMO/NON IMO, pick-up-adres, UN-nummers, colli (# KA naar boven afgerond), bruto en **cubage**. Bij een palletzending is de cubage palletmaat × pallets, bij losse lading doosmaat × dozen.
- **Meerdere containers in één mail:** de reeksen in het onderwerp ("14050-14051 … 60090021-60090044") worden aan elkaar gekoppeld. De DGD wordt aan een zending gekoppeld via vak 4 (master of shipment), niet via de bestandsnaam.
- **Adresboek van terminals** (`TO_TERMINALS`):
  - Rotterdam: RWG, APM MVII, APM Rotterdam, ECT Delta, Euromax, Hutchison Delta II, Kramer Delta en RCD.
  - Antwerpen: Antwerp Gateway K1700, AET K1333 (Grimaldi), MPET K1742, PSA Noordzee K913 en PSA Europa K869.
- **Transportopdracht-PDF:**
  - Neele-Vat Maasvlakte-kop met logo; ref `60WH/<master>`.
  - Laden en lossen op dezelfde dag; laadadres standaard "Neele Vat Maasvlakte 2, Dardanellenstraat 94".
  - Uithaal (depot) en inlever (terminal) met referentie, bootnaam, CLD, ETD/ETA en instructies (1X 40h / IMO / LOOSE LOAD).
  - Pagina 2 bevat de voorwaarden (Duits minimumloon, CMR/POD binnen 7 dagen naar pod@neelevat.com).
  - Datuminvoer "2606" wordt 26/06/jaar; tijd "10" wordt 10:00. Een landcode wordt de landnaam.
- **Mail naar Van der Kaa:** aan `opdrachten@vanderkaa.com`, cc henkel.dc, onderwerp "TO Henkel <datum> - <master>". Bijlagen: TO, CMR en DGD. HTML in Arial 10 pt, kleur `#266791`.
- **CMR (rit loods → zeeterminal):**
  - Geadresseerde (vak 2) is de **terminal**; afzender is Neele-Vat Maasvlakte; vervoerder is Van der Kaa Transport B.V. (Pernis).
  - Bruto bij voorkeur uit de PL (alleen als alle IDH's gedekt zijn).
  - Een ADR-waarschuwing verschijnt als er UN-nummers zijn zonder DGD.
  - Code 128B-barcode; papiermodus print alleen gegevens (voor voorgedrukt CMR-papier).

---

## 8. Module 08 — Documenten delen (BL of douanedocumenten, IC-orders)

- **Invoer:** de mail plus de draft BL. Het onderwerp wordt exact overgenomen.
- **Bijlagen:** alleen draft BL, invoice, PL en overige PDF's. **Excel gaat nooit mee.** Plaatjes kleiner dan 25 kB (handtekeningen) worden overgeslagen.
- **De soort wordt geraden uit de bestandsnaam**, inclusief BL-prefixen van rederijen (MEDU, MAEU, HLCU, CMDU, COSU, EGLV, OOLU, ONEY, YMLU, HMMU, ZIMU).
- **Modus:** vervoersdocument of douanedocumenten (CVO/EUR-1/legalisatie/EXA). Wordt automatisch herkend en is te overrulen.
- **Vaste ontvangers per bestemming** (`BL_ONTVANGERS`):
  - Durban en Coega: 3 ontvangers (Henkel ZA en DSV).
  - Nhava Sheva / India: 6.
  - Jebel Ali: Henkel Consumer & Industrial.
  - Jeddah: **2 varianten** (Henkel Consumer & Industrial, of Henkel Industrial Company). Hier moet je kiezen.
  - Gebze: 3 (Henkel en Doğruer).
- **Aliassen:** INDIA, NHAVA en JNPT → Nhava Sheva; DUBAI → Jebel Ali.
- **Bij een shipping guide** volgt de tool de adressen uit de guide ("Email address for docs") in plaats van de vaste lijst. Pre-alert-opmerkingen worden getoond.
- Bij EC verschijnt een waarschuwing. De mail kan dan alleen met een vinkje toch worden gemaakt.

## 9. Module 10 — Full set
- **Invoer:** de eerste mail (alleen het onderwerp wordt gebruikt), de scan (pdf of afbeelding) en het trackingnummer.
- **Mail:** "Pfa full set scan. Tracking: …".
- Gebruikt dezelfde ontvangerstabel en hulpfuncties als module 08.

## 10. Module 11 — Ocean boeken
- **Boekingsmail aan Steder** (standaard `Mike.Driest@stedergroup.com`), cc henkel.dc.
- **Uit de Excel:**
  - **Master** wordt Ref.
  - Het blok "Uitslag instructies" (Est payload/DG, plts + cntr).
  - Excel-bestanden waarvan de naam begint met `1000…` (forwarderlijsten) worden genegeerd.
- **Consignee uit de guide:** het blok "Consignee & Notify details in FULL".
- **Bijlagen:** alle bijlagen van de inputmail gaan mee (invoice, PL, labels, MSDS, Excel).
- **Onderwerp:** "OCEAN BOEKEN HENKEL " + het oorspronkelijke onderwerp.
- **Slepen vanuit Outlook Classic:** virtuele bestanden worden synchroon uit `dataTransfer.files` en `items` gehaald. Een bestand zonder extensie maar met Outlook-MIME krijgt `.msg`.

## 11. Module 09 — CVO / EUR.1 draft check
- **Invoer:** de mail met de shipping guide en de customsmail met de draft (of losse .docx/.pdf).
- **Type draft:** "CERTIFICATE OF ORIGIN" → COO; "EUR.1 / MOVEMENT CERTIFICATE" → EUR.1.
- **Eisen komen uitsluitend uit de documentregel van de guide:**
  - Invoice No. Een `xxxx`-placeholder wordt ingevuld met het nummer van de bronfactuur.
  - LC en LC Date.
  - HS importing/exporting country.
  - Country of Origin. Uit andere regels wordt die alleen gebruikt als de guide één eenduidige oorsprong noemt.
  - Description of Goods en letterlijke teksten.
  - Revised Rules (vak 7).
- **Veiligheidsregel:** meerdere guides of meerdere drafts geven nooit een AKKOORD.
- **Status:** AKKOORD, NIET AKKOORD of CONTROLEREN.
- **Antwoordmail aan customs** (Customs.Cluster1, cc Henkel.dc, onderwerp "RE: …"), in het Nederlands, met de punten die ontbreken of afwijken. Bij CONTROLEREN wordt geen mail gemaakt.

## 12. Hoofdmodule 02 — Operations Dashboard (planning & signalering)
- Een aparte React-app (Tailwind via CDN) in een iframe.
- **Invoer:** `New_planning_Henkel.xlsm`, werkblad **"Uit"**: rij 2 bevat de koppen, data begint op rij 3, en kolom BP FILE is verplicht.
- **Opslag:** de data staat in `localStorage["henkel_snapshot_v1"]`. Instellingen staan in `henkel_mail_to` en `henkel_dormant_gap` (standaard 250).
- **Uitsluitingen:** Air-zendingen worden uitgesloten. LOADED = Y betekent afgesloten.
- **Slapend:** het shipment-nummer loopt meer dan 250 achter op de eigen reeks (full ≥10000, air, LCL IC, LCL EC) en er is geen laaddatum.
- **Regels:**
  - **MRN:** laaddatum en MRN leeg, ≤3 werkdagen; op de laaddag of later is het "te laat".
  - **Check planning:** NV OCEAN met laaddatum maar niet geboekt (BOOKED ≠ "*" en ≠ N/A).
  - **Boeken:** NV OCEAN zonder laaddatum en niet geboekt.
  - **Documenten:** EC met laaddatum waarbij CVO/EUR1/ATR/COC/HS leeg is. Rood bij ≤2 werkdagen, oranje bij ≤5.
- **Weergave:**
  - Tegels "te laat", "bijna te laat" en "open".
  - Een overzicht van de komende 5 werkdagen.
  - **Boekmaand van de 25e t/m de 24e.**
  - Zoeken op BP, shipment of master; 15 acties per pagina.
  - "Mail export" (mailto, daily overview).
  - Een knop om de cache te wissen.
- **Diagnose:** `__henkel.inspect("bp")` in de console.

---

## 13. Vaste referentiedata

**IC-documenten per bestemming (`IC_DOCS`):**

| Bestemming | Documenten |
|---|---|
| GEBZE | exa, atr, coo |
| DURBAN, COEGA | exa, eur1 |
| JEBEL ALI | exa, coo, kvk |
| NHAVA SHEVA, CHENNAI, MOMBASA, DAKAR | exa, coo |
| ALEXANDRIA | exa, coo, copy, coc, embassy |
| DAMMAM | exa, coo, kvk, gelegaliseerde factuur |
| JEDDAH | exa, coo, copy, coc |
| CASABLANCA | exa, eur1 |
| BEIRUT | exa, coo, copy, eur1, legalised invoice, embassy |
| TUNIS | exa, coo, eur1 |
| BEJAIA | exa, coo, eur1, coc |

- **Aliassen:** COUGA, GQEBERHA en PORT ELIZABETH → COEGA.
- **Kantoor van uitgang:** TUNIS = FR002730; GEBZE = LU701000, behalve bij SEA (dan staat het in het onderwerp).
- **IC krijgt geen draft**, behalve **TUNIS** en **CASABLANCA**.
- **EUR-1-retouradres:** alleen voor Casablanca (Espace Transit Tanger).
- **DGD-consignees:** Gebze, Nhava Sheva, Durban/Coega, Jebel Ali, Mombasa, Dammam/Jeddah (Riyadh), Casablanca en Tunis. Een volgnummer achter de bestemming ("Durban 1") wordt genegeerd.
- **Henkel noodnummer:** 24/7 00492117973350.
- **IMDG-lijst** in de tool: UN 1133, 1139, 1263, 1866, 1950, 2491, 2735, 2865, 3082, 3077, 3175, 3264, 3265, 3266 en 3267 (klasse, toegestane PG's en de LTD-maximumhoeveelheid).

## 14. Praktijklessen (zendingnummers in de code)

| Zending | Les |
|---|---|
| 14010 | COSCO weigerde een niet-geboekt UN-nummer. Een extra UN op de DGD geldt als "DG toegevoegd na boeking". |
| 14035 | Palletgewicht hoort per IDH/HS op de EXA/COO/EUR1. Maximaal 9 IDH's per certificaat. |
| 14076 | Douane inputblad: HS-, gewicht- en colli-normalisatie. |
| 14078 | 8 ST van 964 kg zijn IBC's, geen dozen. Gebruik het gemeten netto, niet 90%. |
| 14092 | Drums met LTD QTY worden afgekeurd (maximaal 30 kg). |
| 14097 / 14107 | MSC eist een vlampunt op aerosol-regels. Netto 0 wordt afgekeurd. |
| 14101 | Zonder zending-Excel ontbrak het Inputblad; het wordt nu gereconstrueerd. |
| 14118 | Een MSDS met 2 stoffen: beide moeten op de DGD. |
| 14141 (Tunis) | De Excel-naam belandde op de EXA. De factuurnaam is leidend. |
| 14142 | CMA eist het percentage salpeterzuur. |
| 14147 / 14177 | Combi-artikel EA 9466 splitsen per UN (Steder, Maersk). |
| 14186 (Balchem) | Invoice- en PL-condities uit de guide horen niet in de aanvraag. |
| 60090018 (Lubumbashi) | Een BEANR-booking moet in Antwerpen worden afgeleverd, niet bij RWG. |
| S330146158 | Grimaldi gebruikt AET Kaai 1333. |

---

## 15. Aandachtspunten / mogelijke fouten die bij het lezen opvielen

1. **De intro van DGD maken v2 klopt niet meer.** De intro en het v2-kopcommentaar zeggen dat netto altijd 90% van bruto is. De code gebruikt de kg-hoeveelheid (KG-regels), dan het gemeten PL-netto, en pas als terugval 90%.
2. **`instructions` op de CMR wordt nooit getekend.** `toBuildCmrModel` vult `cmr.instructions` (eindklant + booking, "vak 13"), maar `cmrBuildPdf` tekent dat veld niet. De eindontvanger staat dus niet op de CMR.
3. **Het "on behalf of"-blok staat vast op Henkel Jebel Ali.** In vak 22 is het standaard Henkel Jebel Ali FZCO, en de afzender in DGD-maker-CMR's is dat ook. Voor andere bestemmingen of entiteiten moet dat met de hand worden aangepast.
4. **Vak 4 van de CMR staat vast** op "MAASVLAKTE/ROTTERDAM", ook als de incoterm-plaats Antwerp is.
5. **Veld "Totaal incl. pallets" verwacht NL-notatie.** Het veld haalt alle punten weg. Wie "2392.294" typt, krijgt 2.392.294 kg.
6. **`toBuildEml` (de transportopdracht-mail)** codeert het onderwerp niet volgens RFC 2047 en zet elke bijlage als `application/pdf`. `buildEml` doet beide wel goed.
7. **De ref600-terugval in `maakReadExcel`** zoekt `r.order`, maar dat veld wordt nooit gevuld. De terugval werkt dus niet.
8. **De `st._doel === 'cmr'`-tak in het transportopdracht-formulier** is onbereikbaar geworden: de CMR heeft nu een eigen dialoog.
9. **Validaties in het Dashboard doen niets.** `VALIDATIONS` is leeg (de TR-zonder-ATR-validatie is verwijderd), dus het tabblad "Validaties" toont altijd 0. Het kopcommentaar noemt "slapend > 200", terwijl de standaard 250 is.
10. **Nummering en titels lopen niet gelijk.**
    - Het Dashboard heeft "09" als tabnummer, net als de CVO-check.
    - De commentaarkoppen gebruiken andere nummers (BL "08", Full set "09", Ocean "10").
    - `NAMEN` (de browsertitel) mist `fs` en `ob`.
11. **De footer is onvolledig.** Hij noemt de bladen Douane inputblad, Pallets en Certificaten niet, en de bibliotheekmelding noemt ExcelJS "de controle-Excel bouwen", terwijl ExcelJS alleen voor het klantrapport wordt gebruikt.
12. **De contactpersoon van de transportopdracht (`TO_CONTACT`) staat vast in de code** (naam en e-mail). Een collega kan alleen de naam per keer overschrijven, het e-mailadres in de kop niet.
13. **Het Dashboard zoekt een kolom die niet meer bestaat.** Het leest de laaddatum uit kolom `LD DATE`, maar in de huidige `New_planning_Henkel.xlsm` heet die kolom **`LAAD DATUM`**.
    - Daardoor heeft geen enkel dossier een laaddatum: de regels MRN, Check planning en Documenten gaan nooit af.
    - Op de data van 10-11-2026 toont het dashboard 15 acties in plaats van 28, en 0 in plaats van 7 te late acties.
    - Zie `KENNISBANK_Planning_en_Documentregistratie.md`, A6.
