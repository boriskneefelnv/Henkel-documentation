# Kennisbank — Henkel-planning (xlsm) en DC Maasvlakte-documentregistratie (csv)

Bronnen:
- `New_planning_Henkel.xlsm` (19,9 MB)
- `DCMAASVLAKTEDocumentregistratie.csv` (17,9 MB)

Peildatum van de analyse: 10-11-2026.
Zie ook `KENNISBANK_Henkel_Operations_Center.md` (de tool die deze planning inleest).

---

# A. New_planning_Henkel.xlsm

## A1. Opbouw

| Blad | Status | Inhoud |
|---|---|---|
| **Uit** | zichtbaar | De planning: Excel-tabel **`HenkelPlanning`** (A2:BB1041), 1.038 dossiers × 54 kolommen |
| Overview | zichtbaar | Dynamische formules: laaddatums vanaf vandaag met aantal dossiers en de som van ORL per dag |
| CCMS new | zichtbaar | Klachtenanalyse (CCMS-notificaties), 113 regels × 28 kolommen, met formules naar een extern werkboek |
| Portbase | zichtbaar | `FILTER` van de dossiers met LAAD DATUM = vandaag én CARR = NV OCEAN (BP, master, shipment, PR, MRN, container, booking) |
| Luchtvracht | zichtbaar | Klein lijstje namen met getallen (vermoedelijk uren/kosten per medewerker), plus één BP-nummer |
| CCMS List | zichtbaar | CCMS-register, 103 regels: nummer, status, comment, root cause, defect type, carrier, datum |
| Fysieke docs | **verborgen** | Invoices waarvoor fysieke ATR/EUR.1 nodig waren (Turkije, Algerije, Zuid-Afrika, Marokko, Libanon) |
| Bijzonderheden afrekenen | **verborgen** | IC-documentenmatrix per bestemming, plus een to-do-lijst |
| Blad2 | **verborgen** | Notities over de betekenis van de veldwaarden |
| Blad1 | **verborgen** | Een oudere versie van de planning (mei 2026, andere kolomindeling). Het XML-bereik loopt tot rij 1.044.721, waardoor het bestand onnodig groot is (166 MB uitgepakt) |

Overige onderdelen van het bestand:
- **VBA:** 3 macro's (zie A5).
- **Opmerkingen:** ~58 threaded comments, vooral laadtijden ("08:00", "10:00") en notities.
- **Opmaak:** 358 regels voor voorwaardelijke opmaak en 24 validatielijsten.
- **Externe koppeling:** naar `CCMS_PA_updated (002).xlsx`, een lokaal Outlook-cachepad op de pc van een collega.

## A2. Kolommen van `HenkelPlanning` (blad Uit, kop op rij 2)

Rij 1 bevat drie groepskoppen: **PLANNING** (vanaf C), **SEAFREIGHT** (vanaf AE) en **WAREHOUSE** (vanaf AQ).

| Groep | Kolommen |
|---|---|
| Identificatie | `BP FILE` (609…), `MASTER` (600…), `SHIPMENT` |
| Planning | `LAAD DATUM`, `PR`, `PB`, `LOADED` (Y = afgesloten), `DG` (Y/N), `EC/IC`, `DEST` (landcode), `TYPE` (Sea/Air/Rail/Truck), `LOS` (losse lading), `CARR` (NV OCEAN/NV AIR/NV TRANS/EXTERNAL), `INC` (incoterm), `ANT` (leeg), `BOOKED` (`*` = geboekt), `REMARKS`, `OPCNTR` (formule), `ORL`, `HF` |
| Documenten | `CVO`, `EUR1`, `ATR`, `COC`, `HS`, `BL`, `MRN`, `DHL REFERENCE` |
| Zeevracht | `SEAL`, `CNTR NR`, `CNTR TYPE`, `ETD`, `ETA`, `CARRIER`, `BOOKING NUMBER`, `B/L NUMBER`, `BC CONFIRM`, `VESSEL`, `POL`, `POD` |
| Warehouse (tellers) | `PLT`, `EUR`, `BP`, `AIR`, `WOOD`, `ADR LBL`, `OWN PL`, `SEAL2`, `UREN`, `LOS2` (formule), `LBL`, `AFG` |
| Hulpkolommen | `FILTER_UITSLUITEN`, `_FILTER_STATUS` (TONEN/VERBERGEN) |

Formules in de tabel:
- `OPCNTR` telt het aantal containernummers in REMARKS: leeg = 0, `*` = 1, anders het aantal komma's + 1.
- `LOS2`: LOS = Y wordt 1, N wordt 0.

**Betekenis van de waarden** (uit Blad2, de commentaren en de data):
- `*` betekent "geregeld" bij PR, PB en BOOKED. NVT en N/A betekenen niet van toepassing.
- In de documentkolommen CVO, EUR1, ATR en COC betekent **A** "aangevraagd", **0** "niet nodig" en een getal waarschijnlijk een aantal of stand.
- In de kolom HS staat het aantal extra HS-codes. De tool kopieert precies deze regel (A/0 + HS) naar het klembord.
- BL: "F = gestuurd naar Henkel".
- `DHL REFERENCE` bevat een koeriertracking, of de tekst "FULL SET SENDED".
- MRN bevat een 18-tekens Douane-MRN (`26NL…`), of teksten als "Applicable, see mail", "gevraagd", "Opgevraagd", "binnen" en "NVT".
- Commentaar bij `ORL`: "**ORL CAP VAN 60 per dag.** Oranje en vanaf 80 rood". `MASTER` geldt ook voor de loading ref.

**Validatielijsten:**
- `LOADED`, `HF`, `LOS`: Y,N.
- `EC/IC`: EC,IC.
- `TYPE`: Sea,Air,Rail,Truck.
- `CARR`: NV OCEAN, NV AIR, NV TRANS, EXTERNAL.
- `INC`: CFR,CIF,CPT,DAP,FCA,FOB.
- `CARRIER`: MSC, HAPAG, CMA CGM, COSCO, EVERGREEN, SSC, FPS, SCI, OOCL, YANG MING, HMM, DIAMOND, BURGER, N/A.
- `CNTR TYPE`: N/A, 20ft, 40ft.
- POL/POD: een lijst met havens.
- `DEST`: een lijst met ISO-landcodes.

**Voorwaardelijke opmaak** (hoofdregels):
- Een rij wordt gekleurd bij `LOADED="Y"`, `LOS="Y"` of `HF="Y"`.
- BP FILE die niet met 609 begint, en MASTER die niet met 600 begint, wordt gemarkeerd. Dubbele BP's en masters ook.
- De laaddatum wordt gemarkeerd als die vandaag is.
- Een MRN die niet 18 tekens is (en niet "Applicable, see mail") wordt gemarkeerd.
- BOOKED wordt gemarkeerd als het leeg is bij NV OCEAN.

## A3. Stand van de data (10-11-2026)

- **Omvang:** 1.038 dossiers.
  - 977 zijn afgesloten (LOADED = Y; 78 daarvan met een kleine "y").
  - **61 zijn open:** 39 Sea, 16 Truck, 4 Rail en 2 Air. 20 daarvan zijn NV OCEAN en 36 EXTERNAL.
- **Laaddatums van de open dossiers:** 36 vandaag of later, 25 leeg, geen enkele in het verleden.
- **Verdeling:** 678 IC en 338 EC.
  - Bestemmingen: IN 232, ZA 126, TR 104, AE 81, SA 59, EG 53, TN 51, MA 50, PK 37.
  - Type: Sea 562, Air 224, Truck 164, Rail 73.
  - Rederijen: MSC 147, CMA CGM 61, Hapag 42, Maersk 32.
- **Placeholderdatums:** 82× 01-01-2026 en 19× 01-01-1900 in LAAD DATUM.
- **Datakwaliteit:**
  - 3 rijen met BP FILE "x".
  - 9 BP-nummers beginnen met 600 in plaats van 609 (60095394–60095424).
  - Dubbele BP's: 60927344 en 60927508. Er zijn 5 dubbele masters.
  - EC/IC bevat foute waarden: "Y", "1", "TR", "MA", "ZA". Bij de laatste drie staan EC/IC en DEST verwisseld (DEST = "IC").
  - In CNTR TYPE staan vrije teksten die de validatielijst omzeilen: "320ft", "6 x 20ft", "LCL", " 20ft".
  - VESSEL/POL/POD bevatten soms complete geplakte zinnen ("ROTTERDAM / RWG - ROTTERDAM ETA: 16 JUL 2026", "Remarks").
  - MRN staat meerdere keren in één cel (komma of slash).
- **ORL-belasting per laaddag:** gemiddeld 69 en mediaan 65.
  - 83 laaddagen zitten op of boven de cap van 60, en 63 dagen op of boven 80 (rood). Het maximum is 189.
  - De datum 01-01-2026 (placeholder) telt ook mee, met 185.

## A4. Klachten (CCMS)

- **CCMS new** (januari–oktober 2026, 113 notificaties):
  - Defect: 73× under shipment, 17× containerschade, 12× over shipment, 5× verkeerd product.
  - Status: 56 Invalid, 44 Valid, 13 Unverifiable.
  - Carrier: Neele-Vat 56, Intermax 28, AST/Avenir 12, DHL Medhub 7.
  - Root cause: 35× "full quantity loaded" (invalid), 16× consolidatie- of picking-fout, 16× vermoedelijk weg- of douane-inspectie, 12× short loading niet gemeld aan Henkel, 10× schadevrij geladen (foto's).
  - Bestemmingen: Durban 22, Jebel Ali 20, Tunis 17, Casablanca 12, Gebze 11, Jeddah 11. 101 van de 113 zijn DG.
  - De formules koppelen aan het externe bestand (Root_Cause_Map, Summary, YTD_2026) voor de acties van het Rotterdam-bezoek van 1 juli 2026 en een voor/na-vergelijking. Zonder dat bestand tonen deze kolommen alleen de laatst opgeslagen waarden.
- **CCMS List** (103 regels): 43 Closed Not Valid, 31 Closed Valid, **20 Open**, 9 Closed Unverifiable. Het meest genoemde Neele-Vat-commentaar is "We sent full ka".
- **Actie uit de externe koppeling:** de picklijst toont nu het aantal pallets. Het warehouse meldt ontbrekend materiaal direct aan Henkel en Henkel beslist (WH-02, picklist-stappen 6–8). Daarnaast verstuurt CargoSnap automatisch een mail bij een missing-stock-rapport.

## A5. VBA-macro's (Module1–3)

1. **`Filter_HenkelPlanning`**:
   - Herstelt de weergave.
   - Sorteert op **`LD DATE`**.
   - Verbergt rijen met LOADED = Y waarbij alle warehouse-tellers (PLT … LBL) én DHL REFERENCE gevuld zijn.
2. **`Toon_Alles_HenkelPlanning`**: toont alles weer en wist de gele markering.
3. **`Controle_Lege_Verplicte_Kolommen_HenkelPlanning`**:
   - Toont alleen de verplichte kolommen (BP FILE, MASTER, SHIPMENT, CVO/EUR1/ATR/COC en de warehouse-tellers).
   - Laat alleen rijen zien met een lege verplichte cel en een lege AFG. Die cellen worden geel.

Geen verdachte code: de "suspicious"-meldingen van olevba zijn vals alarm, het zijn de kolomnamen.

## A6. Belangrijkste bevinding: kolom `LD DATE` heet nu `LAAD DATUM`

De kolom is hernoemd van `LD DATE` naar **`LAAD DATUM`**.
- De formules op Overview en Portbase zijn meegegaan.
- De foutmelding in de Portbase-formule zegt nog "Geen regels met LD DATE = vandaag".
- Twee onderdelen zijn **niet** meegegaan:
  1. **Macro `Filter_HenkelPlanning`**: zoekt `LD DATE`. De macro stopt met de melding "kolommen niet gevonden: LD DATE".
  2. **Operations Dashboard in de tool**:
     - Het dashboard leest `idx["LD DATE"]`, dus elk dossier krijgt **geen laaddatum**.
     - Daardoor gaan de regels MRN, Check planning en Documenten nooit af. Alleen "Boeken" werkt, en elke actie valt in de lijst "Zonder laaddatum".
     - Ook de dagtegels en de boekmaandfilter zijn leeg.
     - De headercommentaar in het dashboard ("LD DATE …") is daardoor ook verouderd.

Simulatie met de huidige data (peildatum 10-11-2026, Air uitgesloten):

| | Zoals nu | Met `LAAD DATUM` gekoppeld |
|---|---|---|
| MRN | 0 | 12 |
| Check planning | 0 | 11 |
| Boeken | 15 | 4 |
| Documenten | 0 | 1 |
| Te laat / bijna / open | 0 / 0 / 15 | 7 / 9 / 12 |

Dit is op te lossen door de kolom in het Excel-bestand terug te hernoemen, of door in het dashboard (en de macro) beide namen te accepteren.

---

# B. DCMAASVLAKTEDocumentregistratie.csv

## B1. Opbouw

- Puntkomma-gescheiden, 16 kolommen.
- **615 echte regels en 1.047.937 lege regels** (`;;;;…`): een export van een volledig Excel-blad, vandaar de 17,9 MB. 309 dossiers.
- **Kolommen:** Dossiernummer (601…), Referentie (order van de klant), Klant, Lossen (losadres-code), Los land, Laaddatum, Documentcode, Omschrijving, Werkdagen voor laden, Uiterste aanvraagdatum, Aanvraagdatum, Datum ontvangst, Datum documenten klant gestuurd, Tracking nummer, Remark, Status.
- **Bevat geen Henkel-zendingen.** Dit is de documentregistratie van DC Maasvlakte voor andere klanten:

| Klant | Dossiers | Regels | Referentieformaat |
|---|---|---|---|
| Solae Europe SARL | 131 | 238 | 811xxxxx |
| Azelis International SAL Offshore | 82 | 203 | SO365xxxxxx |
| Roquette Norway AS | 44 | 70 | 555xxxx |
| Barentz BV | 34 | 65 | NL-SO-xxxxxx |
| International N&H Manufacturing UK | 11 | 27 | 53xxxxx |
| International N&H Mfg. Ireland | 7 | 12 | 555xxxx |

- **Periode:** laaddatums van 01-10-2026 tot 30-11-2026; aanvragen van 28-08 tot 09-11-2026.

## B2. Documentcodes en aanvraagtermijn

"Werkdagen voor laden" bepaalt de uiterste aanvraagdatum: laaddatum min dat aantal werkdagen.

| Code | Document | Werkdagen vóór laden | Aantal | Doorlooptijd aanvraag → ontvangst (mediaan, kalenderdagen) |
|---|---|---|---|---|
| 71001 | EXA | 2 | 282 | 1 |
| 71200 | T-1 | 1 | 26 | 1 |
| 72001 | ATR | 4 | 11 | 9 |
| 72002 | EUR.1 | 2 (soms 1) | 78 | 7 |
| 72200 | CVO | 2 (soms 1) | 113 | 3 |
| 72301 | Legaliseren | 2 | 1 | 1 |
| 72302 | Free Sale-certificaat | 5 | 4 | 13 |
| 72400 | NVWA (fytosanitair/export-certificaat) | 4 | 97 | 4 |
| 72401 | COKZ | 4 | 3 | 10 |

- **Termijn te krap:** de doorlooptijden van EUR.1 (mediaan 7, tot 22 dagen), ATR, NVWA, Free Sale en COKZ zijn duidelijk langer dan de termijn van 2 of 4 werkdagen.
  - 63 documenten kwamen pas **na** de laaddatum binnen (30 NVWA, 11 CVO, 10 EUR.1).
  - 131 regels zijn **na** de uiterste aanvraagdatum aangevraagd. Dat zijn vooral EXA (63) en NVWA (34), het vaakst bij Solae (55) en Roquette (34).
  - In de remarks staat vaak "aanvullende services te laat ingevuld door KB".
- **Laaddatum op zaterdag:** voor 31-10-2026 rekent het bronsysteem de uiterste datum anders. Bij 2 werkdagen wordt dat 29-10 (do), terwijl werkdagen tellen vanaf de zaterdag 28-10 zou geven.
- **Pakketten per dossier:** EXA alleen (84), EXA+NVWA (51), CVO+EXA (40), CVO+EUR.1+EXA (39), EUR.1+EXA (16), T-1 alleen (14).
- **Per land:** Jamaica en Nicaragua krijgen EXA + NVWA (Solae). Algerije en Egypte krijgen CVO + EUR.1 + EXA (Azelis). Turkije krijgt ATR + CVO + EXA. VS-zendingen hebben alleen een EXA.

## B3. Statusmodel

| Status | Aantal | Kenmerk |
|---|---|---|
| Nog aanvragen | 37 | Geen aanvraagdatum. 25 EXA, 5 NVWA, 3 T-1. Het meest bij Azelis met FCA (de koper regelt de export) |
| Aangevraagd - wachten | 91 | Aanvraagdatum gevuld, ontvangst leeg. 29 CVO, 24 EUR.1, 20 NVWA. **13 daarvan liggen al voorbij de laaddatum** |
| Ontvangen - naar klant | 12 | Ontvangen, nog niet doorgestuurd |
| Afgerond | 475 | |

- EXA wordt vrijwel nooit fysiek doorgestuurd: "documenten klant gestuurd" is meestal N/A.
- NVWA, EUR.1, CVO, ATR en Free Sale gaan per koerier. Tracking van 12 cijfers (87…) is vermoedelijk FedEx; tracking van 10 cijfers is vermoedelijk DHL.

**Gangbare remarks:**
- FCA, "wacht op reactie klant (ON HOLD)", "wacht op LTSD" (Egypte), "luchtvracht wacht op KVU".
- "ACID: … op de docs laten zetten" (Egypte).
- "NIET MOGELIJK GOEDEREN UIT US", "GEEN T-1 NODIG", "WACHTEN MET AANVRAGEN – SEA".

**Datakwaliteit:**
- Geannuleerde dossiers staan als "Afgerond" met "CANCEL" in de referentie.
- Gesplitste orders hebben de suffix " A" of " B".
- De laaddatum **30-11-2026** komt 78 keer voor (77× Azelis). Dat is vrijwel zeker een placeholder (einde maand) en geen echte laaddatum.
