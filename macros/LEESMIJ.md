# Macro: Documentregistratie_Aanvullen

Vult op het blad van de documentregistratie de kolommen **BP File, Reference, IC/EC, Dest en Loading date** aan.
De gegevens komen uit blad **Uit** van `New_planning_Henkel.xlsm`, opgezocht op het masternummer (kolom **Master**).

| Registratieblad | Bron (blad Uit) |
|---|---|
| BP File | BP FILE |
| Reference | SHIPMENT |
| IC/EC | EC/IC |
| Dest | DEST |
| Loading date | LAAD DATUM (of LD DATE) |

Kolommen A, G, H en I (Master, Document code, Description, Request date) komen uit de tool, met de knop "Kopieer voor documentregistratie".

## Installeren (eenmalig)
1. Open de werkmap met het registratieblad en druk op **Alt+F11**.
2. Kies **Bestand → Bestand importeren…** en selecteer `Documentregistratie_Aanvullen.bas`.
3. Sla de werkmap op als **.xlsm**.

## Gebruik
1. Plak de regels uit de tool op het registratieblad.
2. Laat het registratieblad actief en druk op **Alt+F8** → `Documentregistratie_Aanvullen` → Uitvoeren.
3. Staat de planning niet open, dan vraagt de macro om het bestand. Hij opent het alleen-lezen.

## Gedrag
- Alleen **lege** cellen worden gevuld. Zet `OVERSCHRIJVEN = True` bovenaan de module om alles opnieuw te vullen.
- Masters die niet in de planning staan, worden **lichtrood** gemarkeerd en in de melding genoemd.
- Komt een master meerdere keren voor in de planning, dan wint de nog niet geladen regel, daarna de laatste laaddatum. De melding noemt die masters.
- Placeholderdatums (bijvoorbeeld 01-01-1900) worden niet overgenomen. Loading date krijgt de opmaak dd-mm-jjjj.
- Kolommen worden op naam gevonden (rij 1 van het registratieblad, en de koprij van tabel `HenkelPlanning`), dus de volgorde mag verschuiven.
