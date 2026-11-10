# Macro: Documentregistratie_Aanvullen

Werkt **alleen op tabblad `DCMAASVLAKTE`**. Hij werkt de groene kolommen bij vanuit blad **Uit** van `New_planning_Henkel.xlsm`, op basis van het masternummer (kolom **Master**).

| Groene kolom op DCMAASVLAKTE | Bron (blad Uit) |
|---|---|
| BP File | BP FILE |
| Reference | SHIPMENT |
| IC/EC | EC/IC |
| Dest | DEST |
| Loading date | LAAD DATUM (of LD DATE) |

Uit de tool komen, met de knop "Kopieer voor documentregistratie": Master, 5 witte kolommen, Document code, Description en Request date (kolom A, G, H en I).

## Regels
- **Alleen tabblad DCMAASVLAKTE wordt aangepast**, en daarop alleen de 5 groene kolommen. Andere kolommen (Master, Document code … Status), andere tabbladen en de planning zelf blijven onaangeroerd: geen waarden en geen opmaak.
- **De groene kolommen volgen altijd blad Uit.** Is iets in de planning veranderd (laaddatum, shipment, …), dan wordt de groene cel bij de volgende run aangepast. Is het veld in de planning leeg geworden, dan wordt de groene cel ook leeg. De opmaak blijft staan.
- Staat een master niet in de planning, dan blijft die regel helemaal ongewijzigd. Hij wordt alleen in de melding genoemd.
- Komt een master meerdere keren voor in de planning, dan wint de nog niet geladen regel, daarna de laatste laaddatum. De melding noemt die masters.
- Placeholderdatums (bijvoorbeeld 01-01-1900) gelden als "geen laaddatum".
- Kolommen worden op naam gevonden, dus de volgorde mag verschuiven.

## Installeren (eenmalig)
1. Open de werkmap met tabblad DCMAASVLAKTE en druk op **Alt+F11**.
2. Kies **Bestand → Bestand importeren…** en selecteer `Documentregistratie_Aanvullen.bas`. Staat er al een oude versie in, verwijder die module dan eerst.
3. Sla de werkmap op als **.xlsm**.

## Gebruik
- Druk op **Alt+F8** → `Documentregistratie_Aanvullen` → Uitvoeren. Dat kan vanaf elk tabblad: de macro zoekt DCMAASVLAKTE zelf.
- Staat de planning niet open, dan vraagt hij om het bestand en opent hij het alleen-lezen.
- Draai de macro opnieuw wanneer de planning is bijgewerkt. Zo blijven de groene kolommen actueel.
