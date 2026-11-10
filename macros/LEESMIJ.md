# Documentregistratie bijwerken

Werkt **alleen op tabblad `DCMAASVLAKTE`**. Het script werkt de groene kolommen bij vanuit blad **Uit** van de Henkel-planning, op basis van het masternummer (kolom **Master**).

| Groene kolom op DCMAASVLAKTE | Bron (blad Uit) |
|---|---|
| BP File | BP FILE |
| Reference | SHIPMENT |
| IC/EC | EC/IC |
| Dest | DEST |
| Loading date | LAAD DATUM (of LD DATE) |

Uit de tool komen, met de knop "Kopieer voor documentregistratie": Master, 5 witte kolommen, Document code, Description en Request date.

Er zijn twee versies met hetzelfde gedrag:

| Bestand | Voor |
|---|---|
| `Documentregistratie_Aanvullen.osts.ts` | **Excel-script** (Office Script, Excel voor het web / Microsoft 365, tabblad *Automatiseren*) |
| `Planning_Exporteren.osts.ts` | Hulpscript, alleen nodig als DCMAASVLAKTE en blad Uit in **verschillende bestanden** staan (via Power Automate) |
| `Documentregistratie_Aanvullen.bas` | De VBA-macro (Excel desktop), voor wie die liever gebruikt |

## Regels (beide versies)
- **Alleen tabblad DCMAASVLAKTE wordt aangepast**, en daarop alleen de 5 groene kolommen. Andere kolommen, andere tabbladen en de planning zelf blijven onaangeroerd: geen waarden en geen opmaak.
- **De groene kolommen volgen altijd blad Uit.** Wat in de planning verandert, wordt bij de volgende run bijgewerkt. Is een veld in de planning leeg geworden, dan wordt de groene cel ook leeg. De opmaak blijft staan.
- Staat een master niet in de planning, dan blijft die regel ongewijzigd. Hij wordt alleen gemeld.
- Dubbele masters in de planning: de nog niet geladen regel wint, daarna de laatste laaddatum.
- Placeholderdatums (bijvoorbeeld 01-01-1900) gelden als leeg.
- Kolommen worden op naam gevonden, dus de volgorde mag verschuiven.

## Excel-script installeren en gebruiken
1. Open de werkmap in Excel (web of desktop met Microsoft 365) → tabblad **Automatiseren** → **Nieuw script**.
2. Vervang de voorbeeldcode door de inhoud van `Documentregistratie_Aanvullen.osts.ts` en sla op als *Documentregistratie aanvullen*.
3. Klik op **Uitvoeren**. De melding (aantal bijgewerkte regels, niet-gevonden en dubbele masters) staat in het uitvoervenster.
   Tip: via **…** → *Knop toevoegen* zet je het script als knop op een werkblad.

**Belangrijk:** een Excel-script kan alleen de werkmap lezen waarin het draait. Blad **Uit** (tabel `HenkelPlanning`) moet dus in **dezelfde werkmap** staan als DCMAASVLAKTE.

### Staat de planning in een ander bestand? Gebruik Power Automate
1. Maak een flow met **Excel Online (Business) → Script uitvoeren** op het planningbestand, met het script `Planning_Exporteren` (opslaan zoals hierboven).
2. Voeg een tweede **Script uitvoeren** toe op het registratiebestand met `Documentregistratie aanvullen`. Vul bij de parameter **planning** het *result* van stap 1 in.
3. De flow kun je handmatig starten of laten inplannen (bijvoorbeeld elk uur).

Power Automate werkt met bestanden in OneDrive of SharePoint. Draait het script daar niet op een .xlsm-bestand, sla de werkmap dan op als .xlsx (de VBA-macro gaat daarbij verloren).
