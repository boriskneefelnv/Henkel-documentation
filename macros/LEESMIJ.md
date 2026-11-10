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
| `Documentregistratie_Aanvullen.bas` | De VBA-macro (Excel desktop), voor wie die liever gebruikt |
| `../tool/Documentregistratie_bijwerken.html` | **HTML-tool** in de browser, zonder Excel-script of macro |

## HTML-tool gebruiken
1. Open `tool/Documentregistratie_bijwerken.html` in Chrome of Edge.
2. Kies de documentregistratie (.xlsx met tabblad DCMAASVLAKTE) en de Henkel-planning (.xlsx/.xlsm/.xls met blad Uit). Staat blad Uit al in de documentregistratie, dan is de planning niet nodig.
3. Klik op **Controleer wijzigingen**. Je ziet per cel wat er was en wat het wordt, plus de niet-gevonden en dubbele masters.
4. Klik op **Download bijgewerkte documentregistratie**. Je krijgt een nieuwe kopie (`… - bijgewerkt.xlsx`). Het origineel blijft ongewijzigd.

De planning en de documentregistratie mogen hier dus in **verschillende werkmappen** staan. Alles blijft lokaal in de browser. Alleen de Excel-bibliotheken worden bij het openen van cdnjs geladen, dus je hebt internet nodig.
Let op: de kopie wordt opnieuw opgebouwd. Waarden en celopmaak blijven staan, maar draaitabellen, grafieken en macro's in de documentregistratie kunnen daarbij verloren gaan. Gebruik in dat geval het Excel-script.

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

Het script start direct, zonder invoerscherm. Staat er nog een oude versie in Excel die om "planning" vraagt? Vervang dan de hele code door de nieuwe versie.
