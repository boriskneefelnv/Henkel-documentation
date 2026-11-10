/**
 * Office Script voor het PLANNINGBESTAND (New_planning_Henkel).
 * Alleen nodig als DCMAASVLAKTE in een ANDER bestand staat dan blad Uit.
 *
 * Geeft de kolommen die de documentregistratie nodig heeft terug als tabel (array, als tekst;
 * datums als Excel-datumgetal).
 * In Power Automate:
 *   1. "Script uitvoeren" op het planningbestand met dit script.
 *   2. "Script uitvoeren" op het registratiebestand met Documentregistratie_Aanvullen,
 *      en vul bij de parameter "planning" de uitkomst (result) van stap 1 in.
 * Dit script leest alleen; het verandert niets aan de planning.
 */

type Cel = string | number | boolean;

function main(workbook: ExcelScript.Workbook): string[][] {
  let data: Cel[][] | undefined;
  const t = workbook.getTable("HenkelPlanning");
  if (t) data = t.getRange().getValues() as Cel[][];
  else {
    const ws = workbook.getWorksheets().find(w => w.getName().trim().toUpperCase() === "UIT");
    const used = ws ? ws.getUsedRange(true) : undefined;
    if (ws && used) {
      const alles = ws.getRangeByIndexes(0, 0, used.getRowIndex() + used.getRowCount(),
        used.getColumnIndex() + used.getColumnCount()).getValues() as Cel[][];
      const kopRij = alles.findIndex((r, i) => i < 10 && r.some(c => norm(c) === "MASTER") && r.some(c => norm(c) === "BPFILE"));
      if (kopRij >= 0) data = alles.slice(kopRij);
    }
  }
  if (!data) throw new Error("Tabel HenkelPlanning / blad Uit niet gevonden in dit bestand.");

  const nodig = [["MASTER"], ["BP FILE"], ["SHIPMENT"], ["EC/IC", "IC/EC"], ["DEST"], ["LAAD DATUM", "LD DATE", "LD. DATE"], ["LOADED"]];
  const idx = nodig.map(namen => data![0].findIndex(c => namen.map(norm).indexOf(norm(c)) >= 0));
  return data
    .filter((r, i) => i === 0 || String(r[idx[0]] ?? "").trim() !== "")
    .map((r, i) => idx.map((c, j) => (i === 0 ? nodig[j][0] : (c >= 0 ? String(r[c] ?? "") : ""))));
}

function norm(v: Cel): string {
  return String(v ?? "").replace(/[\s ]/g, "").toUpperCase();
}
