/**
 * Office Script: Documentregistratie bijwerken vanuit de Henkel-planning
 *
 * Werkt UITSLUITEND op tabblad "DCMAASVLAKTE" (kop op rij 1):
 *   Master | BP File | Reference | IC/EC | Dest | Loading date | Document code | ...
 *
 * Per regel wordt het masternummer (kolom Master) opgezocht in de planning en worden
 * de vijf GROENE kolommen bijgewerkt, ook als er al iets in staat, zodat een wijziging
 * in blad Uit altijd doorkomt:
 *   BP File      <- BP FILE
 *   Reference    <- SHIPMENT
 *   IC/EC        <- EC/IC
 *   Dest         <- DEST
 *   Loading date <- LAAD DATUM (of LD DATE)
 *
 * Alle andere kolommen, tabbladen en de planning zelf worden niet aangeraakt.
 * Een master die niet in de planning staat, blijft ongewijzigd (alleen gemeld).
 * Dubbele masters in de planning: de nog niet geladen regel wint, daarna de laatste
 * laaddatum. Placeholderdatums (jaar < 2000, zoals 01-01-1900) gelden als leeg.
 *
 * Bron: tabel "HenkelPlanning" / blad "Uit" in DEZELFDE werkmap.
 * Het script start direct, zonder invoerscherm.
 */

type Cel = string | number | boolean;

const DOEL_BLAD = "DCMAASVLAKTE";
const BRON_BLAD = "Uit";
const BRON_TABEL = "HenkelPlanning";
const DATUM_GRENS = 36526; // 01-01-2000 als Excel-datumgetal

function main(workbook: ExcelScript.Workbook) {
  // --- alleen tabblad DCMAASVLAKTE ---
  const doel = workbook.getWorksheets().find(ws => ws.getName().trim().toUpperCase() === DOEL_BLAD);
  if (!doel) return meld(`Tabblad '${DOEL_BLAD}' niet gevonden. Er is niets aangepast.`);

  const used = doel.getUsedRange(true);
  if (!used) return meld(`Tabblad '${DOEL_BLAD}' is leeg.`);
  const nRijen = used.getRowIndex() + used.getRowCount();
  const nKols = used.getColumnIndex() + used.getColumnCount();
  const doelData: Cel[][] = doel.getRangeByIndexes(0, 0, nRijen, nKols).getValues() as Cel[][];

  // --- kolommen op het doelblad (rij 1) ---
  const kop = doelData[0];
  const dMaster = kolom(kop, ["Master"]);
  const dBp = kolom(kop, ["BP File"]);
  const dRef = kolom(kop, ["Reference"]);
  const dEcIc = kolom(kop, ["IC/EC", "EC/IC"]);
  const dDest = kolom(kop, ["Dest"]);
  const dLaad = kolom(kop, ["Loading date"]);
  if (dMaster < 0) return meld(`Kolom 'Master' niet gevonden op rij 1 van '${DOEL_BLAD}'.`);
  if (dBp < 0 && dRef < 0 && dEcIc < 0 && dDest < 0 && dLaad < 0) {
    return meld(`Geen van de kolommen BP File, Reference, IC/EC, Dest of Loading date gevonden op rij 1 van '${DOEL_BLAD}'.`);
  }

  // --- planning lezen ---
  let bron: Cel[][] | undefined;
  let bronNaam = `tabel ${BRON_TABEL}`;
  const t = workbook.getTable(BRON_TABEL);
  if (t) {
    bron = t.getRange().getValues() as Cel[][];
  } else {
    bron = leesBlad(workbook);
    bronNaam = `blad ${BRON_BLAD}`;
  }
  if (!bron) {
    return meld(`Geen planning gevonden: tabel '${BRON_TABEL}' en blad '${BRON_BLAD}' staan niet in deze werkmap. Er is niets aangepast.`);
  }

  const bKop = bron[0];
  const cMaster = kolom(bKop, ["MASTER"]);
  const cBp = kolom(bKop, ["BP FILE"]);
  const cShip = kolom(bKop, ["SHIPMENT"]);
  const cEcIc = kolom(bKop, ["EC/IC", "IC/EC"]);
  const cDest = kolom(bKop, ["DEST"]);
  const cLaad = kolom(bKop, ["LAAD DATUM", "LD DATE", "LD. DATE"]);
  const cLoaded = kolom(bKop, ["LOADED"]);
  if (cMaster < 0) return meld(`Kolom MASTER niet gevonden in de planning (${bronNaam}).`);

  // --- index master -> beste bronregel ---
  const idx = new Map<string, number>();
  const dubbel = new Set<string>();
  for (let r = 1; r < bron.length; r++) {
    const s = sleutel(bron[r][cMaster]);
    if (!s) continue;
    const oud = idx.get(s);
    if (oud === undefined) idx.set(s, r);
    else {
      dubbel.add(s);
      if (beter(bron, r, oud, cLoaded, cLaad)) idx.set(s, r);
    }
  }

  // --- regels bijwerken (alleen de groene kolommen) ---
  let nRegels = 0, nRegelsGewijzigd = 0, nCellen = 0;
  const nietGevonden: string[] = [];
  const dubbelOpBlad: string[] = [];
  for (let r = 1; r < doelData.length; r++) {
    const s = sleutel(doelData[r][dMaster]);
    if (!s) continue;
    nRegels++;
    if (dubbel.has(s) && dubbelOpBlad.indexOf(s) < 0) dubbelOpBlad.push(s);
    const br = idx.get(s);
    if (br === undefined) {
      if (nietGevonden.indexOf(s) < 0) nietGevonden.push(s);
      continue;
    }
    const rij = bron[br];
    let nu = 0;
    // alleen bijwerken als de kolom in de planning bestaat (anders niets wissen)
    if (cBp >= 0) nu += bijwerken(doel, doelData, r, dBp, getal(rij[cBp]), false);
    if (cShip >= 0) nu += bijwerken(doel, doelData, r, dRef, getal(rij[cShip]), false);
    if (cEcIc >= 0) nu += bijwerken(doel, doelData, r, dEcIc, tekst(rij[cEcIc]), false);
    if (cDest >= 0) nu += bijwerken(doel, doelData, r, dDest, tekst(rij[cDest]), false);
    if (cLaad >= 0) nu += bijwerken(doel, doelData, r, dLaad, datum(rij[cLaad]), true);
    nCellen += nu;
    if (nu > 0) nRegelsGewijzigd++;
  }

  // --- melding ---
  let m = `${nRegels} regel(s) met een masternummer gecontroleerd op '${doel.getName()}'.\n` +
    `${nRegelsGewijzigd} regel(s) bijgewerkt, ${nCellen} cel(len) gewijzigd in BP File / Reference / IC/EC / Dest / Loading date.\n` +
    `Bron: ${bronNaam}.`;
  if (nietGevonden.length) m += `\n\n${nietGevonden.length} master(s) niet gevonden in de planning (ongewijzigd gelaten):\n${nietGevonden.join(", ")}`;
  if (dubbelOpBlad.length) m += `\n\nLet op: deze masters staan meerdere keren in de planning (open regel / laatste laaddatum gebruikt):\n${dubbelOpBlad.join(", ")}`;
  if (cLaad < 0) m += `\n\nKolom LAAD DATUM / LD DATE niet gevonden in de planning: Loading date is niet bijgewerkt.`;
  return meld(m);
}

/** Zet de waarde uit de planning in een groene cel als die afwijkt; 1 = gewijzigd. */
function bijwerken(ws: ExcelScript.Worksheet, data: Cel[][], r: number, c: number, nieuw: Cel | null, isDatum: boolean): number {
  if (c < 0) return 0;
  const huidig: Cel = c < data[r].length ? data[r][c] : "";
  if (gelijk(huidig, nieuw, isDatum)) return 0;
  const cel = ws.getCell(r, c);
  if (nieuw === null) {
    cel.clear(ExcelScript.ClearApplyTo.contents); // waarde weg, groene opmaak blijft
  } else {
    cel.setValue(nieuw);
    if (isDatum) cel.setNumberFormat("dd-mm-yyyy");
  }
  return 1;
}

function gelijk(huidig: Cel, nieuw: Cel | null, isDatum: boolean): boolean {
  const leeg = String(huidig ?? "").trim() === "";
  if (nieuw === null) return leeg;
  if (leeg) return false;
  if (isDatum) return typeof huidig === "number" && Math.floor(huidig) === Math.floor(nieuw as number);
  return String(huidig).trim().toUpperCase() === String(nieuw).trim().toUpperCase();
}

/** Masternummer -> alleen cijfers (vangt tekst/getal, spaties en regeleindes af). */
function sleutel(v: Cel): string {
  if (v === null || v === undefined) return "";
  const s = typeof v === "number" ? Math.round(v).toString() : String(v);
  return s.replace(/\D/g, "");
}

/** BP File / Reference: cijfers als getal, anders de tekst; leeg = null. */
function getal(v: Cel): Cel | null {
  if (v === null || v === undefined) return null;
  if (typeof v === "number") return v;
  const s = String(v).replace(/[\r\n]/g, "").trim();
  if (!s) return null;
  return /^\d+$/.test(s) ? Number(s) : s;
}

function tekst(v: Cel): Cel | null {
  const s = String(v ?? "").replace(/[\t\r\n]/g, "").trim().toUpperCase();
  return s ? s : null;
}

/** Laaddatum als Excel-datumgetal; placeholders (voor 2000) en tekst zonder datum = null. */
function datum(v: Cel): number | null {
  if (typeof v === "number") return v >= DATUM_GRENS ? Math.floor(v) : null;
  const s = String(v ?? "").trim();
  const m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/.exec(s) || null;
  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  let d: Date | null = null;
  if (m) d = new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
  else if (iso) d = new Date(Date.UTC(+iso[1], +iso[2] - 1, +iso[3]));
  if (!d || isNaN(d.getTime())) return null;
  const serieel = Math.round(d.getTime() / 86400000) + 25569;
  return serieel >= DATUM_GRENS ? serieel : null;
}

/** Bij dubbele masters: open regel (LOADED <> Y) wint, daarna de laatste laaddatum. */
function beter(bron: Cel[][], nieuw: number, oud: number, cLoaded: number, cLaad: number): boolean {
  if (cLoaded >= 0) {
    const openN = String(bron[nieuw][cLoaded] ?? "").trim().toUpperCase() !== "Y";
    const openO = String(bron[oud][cLoaded] ?? "").trim().toUpperCase() !== "Y";
    if (openN !== openO) return openN;
  }
  const dN = cLaad >= 0 ? (datum(bron[nieuw][cLaad]) ?? 0) : 0;
  const dO = cLaad >= 0 ? (datum(bron[oud][cLaad]) ?? 0) : 0;
  return dN > dO;
}

/** Kolomindex (0-gebaseerd) op naam; hoofdletter- en spatieongevoelig. -1 = niet gevonden. */
function kolom(kop: Cel[], namen: string[]): number {
  const norm = (v: Cel) => String(v ?? "").replace(/[\s\u00a0]/g, "").toUpperCase();
  const gezocht = namen.map(n => norm(n));
  for (let c = 0; c < kop.length; c++) if (gezocht.indexOf(norm(kop[c])) >= 0) return c;
  return -1;
}

/** Blad Uit zonder tabel: zoek de koprij (MASTER + BP FILE) in de eerste 10 rijen. */
function leesBlad(workbook: ExcelScript.Workbook): Cel[][] | undefined {
  const ws = workbook.getWorksheets().find(w => w.getName().trim().toUpperCase() === BRON_BLAD.toUpperCase());
  if (!ws) return undefined;
  const used = ws.getUsedRange(true);
  if (!used) return undefined;
  const n = used.getRowIndex() + used.getRowCount();
  const k = used.getColumnIndex() + used.getColumnCount();
  const data = ws.getRangeByIndexes(0, 0, n, k).getValues() as Cel[][];
  for (let r = 0; r < Math.min(10, data.length); r++) {
    if (kolom(data[r], ["MASTER"]) >= 0 && kolom(data[r], ["BP FILE"]) >= 0) return data.slice(r);
  }
  return undefined;
}

function meld(tekstMelding: string): void {
  console.log(tekstMelding);
}
