/***************************************************************
 * RecetteMtServer.js — TSARA Entry web app
 * Screen: Recette provende MT (impression)
 *
 * Prints the MT feed recipe of one lot: the stock solution and the
 * four weekly batches of the 21-day masculinisation phase.
 *
 * NOTHING IS CALCULATED HERE. Every figure comes from the lot file's
 * own "Impression MT" tab, which is only a set of plain references to
 * the period tabs 1-5, 6-10, 11-15 and 16-21. Those in turn read the
 * "stock & working solution" tabs. The arithmetic has never moved.
 *
 * The tab is built by buildImpressionMT() in the Apps Script project
 * "creationnouveaufichierlot". It reaches a lot file because every lot
 * file is a copy of the template. A lot created BEFORE that change has
 * no such tab and this screen says so plainly.
 *
 * DEPENDENCIES (deliberate, so there is one of each in the project):
 *   getLotFileList()  Lot.js               — the lot list
 *   tempFetchPdf()    TemperaturesServer.js — the PDF fetch + encode
 ***************************************************************/

const RECETTE_MT_CFG = {
  SHEET: "Impression MT",   // built by buildImpressionMT()
  SOURCE_SHEET: "1-5",      // holds the 21-day total
  TOTAL_CELL: "C24",        // provende MT totale — 0 means no fish count yet
  PRINT_LAST_ROW: 20,       // A1:F20 is the whole printable block
  PRINT_LAST_COL: 6
};

/**
 * Lot files that can have an MT recipe: growout lots only.
 *
 * A broodstock lot id is letters only (Mirana). Those lots never go
 * through masculinisation, so they are not offered.
 *
 * @return {Array<Object>} [{ lotNumber, fileId, fileName }]
 */
function recetteListLots() {
  return getLotFileList().filter(function (l) {
    return /^[0-9]/.test(String(l.lotNumber));
  });
}

/**
 * PDF of one lot's "Impression MT" tab, A1:F20, A4 portrait.
 *
 * Refuses before building anything when the lot has no fish count:
 * every quantity on the tab is derived from it, so the page would be
 * a sheet of zeros — worse than no page at all.
 *
 * @param {string} fileId  Lot file id, from recetteListLots().
 * @return {Object} { base64, filename, size }
 */
function recetteBuildPdf(fileId) {
  if (!fileId) throw new Error("Aucun lot choisi.");

  const ss = SpreadsheetApp.openById(fileId);

  const src = ss.getSheetByName(RECETTE_MT_CFG.SOURCE_SHEET);
  if (!src) {
    throw new Error("Onglet « " + RECETTE_MT_CFG.SOURCE_SHEET +
      " » introuvable dans " + ss.getName() + ".");
  }

  const sh = ss.getSheetByName(RECETTE_MT_CFG.SHEET);
  if (!sh) {
    throw new Error("Ce lot n'a pas l'onglet « " + RECETTE_MT_CFG.SHEET +
      " ». Il a été créé avant la mise en place de la recette MT.");
  }

  const total = Number(src.getRange(RECETTE_MT_CFG.TOTAL_CELL).getValue()) || 0;
  if (total <= 0) {
    throw new Error("Effectif du lot pas encore saisi : impossible de calculer la recette. " +
      "Saisissez le nombre de poissons, puis imprimez.");
  }

  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export"
    + "?format=pdf"
    + "&gid=" + sh.getSheetId()
    + "&size=A4"
    + "&portrait=true"
    + "&fitw=true"
    + "&gridlines=false"
    + "&printtitle=false"
    + "&sheetnames=false"
    + "&pagenum=UNDEFINED"
    + "&attachment=false"
    + "&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5"
    + "&r1=0"
    + "&c1=0"
    + "&r2=" + RECETTE_MT_CFG.PRINT_LAST_ROW
    + "&c2=" + RECETTE_MT_CFG.PRINT_LAST_COL;

  return tempFetchPdf(url, "RecetteMT_" + ss.getName());
}

/**
 * Read-only check. Run from the editor after the push and BEFORE the
 * new deployment: it writes nothing, and it triggers any new
 * permission prompt while /exec still serves the old version.
 *
 * Script: tsaraentry · file: RecetteMtServer.js
 */
function testRecetteMtServer() {
  const lots = recetteListLots();
  Logger.log("Lots proposés : " + lots.length);
  lots.slice(-5).forEach(function (l) {
    Logger.log("  " + l.fileName + " · " + l.fileId);
  });
  if (!lots.length) return;

  const last = lots[lots.length - 1];
  Logger.log("Test sur " + last.fileName);
  try {
    const pdf = recetteBuildPdf(last.fileId);
    Logger.log("PDF construit : " + pdf.size + " octets · " + pdf.filename);
  } catch (e) {
    Logger.log("Refus (attendu si effectif vide ou onglet absent) : " + e.message);
  }
}
