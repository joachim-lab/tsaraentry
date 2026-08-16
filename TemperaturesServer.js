/***************************************************************
 * TemperaturesServer.js — TSARA Entry web app, screen 7
 *
 * Two jobs, one screen:
 *   1. Write the two water temperatures into Stock poisson,
 *      sheet "lot", cells R1 (bassin) and T1 (lac).
 *   2. Build a PDF of that sheet, columns N to X, A4 landscape,
 *      and hand it to the browser for printing.
 *
 * R1 and T1 are NOT a new convention. WARNINGSYSTEM reads them
 * (WS_CFG.TEMP) and the cockpit reads them (KPI_CFG.STOCK).
 * This screen writes the same two cells, so there is one source
 * of truth. Do not move them without changing those two projects.
 *
 * The engine writes the "lot" sheet from row 3 down, so row 1 is
 * free. The spreadsheet ID comes from CFG in Code.js — no copy.
 *
 * No date stamp is written. The cockpit already logs both values
 * with the date in its hidden "_TempLog" sheet.
 ***************************************************************/

const TEMP_CFG = {
  BASSIN_CELL: "R1",
  LAC_CELL: "T1",
  MIN_C: 0,               // refuse a typo, not a real reading
  MAX_C: 45,
  PRINT_FIRST_COL: 14,    // N
  PRINT_LAST_COL: 24      // X
};

/** The Stock poisson "lot" sheet. Throws if it is missing. */
function tempSheet() {
  const sh = SpreadsheetApp
    .openById(CFG.STOCK_POISSON_SS_ID)
    .getSheetByName(CFG.STOCK_SHEET);
  if (!sh) {
    throw new Error('Onglet "' + CFG.STOCK_SHEET + '" introuvable dans Stock poisson.');
  }
  return sh;
}

/** Current values, as shown in the sheet. Used to fill the form. */
function tempGetCurrent() {
  const sh = tempSheet();
  return {
    bassin: sh.getRange(TEMP_CFG.BASSIN_CELL).getDisplayValue(),
    lac: sh.getRange(TEMP_CFG.LAC_CELL).getDisplayValue()
  };
}

/** "26,5" or "26.5" -> 26.5. Anything else -> null. */
function tempParseNumber(v) {
  const s = String(v === null || v === undefined ? "" : v).trim().replace(",", ".");
  if (s === "") return null;
  const n = Number(s);
  return isFinite(n) ? n : null;
}

/** Checks one reading and returns it, or throws a French message. */
function tempCheckReading(label, raw) {
  const n = tempParseNumber(raw);
  if (n === null) {
    throw new Error("Température " + label + " : valeur non numérique.");
  }
  if (n < TEMP_CFG.MIN_C || n > TEMP_CFG.MAX_C) {
    throw new Error(
      "Température " + label + " : " + n + " °C est hors plage (" +
      TEMP_CFG.MIN_C + " à " + TEMP_CFG.MAX_C + " °C). Vérifiez la saisie."
    );
  }
  return n;
}

/** Writes both cells. Both are checked before either is written. */
function tempSave(bassinRaw, lacRaw) {
  const bassin = tempCheckReading("bassin", bassinRaw);
  const lac = tempCheckReading("lac", lacRaw);

  const sh = tempSheet();
  sh.getRange(TEMP_CFG.BASSIN_CELL).setValue(bassin);
  sh.getRange(TEMP_CFG.LAC_CELL).setValue(lac);
  SpreadsheetApp.flush();

  return { bassin: bassin, lac: lac };
}

/**
 * PDF of the "lot" sheet, columns N to X, all rows, A4 landscape,
 * narrow margins. Returned base64-encoded, because the browser
 * cannot read the sheet itself — the web app runs as the deploying
 * user and the workers have no access to the file.
 *
 * r1/c1/r2/c2 are 0-based; c1 is inclusive and c2 is exclusive.
 */
function tempBuildPdf() {
  const sh = tempSheet();
  const lastRow = sh.getLastRow();

  const url = "https://docs.google.com/spreadsheets/d/" + CFG.STOCK_POISSON_SS_ID + "/export"
    + "?format=pdf"
    + "&gid=" + sh.getSheetId()
    + "&size=A4"
    + "&portrait=false"
    + "&fitw=true"
    + "&gridlines=true"
    + "&printtitle=false"
    + "&sheetnames=false"
    + "&pagenum=UNDEFINED"
    + "&attachment=false"
    + "&top_margin=0.25&bottom_margin=0.25&left_margin=0.25&right_margin=0.25"
    + "&r1=0"
    + "&c1=" + (TEMP_CFG.PRINT_FIRST_COL - 1)
    + "&r2=" + lastRow
    + "&c2=" + TEMP_CFG.PRINT_LAST_COL;

  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error("Google a refusé l'export PDF (code " + res.getResponseCode() + ").");
  }

  const bytes = res.getBlob().getBytes();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");

  return {
    base64: Utilities.base64Encode(bytes),
    filename: "Stock_poisson_" + stamp + ".pdf",
    size: bytes.length
  };
}

/**
 * Read-only check. Run this from the editor after the push and
 * BEFORE the new deployment. It writes nothing. It proves three
 * things: the sheet resolves, R1 and T1 are not merged cells, and
 * the PDF builds — which also triggers the new UrlFetchApp
 * permission prompt while /exec still serves the old version.
 */
function testTemperaturesServer() {
  const sh = tempSheet();
  Logger.log("Onglet : " + sh.getName() + " · gid " + sh.getSheetId() + " · dernière ligne " + sh.getLastRow());
  Logger.log("R1 fusionnée : " + sh.getRange(TEMP_CFG.BASSIN_CELL).isPartOfMerge());
  Logger.log("T1 fusionnée : " + sh.getRange(TEMP_CFG.LAC_CELL).isPartOfMerge());

  const cur = tempGetCurrent();
  Logger.log("Valeurs lues — bassin : " + cur.bassin + " · lac : " + cur.lac);

  const pdf = tempBuildPdf();
  Logger.log("PDF construit : " + pdf.size + " octets · " + pdf.filename);
}
