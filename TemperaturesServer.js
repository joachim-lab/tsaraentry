/***************************************************************
 * TemperaturesServer.js — TSARA Entry web app, screen 7
 *
 * Three jobs, one screen:
 *   1. Write the two water temperatures into Stock poisson,
 *      sheet "lot", cells R1 (bassin) and T1 (lac).
 *   2. Build a PDF of that sheet, columns N to X, A4 landscape,
 *      and hand it to the browser for printing.
 *   3. Build a PDF of the Programme "planning" sheet, print area
 *      sized to its actual content (not a fixed range — the sheet
 *      is short-lived and its row count changes week to week).
 *
 * R1 and T1 are NOT a new convention. WARNINGSYSTEM reads them
 * (WS_CFG.TEMP) and the cockpit reads them (KPI_CFG.STOCK).
 * This screen writes the same two cells, so there is one source
 * of truth. Do not move them without changing those two projects.
 *
 * The engine writes the "lot" sheet from row 3 down, so row 1 is
 * free. Both spreadsheet IDs come from CFG in Code.js — no copy.
 *
 * No date stamp is written. The cockpit already logs both values
 * with the date in its hidden "_TempLog" sheet.
 ***************************************************************/

const TEMP_CFG = {
  BASSIN_CELL: "R1",
  LAC_CELL: "T1",
  MIN_C: 0,               // refuse a typo, not a real reading
  MAX_C: 45,
  SAVED_PROP: "TEMP_SAVED_DATE",   // Script Property: date of the last tempSave
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

/** Today as yyyy-MM-dd, in the script timezone. One definition. */
function tempToday() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/**
 * The date of the last tempSave, as yyyy-MM-dd, or "" when nobody has
 * saved since this gate was installed.
 *
 * The date lives in a Script Property, NOT in the sheet. R1 and T1 are
 * read by WARNINGSYSTEM (WS_CFG.TEMP) and by tsaracockpit (KPI_CFG.TEMP).
 * Those two cells must hold a temperature and nothing else.
 */
function tempSavedDate() {
  const v = PropertiesService.getScriptProperties().getProperty(TEMP_CFG.SAVED_PROP);
  return v ? String(v) : "";
}

/**
 * Current values, as shown in the sheet, plus the date they were saved.
 * The screen shows them as text only. It does NOT preselect them: the
 * worker must read the thermometer and choose both values again.
 */
function tempGetCurrent() {
  const sh = tempSheet();
  const saved = tempSavedDate();
  return {
    bassin: sh.getRange(TEMP_CFG.BASSIN_CELL).getDisplayValue(),
    lac: sh.getRange(TEMP_CFG.LAC_CELL).getDisplayValue(),
    savedDate: saved,
    savedToday: saved === tempToday()
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
  const s = String(raw === null || raw === undefined ? "" : raw).trim();
  if (s === "") {
    throw new Error("Température " + label + " : choisissez une valeur.");
  }
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

  // The save date is what the print gate reads. Write it only after both
  // cells are written, so a failed write never opens the gate.
  const today = tempToday();
  PropertiesService.getScriptProperties().setProperty(TEMP_CFG.SAVED_PROP, today);

  return { bassin: bassin, lac: lac, savedDate: today };
}

/**
 * Fetches a Sheets PDF export URL and returns it base64-encoded.
 * Shared by both PDF buttons on this screen. Throws a French message
 * on a non-200 response instead of returning a broken PDF.
 */
function tempFetchPdf(url, filenamePrefix) {
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
    filename: filenamePrefix + "_" + stamp + ".pdf",
    size: bytes.length
  };
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

  return tempFetchPdf(url, "Stock_poisson");
}

/**
 * The Stock poisson print, as called by the Impressions screen.
 *
 * The temperatures must be saved on the same calendar day. That is the
 * whole purpose of the gate: the printed sheet carries a water
 * temperature, so that temperature must be today's reading.
 *
 * The build stays in tempBuildPdf, with no gate, so that
 * testTemperaturesServer proves the export on any day.
 */
function tempPrintStockPdf() {
  if (tempSavedDate() !== tempToday()) {
    throw new Error(
      "Températures pas à jour. Enregistrez la température bassin et la " +
      "température lac aujourd'hui, avant d'imprimer le stock poisson."
    );
  }
  return tempBuildPdf();
}

/** The Programme "planning" sheet. Throws if it is missing. */
function tempProgrammeSheet() {
  const sh = SpreadsheetApp
    .openById(CFG.PROGRAMME_SS_ID)
    .getSheetByName(CFG.PROGRAMME_SHEET);
  if (!sh) {
    throw new Error('Onglet "' + CFG.PROGRAMME_SHEET + '" introuvable dans Programme.');
  }
  return sh;
}

/**
 * PDF of the Programme "planning" sheet. Print area is sized to the
 * sheet's actual content (getLastRow/getLastColumn), not a fixed
 * range — the sheet is short and its size changes week to week.
 * A4 landscape, same narrow margins as the Stock poisson PDF.
 */
function tempBuildProgrammePdf() {
  const sh = tempProgrammeSheet();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  const url = "https://docs.google.com/spreadsheets/d/" + CFG.PROGRAMME_SS_ID + "/export"
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
    + "&c1=0"
    + "&r2=" + lastRow
    + "&c2=" + lastCol;

  return tempFetchPdf(url, "Programme");
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

  Logger.log("Dernière saisie température : " + (tempSavedDate() || "(aucune)") +
             " · aujourd'hui : " + tempToday() +
             " · impression autorisée : " + (tempSavedDate() === tempToday()));

  const pdf = tempBuildPdf();
  Logger.log("PDF Stock poisson construit : " + pdf.size + " octets · " + pdf.filename);

  const progSh = tempProgrammeSheet();
  Logger.log("Programme onglet : " + progSh.getName() + " · dernière ligne " + progSh.getLastRow() + " · dernière colonne " + progSh.getLastColumn());
  const progPdf = tempBuildProgrammePdf();
  Logger.log("PDF Programme construit : " + progPdf.size + " octets · " + progPdf.filename);
}
