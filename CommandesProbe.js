/***************************************************************
 * CommandesProbe.js — READ ONLY probe of the Commandes spreadsheet.
 * Temporary: used to design Screen 3. Delete once the layout is known.
 ***************************************************************/

const CMD_PROBE = {
  SS_ID: "1QYxnnfMoYidqN8l0ZQZZe5MHER5EBeCEXXvTzyhKan8",
  SHEET: "2026"
};

/** Dump tab list, header row, and the last few data rows. */
function probeCommandes() {
  const ss = SpreadsheetApp.openById(CMD_PROBE.SS_ID);
  Logger.log("TABS: " + ss.getSheets().map(s => s.getName()).join(", "));

  const sh = ss.getSheetByName(CMD_PROBE.SHEET);
  if (!sh) { Logger.log("Sheet '" + CMD_PROBE.SHEET + "' not found"); return; }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  Logger.log("Sheet '" + CMD_PROBE.SHEET + "': lastRow=" + lastRow + " lastCol=" + lastCol);

  // Header row(s)
  for (let r = 1; r <= 2; r++) {
    const vals = sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    const named = vals.map((v, i) =>
      v ? colLetter(i + 1) + "=" + v : null).filter(Boolean);
    Logger.log("ROW " + r + ": " + JSON.stringify(named));
  }

  // Last 3 data rows, to see what real entries look like
  const start = Math.max(2, lastRow - 2);
  for (let r = start; r <= lastRow; r++) {
    const vals = sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    const named = vals.map((v, i) =>
      v ? colLetter(i + 1) + "=" + v : null).filter(Boolean);
    Logger.log("DATA r" + r + ": " + JSON.stringify(named));
  }
}

/** Which columns carry dropdowns, and what are the options? */
function probeCommandesDropdowns() {
  const ss = SpreadsheetApp.openById(CMD_PROBE.SS_ID);
  const sh = ss.getSheetByName(CMD_PROBE.SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const probeRow = Math.max(2, lastRow - 1);

  for (let c = 1; c <= lastCol; c++) {
    const rule = sh.getRange(probeRow, c).getDataValidation();
    if (!rule) continue;
    let opts = "";
    try {
      const cv = rule.getCriteriaValues();
      opts = Array.isArray(cv[0]) ? cv[0].slice(0, 15).join(" | ") : String(cv[0]);
    } catch (e) { opts = "(non listable)"; }
    Logger.log(colLetter(c) + " (" + rule.getCriteriaType() + "): " + opts);
  }
}

function colLetter(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
