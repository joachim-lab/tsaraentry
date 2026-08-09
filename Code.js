/***************************************************************
 * Code.js — TSARA Entry web app
 * Screen 1: Nourrissage (feed quantity entry)
 *
 * Writes rows to the Nourrissage file's "Consommation provende"
 * sheet (columns C:F + formula in I), then calls the
 * controleconsoprovende library (bound at HEAD) to fill column H —
 * the exact same fill logic the manual-edit trigger uses.
 ***************************************************************/

const CFG = {
  STOCK_POISSON_SS_ID: "1Kfs5beQorhdheqzEDibgnBd5wQjy79MecncNlgKjRIE",
  STOCK_SHEET: "lot",
  STOCK_LOT_COL: 14,       // N = lot id
  STOCK_LOT_START_ROW: 3,
  STOCK_LOT_END_ROW: 50,

  NOURRISSAGE_SS_ID: "1JBoH5c7BqZc2V5czcDAnEt-2hvkNKlJAxuDSupttTfs",
  CONSO_SHEET: "Consommation provende",
  CONSO_START_ROW: 2,
  CONSO_KEY_COL: 3,        // C = lot key (write block C:F starts here)
  CONSO_TYPE_COL: 5,       // E = type provende (for reading the dropdown list)
  CONSO_PCT_COL: 9         // I = % différence (formula =IFERROR(F/H,""))
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('TSARA Entry')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Lot list for the dropdown, from Stock Poisson N3:N50 — same source actionsurstock uses. */
function getLotList() {
  const ss = SpreadsheetApp.openById(CFG.STOCK_POISSON_SS_ID);
  const sh = ss.getSheetByName(CFG.STOCK_SHEET);
  if (!sh) throw new Error('Stock Poisson sheet not found: "' + CFG.STOCK_SHEET + '"');
  const n = CFG.STOCK_LOT_END_ROW - CFG.STOCK_LOT_START_ROW + 1;
  const vals = sh.getRange(CFG.STOCK_LOT_START_ROW, CFG.STOCK_LOT_COL, n, 1).getValues();
  return vals.map(r => r[0]).filter(v => v !== "" && v !== null);
}

/** Feed type list, read live from the existing dropdown rule on column E — never hardcoded, so it can't drift. */
function getFeedTypes() {
  const ss = SpreadsheetApp.openById(CFG.NOURRISSAGE_SS_ID);
  const sh = ss.getSheetByName(CFG.CONSO_SHEET);
  if (!sh) throw new Error('Sheet not found: "' + CFG.CONSO_SHEET + '"');
  const rule = sh.getRange(2, CFG.CONSO_TYPE_COL).getDataValidation();
  if (!rule) return [];
  const criteria = rule.getCriteriaValues(); // [ [list values], ... ] for requireValueInList
  return criteria[0] || [];
}

/**
 * Append one or more feed entries:
 *  1) values into C:F,
 *  2) the % différence formula into I (same formula as existing rows),
 *  3) fill column H via the controleconsoprovende library — same code
 *     path as manual edits.
 *
 * entries: [{ lot, date, type, qty }], date is "yyyy-MM-dd" from the browser.
 * Returns { written, startRow, endRow }.
 * If the library call fails, the rows are NOT rolled back — losing a
 * feed record is worse than a blank H — the error is surfaced instead.
 */
function submitNourrissage(entries) {
  if (!entries || !entries.length) throw new Error("No entries to submit.");

  const ss = SpreadsheetApp.openById(CFG.NOURRISSAGE_SS_ID);
  const sh = ss.getSheetByName(CFG.CONSO_SHEET);
  if (!sh) throw new Error('Sheet not found: "' + CFG.CONSO_SHEET + '"');

  const rows = entries.map(en => {
    if (!en.lot || !en.date || !en.type || !en.qty) {
      throw new Error("Incomplete entry: lot, date, type and qty are all required.");
    }
    const qty = Number(en.qty);
    if (!isFinite(qty) || qty <= 0) {
      throw new Error("Qty must be a positive number (got: " + en.qty + ").");
    }
    return [en.lot, new Date(en.date), en.type, qty];
  });

  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, CFG.CONSO_KEY_COL, rows.length, 4).setValues(rows); // C:F

  // Column I: same formula as existing rows, in R1C1 so it adapts per row.
  // I = F / H  ->  RC[-3] / RC[-1]
  sh.getRange(startRow, CFG.CONSO_PCT_COL, rows.length, 1)
    .setFormulaR1C1('=IFERROR(RC[-3]/RC[-1],"")');

  const endRow = startRow + rows.length - 1;

  try {
    ConsoProvende.fillHForRows(startRow, endRow);
  } catch (err) {
    throw new Error(
      "Rows " + startRow + "-" + endRow + " were saved, but the theoretical-quantity " +
      "fill (column H) failed: " + err + ". Please tell Kim."
    );
  }

  return { written: rows.length, startRow: startRow, endRow: endRow };
}

/**
 * RUN FROM EDITOR once after first push: verifies the library binding
 * resolves and is callable. Calls fillHForRows with an empty range
 * (endRow < startRow) so it processes 0 rows and writes nothing.
 */
function testLibraryBinding() {
  const result = ConsoProvende.fillHForRows(2, 1);
  Logger.log("Library OK, fillHForRows returned: " + result);
  return "Library OK (processed " + result + " rows, as expected 0)";
}
