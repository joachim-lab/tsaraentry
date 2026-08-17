/***************************************************************
 * Code.js — TSARA Entry web app
 * Screen 1: Nourrissage (feed quantity entry)
 *
 * Writes rows to the Nourrissage file's "Consommation provende"
 * sheet (columns C:F), copies the % formula in I from the row
 * above (PASTE_FORMULA — the sheet's own formula stays the single
 * source of truth), then calls the controleconsoprovende library
 * (bound at HEAD) to fill column H — the exact same fill logic
 * the manual-edit trigger uses.
 ***************************************************************/

const CFG = {
  STOCK_POISSON_SS_ID: "1Kfs5beQorhdheqzEDibgnBd5wQjy79MecncNlgKjRIE",
  STOCK_SHEET: "lot",
  STOCK_LOT_COL: 14,       // N = lot id
  STOCK_LOT_START_ROW: 3,
  STOCK_LOT_END_ROW: 50,

  PROGRAMME_SS_ID: "1Ky1DKbNxeBnqTXUxgk8PedPSuwISgS4rvVhFMUHBdsM",
  PROGRAMME_SHEET: "planning",

  NOURRISSAGE_SS_ID: "1JBoH5c7BqZc2V5czcDAnEt-2hvkNKlJAxuDSupttTfs",
  CONSO_SHEET: "Consommation provende",
  CONSO_START_ROW: 2,
  CONSO_KEY_COL: 3,        // C = lot key (write block C:F starts here)
  CONSO_TYPE_COL: 5,       // E = type provende (for reading the dropdown list)
  CONSO_F_COL: 6,          // F = qty given
  CONSO_PCT_COL: 9         // I = % différence (formula copied from row above)
};

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
 * First row after the last real entry, scanning columns C (lot key)
 * and F (qty). getLastRow() is unusable here: column G holds
 * pre-inserted checkboxes far below the data, so it reports the
 * bottom of the sheet, not the last entry.
 */
function findNextConsoRow(sh) {
  const lastPhysical = sh.getLastRow();
  const n = lastPhysical - CFG.CONSO_START_ROW + 1;
  if (n < 1) return CFG.CONSO_START_ROW;

  const cVals = sh.getRange(CFG.CONSO_START_ROW, CFG.CONSO_KEY_COL, n, 1).getValues();
  const fVals = sh.getRange(CFG.CONSO_START_ROW, CFG.CONSO_F_COL, n, 1).getValues();

  let lastData = CFG.CONSO_START_ROW - 1;
  for (let i = 0; i < n; i++) {
    const c = cVals[i][0];
    const f = fVals[i][0];
    if ((c !== "" && c !== null) || (f !== "" && f !== null)) {
      lastData = CFG.CONSO_START_ROW + i;
    }
  }
  return lastData + 1;
}

/**
 * Append one or more feed entries:
 *  1) values into C:F at the first row after the last real entry,
 *  2) column I formula copied from the row above (PASTE_FORMULA),
 *  3) fill column H via the controleconsoprovende library — same
 *     code path as manual edits.
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

  const startRow = findNextConsoRow(sh);
  sh.getRange(startRow, CFG.CONSO_KEY_COL, rows.length, 4).setValues(rows); // C:F
  const endRow = startRow + rows.length - 1;

  // Column I: copy the formula from the row above so the sheet's own
  // formula (locale, exact form) remains the single source of truth.
  if (startRow > CFG.CONSO_START_ROW) {
    const src = sh.getRange(startRow - 1, CFG.CONSO_PCT_COL);
    const dst = sh.getRange(startRow, CFG.CONSO_PCT_COL, rows.length, 1);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  }

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
