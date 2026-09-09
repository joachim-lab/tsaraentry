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

  DATA_BASSINS_SS_ID: "1n4Xj4hZnGthAsFig3ecK9TCzegYdyh-0y9ISWfU36D0",

  NOURRISSAGE_SS_ID: "1JBoH5c7BqZc2V5czcDAnEt-2hvkNKlJAxuDSupttTfs",
  CONSO_SHEET: "Consommation provende",
  CONSO_START_ROW: 2,
  CONSO_KEY_COL: 3,        // C = lot key (write block C:F starts here)
  CONSO_TYPE_COL: 5,       // E = type provende (for reading the dropdown list)
  CONSO_F_COL: 6,          // F = qty given
  CONSO_PCT_COL: 9,        // I = % différence (formula copied from row above)

  /* ---- Correction of an entry already saved ---- */
  CONSO_DATE_COL: 4,       // D = date of the entry
  CORR_WINDOW_DAYS: 1,     // 0 = today only, 1 = today and yesterday
  CORR_SCAN_ROWS: 400,     // how far back to read when listing recent entries
  CORR_LOG_SHEET: "Corrections"
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
  try {
    sh.getRange(startRow, CFG.CONSO_KEY_COL, rows.length, 4).setValues(rows); // C:F
    SpreadsheetApp.flush(); // apply the write NOW: a rejected write must fail here, not later
  } catch (err) {
    throw new Error(
      "NO rows were saved - the sheet rejected the write: " + err + ". Please tell Kim."
    );
  }
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


/***************************************************************
 * CORRECTION OF A SAVED ENTRY
 *
 * A worker who typed the wrong lot, feed type or quantity fixes
 * the row in place. There is one row per feeding event before and
 * after a correction, so every consumer of "Consommation provende"
 * (projections, FCR, feed guards, monthly report) keeps working
 * with no change.
 *
 * FOUR RULES, and the reason for each:
 *
 * 1. The date is NOT editable. The rows are in date order because
 *    they are appended; changing a date would put a row out of
 *    order without moving it.
 *
 * 2. Only today and yesterday (CORR_WINDOW_DAYS). Anything older
 *    has already been read by the monthly report, the forecast
 *    snapshot and the stock backup.
 *
 * 3. The row is identified by its number AND by a signature of its
 *    own C:F values. If another worker appended a row in between,
 *    or edited this one, the signature no longer matches and the
 *    correction is refused instead of overwriting a stranger.
 *
 * 4. Column H (theoretical quantity) is refilled ONLY when the lot
 *    changed. H belongs to the lot, and the library fills it from
 *    TODAY's Stock Poisson value. Refilling it after a
 *    quantity-only correction would silently replace the target
 *    recorded on the day of the entry, and change column I for a
 *    reason the worker never asked for.
 *
 * The operator identity comes from tracCurrentOperatorEmail() in
 * TracabiliteServer.js — the project already has exactly one way
 * to answer "who is this", and this screen uses it rather than
 * declaring a second one.
 ***************************************************************/

/** yyyy-MM-dd for a cell date, in the sheet's own timezone (that is what the dates in D mean). */
function corrDayKey_(value, tz) {
  if (!(value instanceof Date)) return "";
  return Utilities.formatDate(value, tz, "yyyy-MM-dd");
}

/** The days a correction is still allowed on, newest first. */
function corrAllowedDays_(tz) {
  const now = new Date();
  const days = [];
  for (let d = 0; d <= CFG.CORR_WINDOW_DAYS; d++) {
    days.push(Utilities.formatDate(new Date(now.getTime() - d * 86400000), tz, "yyyy-MM-dd"));
  }
  return days;
}

/**
 * Short fingerprint of one row's C:F values. Sent to the browser with
 * the row, and checked again at write time: if it changed, the row is
 * not the row the worker was looking at.
 */
function corrSignature_(vals) {
  const parts = vals.map(function (v) {
    if (v instanceof Date) return String(v.getTime());
    return String(v === null || v === undefined ? "" : v);
  });
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5, parts.join(""), Utilities.Charset.UTF_8);
  return bytes.slice(0, 6).map(function (b) {
    return ("0" + (b & 0xff).toString(16)).slice(-2);
  }).join("");
}

/** The audit tab. Created on first use; nothing else reads it. */
function corrLogSheet_(ss) {
  let sh = ss.getSheetByName(CFG.CORR_LOG_SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(CFG.CORR_LOG_SHEET);
  sh.appendRow(["Date/heure", "Opérateur", "Ligne", "Jour de la saisie",
                "Ancien lot", "Ancien type", "Ancienne qté",
                "Nouveau lot", "Nouveau type", "Nouvelle qté", "H recalculé"]);
  sh.setFrozenRows(1);
  return sh;
}

/**
 * Entries still inside the correction window, newest first.
 * Returns { days: [...], rows: [{ row, lot, date, type, qty, sig }] }.
 */
function corrListRecent() {
  const ss = SpreadsheetApp.openById(CFG.NOURRISSAGE_SS_ID);
  const sh = ss.getSheetByName(CFG.CONSO_SHEET);
  if (!sh) throw new Error('Sheet not found: "' + CFG.CONSO_SHEET + '"');

  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const allowed = corrAllowedDays_(tz);

  const lastRow = findNextConsoRow(sh) - 1;   // the app's one way to find the end of the data
  if (lastRow < CFG.CONSO_START_ROW) return { days: allowed, rows: [] };

  const first = Math.max(CFG.CONSO_START_ROW, lastRow - CFG.CORR_SCAN_ROWS + 1);
  const block = sh.getRange(first, CFG.CONSO_KEY_COL, lastRow - first + 1, 4).getValues(); // C:F

  const rows = [];
  for (let i = block.length - 1; i >= 0; i--) {
    const v = block[i];
    const day = corrDayKey_(v[1], tz);
    if (!day || allowed.indexOf(day) < 0) continue;
    rows.push({
      row: first + i,
      lot: String(v[0]),
      date: day,
      type: String(v[2]),
      qty: v[3],
      sig: corrSignature_(v)
    });
  }
  return { days: allowed, rows: rows };
}

/**
 * Apply one correction.
 * req: { row, sig, lot, type, qty } — lot/type/qty may be omitted to keep the current value.
 * Returns { row, sig, hRefilled }.
 */
function corrApply(req) {
  if (!req || !req.row || !req.sig) throw new Error("Requête incomplète.");
  const row = Number(req.row);
  if (!isFinite(row) || row < CFG.CONSO_START_ROW) throw new Error("Ligne invalide.");

  const email = tracCurrentOperatorEmail();   // throws if the user is not identified

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error("Le système est occupé. Réessayez dans un instant.");
  }
  try {
    const ss = SpreadsheetApp.openById(CFG.NOURRISSAGE_SS_ID);
    const sh = ss.getSheetByName(CFG.CONSO_SHEET);
    if (!sh) throw new Error('Sheet not found: "' + CFG.CONSO_SHEET + '"');
    const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();

    const cur = sh.getRange(row, CFG.CONSO_KEY_COL, 1, 4).getValues()[0]; // C:F
    if (corrSignature_(cur) !== String(req.sig)) {
      throw new Error("Cette ligne a changé depuis l'affichage. Rechargez l'écran, puis recommencez.");
    }

    const day = corrDayKey_(cur[1], tz);
    if (corrAllowedDays_(tz).indexOf(day) < 0) {
      throw new Error("Cette saisie est trop ancienne pour être corrigée ici. Prévenez Kim.");
    }

    const blank = function (x) { return x === undefined || x === null || x === ""; };
    const newLot  = blank(req.lot)  ? String(cur[0]) : String(req.lot);
    const newType = blank(req.type) ? String(cur[2]) : String(req.type);
    const newQty  = blank(req.qty)  ? Number(cur[3]) : Number(req.qty);

    if (!isFinite(newQty) || newQty <= 0) {
      throw new Error("La quantité doit être un nombre positif (reçu : " + req.qty + ").");
    }
    if (getLotList().map(String).indexOf(newLot) < 0) {
      throw new Error("Lot inconnu : " + newLot);
    }
    if (getFeedTypes().map(String).indexOf(newType) < 0) {
      throw new Error("Type de provende inconnu : " + newType);
    }

    const lotChanged = newLot !== String(cur[0]);
    if (!lotChanged && newType === String(cur[2]) && newQty === Number(cur[3])) {
      throw new Error("Rien n'a changé.");
    }

    // The audit tab is created BEFORE the write, so the one failure that
    // would leave a correction unrecorded happens before anything moves.
    const log = corrLogSheet_(ss);

    sh.getRange(row, CFG.CONSO_KEY_COL).setValue(newLot);   // C
    sh.getRange(row, CFG.CONSO_TYPE_COL).setValue(newType); // E
    sh.getRange(row, CFG.CONSO_F_COL).setValue(newQty);     // F
    SpreadsheetApp.flush();

    let hRefilled = false;
    if (lotChanged) {
      ConsoProvende.fillHForRows(row, row);   // same code path as a manual sheet edit
      hRefilled = true;
    }

    log.appendRow([new Date(), email, row, day,
                   String(cur[0]), String(cur[2]), cur[3],
                   newLot, newType, newQty, hRefilled ? "oui" : "non"]);

    const after = sh.getRange(row, CFG.CONSO_KEY_COL, 1, 4).getValues()[0];
    return { row: row, sig: corrSignature_(after), hRefilled: hRefilled };
  } finally {
    lock.releaseLock();
  }
}
