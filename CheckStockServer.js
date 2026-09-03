/***************************************************************
 * CheckStockServer.js — TSARA Entry web app
 * Screen 8b: Check stock physique (comptage des sacs)
 *
 * Appends ONE NEW COLUMN to the "Nourrissage & inventaire" file,
 * sheet "check stock physique":
 *   row 2  = date of the count
 *   rows 3-18 = bags counted, one per feed type in column A
 *
 * The contract is owned by WARNINGSYSTEM/07_checks_stockdiff.js,
 * which scans row 2 from column D rightward, skips column C, and
 * compares the FROZEN theoretical figures against the column with
 * the NEWEST date. Two consequences drive the code below:
 *
 *   1) A count dated before an existing column would be written
 *      and then ignored. Such a date is refused.
 *   2) An empty cell is skipped by the comparison
 *      (`if (ph === null) continue`), while 0 is compared. A feed
 *      type left blank therefore means "not counted", not "zero
 *      bags". Blanks are written as empty, never as 0.
 ***************************************************************/

const CS_CFG = {
  SS_ID: "1JBoH5c7BqZc2V5czcDAnEt-2hvkNKlJAxuDSupttTfs",
  SHEET: "check stock physique",
  DATE_ROW: 2,          // row 2 holds the date of each count column
  START_ROW: 3,         // feed-type rows
  END_ROW: 18,
  NAME_COL: 1,          // A = feed type
  THEO_COL: 3,          // C = frozen theoretical — skipped by the scan
  FIRST_PHYS_COL: 4,    // D = first count column

  SNAP_SHEET: "snapshots stock",   // hidden tab, written by WARNINGSYSTEM
  SNAP_DATE_COL: 1,                // A = date of the snapshot
  SNAP_STATUS_COL: 4,              // D = statut
  SNAP_STATUS_PENDING: "en attente",
  SNAP_STATUS_DONE: "compare",     // = WS_CFG.STOCK_SNAPSHOT.STATUS_DONE
  SNAP_ANSWERED_COL: 5,            // E = date of the count that closed the snapshot
  MAX_LAG_DAYS: 2                  // WARNINGSYSTEM voids a count later than this
};

/** Open the count sheet, or fail loudly. */
function openCheckStockSheet() {
  const sh = SpreadsheetApp.openById(CS_CFG.SS_ID).getSheetByName(CS_CFG.SHEET);
  if (!sh) throw new Error('Onglet introuvable: "' + CS_CFG.SHEET + '"');
  return sh;
}

/**
 * Newest count column. Scans row 2 from FIRST_PHYS_COL, skips THEO_COL,
 * and keeps a STRICTLY greater date, so two columns of one date resolve
 * to the leftmost — the exact rule WARNINGSYSTEM uses. Both projects
 * must select the same column or a correction lands on a column that is
 * never read.
 * Returns { col, date } or null.
 */
function findNewestCount(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < CS_CFG.FIRST_PHYS_COL) return null;

  const dateRow = sh.getRange(CS_CFG.DATE_ROW, 1, 1, lastCol).getValues()[0];
  let best = null;
  for (let c = CS_CFG.FIRST_PHYS_COL; c <= lastCol; c++) {
    if (c === CS_CFG.THEO_COL) continue;
    const v = dateRow[c - 1];
    if (Object.prototype.toString.call(v) !== "[object Date]") continue;
    if (best === null || v.getTime() > best.date.getTime()) best = { col: c, date: v };
  }
  return best;
}

/** Midnight of a date, so two counts of one day compare equal. */
function dayKeyOf(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

/**
 * Re-open the comparison after an in-place correction.
 *
 * WARNINGSYSTEM closes a snapshot with status "compare" and stamps the
 * date of the count that closed it into column E. A correction re-opens
 * ONLY the rows whose stamp equals the date of the column just corrected.
 *
 * The stamp is what makes this exact. Matching on the snapshot date
 * instead means guessing, and the guess re-opens a snapshot that an
 * earlier count already answered.
 *
 * Rows written before the stamp existed have an empty column E. They
 * match nothing, which is the safe default. A voided snapshot carries no
 * stamp either: a count that missed its window must not revive.
 *
 * The stamp is cleared on re-arming, so WARNINGSYSTEM writes a fresh one
 * when it compares again.
 * Returns true when at least one row was re-armed.
 */
function rearmSnapshotForCount(ss, countDate) {
  const snapSh = ss.getSheetByName(CS_CFG.SNAP_SHEET);
  if (!snapSh || snapSh.getLastRow() < 2) return false;

  const rows = snapSh.getLastRow() - 1;
  const sv = snapSh.getRange(2, 1, rows, CS_CFG.SNAP_ANSWERED_COL).getValues();
  const countKey = dayKeyOf(countDate);

  let n = 0;
  for (let i = 0; i < rows; i++) {
    const status = String(sv[i][CS_CFG.SNAP_STATUS_COL - 1] || "").trim().toLowerCase();
    if (status !== CS_CFG.SNAP_STATUS_DONE) continue;

    const answered = sv[i][CS_CFG.SNAP_ANSWERED_COL - 1];
    if (Object.prototype.toString.call(answered) !== "[object Date]") continue;
    if (dayKeyOf(answered) !== countKey) continue;

    snapSh.getRange(i + 2, CS_CFG.SNAP_STATUS_COL).setValue(CS_CFG.SNAP_STATUS_PENDING);
    snapSh.getRange(i + 2, CS_CFG.SNAP_ANSWERED_COL).clearContent();
    n++;
  }
  return n > 0;
}

/**
 * Everything the screen needs on load: the feed-type list read live
 * from column A, the date of the last count, and the open count
 * request if WARNINGSYSTEM has one waiting.
 *
 * The frozen theoretical figures are deliberately NOT returned. The
 * count must be blind, or the person counting is anchored by the
 * number the count exists to verify.
 */
function getCheckStockContext() {
  const sh = openCheckStockSheet();
  const ss = sh.getParent();
  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();

  const n = CS_CFG.END_ROW - CS_CFG.START_ROW + 1;
  const names = sh.getRange(CS_CFG.START_ROW, CS_CFG.NAME_COL, n, 1).getValues();
  const feeds = [];
  for (let i = 0; i < n; i++) {
    const nm = String(names[i][0] || "").trim();
    if (nm) feeds.push(nm);
  }

  const newest = findNewestCount(sh);

  // Figures of the newest column, so the screen can load them back for a
  // correction. The FROZEN THEORETICAL is still never returned: the person
  // sees only what he entered himself, never the number he is verifying.
  const lastCountValues = [];
  if (newest) {
    const pv = sh.getRange(CS_CFG.START_ROW, newest.col, n, 1).getValues();
    for (let i = 0; i < n; i++) {
      const nm = String(names[i][0] || "").trim();
      if (!nm) continue;
      const raw = pv[i][0];
      lastCountValues.push({ name: nm, value: (raw === "" || raw === null) ? "" : String(raw) });
    }
  }

  let pending = null;
  const snapSh = ss.getSheetByName(CS_CFG.SNAP_SHEET);
  if (snapSh && snapSh.getLastRow() > 1) {
    const sv = snapSh.getRange(2, 1, snapSh.getLastRow() - 1, CS_CFG.SNAP_STATUS_COL).getValues();
    let snapDate = null;
    for (let i = 0; i < sv.length; i++) {
      const status = String(sv[i][CS_CFG.SNAP_STATUS_COL - 1] || "").trim().toLowerCase();
      if (status !== CS_CFG.SNAP_STATUS_PENDING) continue;
      const d = sv[i][CS_CFG.SNAP_DATE_COL - 1];
      if (Object.prototype.toString.call(d) !== "[object Date]") continue;
      if (snapDate === null || d.getTime() > snapDate.getTime()) snapDate = d;
    }
    if (snapDate) {
      const today = new Date();
      const ageDays = Math.round(
        (new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime() -
         new Date(snapDate.getFullYear(), snapDate.getMonth(), snapDate.getDate()).getTime()) / 86400000
      );
      pending = {
        date: Utilities.formatDate(snapDate, tz, "dd/MM/yyyy"),
        ageDays: ageDays,
        late: ageDays > CS_CFG.MAX_LAG_DAYS
      };
    }
  }

  return {
    feeds: feeds,
    lastCountDate: newest ? Utilities.formatDate(newest.date, tz, "dd/MM/yyyy") : "",
    lastCountDateISO: newest ? Utilities.formatDate(newest.date, tz, "yyyy-MM-dd") : "",
    lastCountValues: lastCountValues,
    pending: pending,
    maxLagDays: CS_CFG.MAX_LAG_DAYS
  };
}

/**
 * Write one count column — ONE COLUMN PER DATE.
 *
 * A date that already has a column OVERWRITES that column, so a miscount
 * is corrected in place. A second column for the same date would be
 * ignored: WARNINGSYSTEM keeps the leftmost of two equal dates.
 *
 * payload: { date: "yyyy-MM-dd", counts: [{ name, value }] }
 * A count with an empty value is written as an empty cell, so the
 * comparison skips that feed type instead of reading it as zero.
 *
 * Counts are matched to rows BY NAME, never by position, so a row
 * inserted in the sheet after the screen loaded cannot shift a
 * count onto the wrong feed type.
 *
 * Returns { column, date, filled }.
 */
function submitCheckStock(payload) {
  if (!payload) throw new Error("Aucune donnée reçue.");

  const dateStr = String(payload.date || "").trim();
  if (!dateStr) throw new Error("Date du comptage obligatoire.");

  const counts = payload.counts || [];
  if (!counts.length) throw new Error("Aucun comptage saisi.");

  const sh = openCheckStockSheet();
  const tz = sh.getParent().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const countDate = new Date(dateStr);
  if (isNaN(countDate.getTime())) throw new Error("Date invalide (reçu: " + dateStr + ").");

  const newest = findNewestCount(sh);
  if (newest && dayKeyOf(countDate) < dayKeyOf(newest.date)) {
    throw new Error(
      "Date trop ancienne. Un comptage du " + Utilities.formatDate(newest.date, tz, "dd/MM/yyyy") +
      " existe déjà, et le système compare toujours le comptage le plus récent. " +
      "Une colonne plus ancienne serait ignorée."
    );
  }

  const n = CS_CFG.END_ROW - CS_CFG.START_ROW + 1;
  const names = sh.getRange(CS_CFG.START_ROW, CS_CFG.NAME_COL, n, 1).getValues();
  const rowIndexByName = {};
  for (let i = 0; i < n; i++) {
    const nm = String(names[i][0] || "").trim();
    if (nm) rowIndexByName[nm] = i;
  }

  const values = [];
  for (let i = 0; i < n; i++) values.push([""]);

  let filled = 0;
  for (let k = 0; k < counts.length; k++) {
    const nm = String(counts[k].name || "").trim();
    if (!nm) continue;
    if (!(nm in rowIndexByName)) {
      throw new Error('Type de provende absent de l\'onglet: "' + nm + '". Rechargez l\'écran.');
    }
    const raw = counts[k].value;
    if (raw === "" || raw === null || raw === undefined) continue;
    const v = Number(raw);
    if (!isFinite(v) || v < 0) {
      throw new Error("Comptage invalide pour " + nm + " (reçu: " + raw + ").");
    }
    values[rowIndexByName[nm]] = [v];
    filled++;
  }

  if (!filled) throw new Error("Aucun comptage saisi.");

  const isUpdate = !!(newest && dayKeyOf(countDate) === dayKeyOf(newest.date));
  const col = isUpdate ? newest.col : sh.getLastColumn() + 1;

  sh.getRange(CS_CFG.DATE_ROW, col).setValue(countDate);
  sh.getRange(CS_CFG.START_ROW, col, n, 1).setValues(values);

  const rearmed = isUpdate ? rearmSnapshotForCount(sh.getParent(), countDate) : false;

  return {
    column: col,
    date: Utilities.formatDate(countDate, tz, "dd/MM/yyyy"),
    filled: filled,
    updated: isUpdate,
    rearmed: rearmed
  };
}
