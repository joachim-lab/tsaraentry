/***************************************************************
 * MortsServer.js — TSARA Entry web app
 * Screen 4: Morts (mortalité) — server side.
 *
 * Appends rows to the "Morts" tab inside the Commandes spreadsheet —
 * the same tab tt_processMortsWorker_ reads (engine_core.js,
 * TT_MORTS_SHEET_NAME = "Morts", CFG.COMMANDES_SS_ID). This screen
 * never writes to a lot file directly; the nightly engine does the
 * actual deduction, exactly as it does for Commandes.
 *
 * Column layout A:F (engine_core.js, live-verified 2026-08-13,
 * 112,644 bytes):
 *   A Lot   B Date   C Qty   D Cause   E Log   F Erreur
 * A row is only processed by the engine when E and F are both empty.
 *
 * VALIDATION DIFFERS FROM COMMANDES: the engine's Morts path does
 * NOT check reservations (tt_processMortsWorker_ has no reservation
 * read at all — a death removes fish whether or not they are held
 * for a customer). So the block rule here is simply
 *   qty > available   -> BLOCK
 * with no TOUT check and no PM/type gate (deaths have no AL/GR type).
 *
 * Reuses, unchanged, from elsewhere in this project:
 *   cmdCanonKey, findSubLotColumnByOrderKey   (Lot.js)
 *   getLotFileList                            (Lot.js)
 *   cmdToNum                                  (CommandesServer.js)
 *   CMD_CFG.SS_ID                             (CommandesServer.js)
 ***************************************************************/

const MORTS_CFG = {
  SS_ID: CMD_CFG.SS_ID,   // Commandes spreadsheet — same file the engine reads
  SHEET: "Morts",
  START_ROW: 2,
  COL: { LOT: 1, DATE: 2, QTY: 3, CAUSE: 4, LOG: 5, ERROR: 6 }
};

function mortsSheet() {
  const ss = SpreadsheetApp.openById(MORTS_CFG.SS_ID);
  const sh = ss.getSheetByName(MORTS_CFG.SHEET);
  if (!sh) throw new Error('Onglet introuvable: "' + MORTS_CFG.SHEET + '"');
  return sh;
}

/**
 * Same yyyy-MM-dd -> local-midnight parser as Commandes (cmdParseDate).
 * Never new Date(str) directly — that parses as UTC midnight, which
 * shows as 03:00 once Sheets renders it in Africa/Nairobi (UTC+3).
 */
function mortsParseDate(isoStr) {
  if (!isoStr) return undefined;
  const m = String(isoStr).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return undefined;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/**
 * First row after the last real entry, scanning column A (lot).
 * Same defensive pattern as Nourrissage/Commandes — getLastRow() is
 * not trusted anywhere in this project; two other sheets in this same
 * spreadsheet family are known to overshoot it.
 */
function findNextMortsRow(sh) {
  const physical = sh.getLastRow();
  if (physical < MORTS_CFG.START_ROW) return MORTS_CFG.START_ROW;
  const vals = sh.getRange(MORTS_CFG.START_ROW, MORTS_CFG.COL.LOT,
    physical - MORTS_CFG.START_ROW + 1, 1).getDisplayValues();
  let lastData = MORTS_CFG.START_ROW - 1;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || "").trim() !== "") lastData = MORTS_CFG.START_ROW + i;
  }
  return lastData + 1;
}

/**
 * Lot dropdown: read live off the sheet's own validation rule on
 * column A — same technique as getFeedTypes (Code.js) and
 * cmdGetOptions (CommandesServer.js). tt_applyMortsDropdown_
 * (engine_core.js) is what keeps that rule current; this never
 * hardcodes the list.
 */
function mortsGetOptions() {
  const sh = mortsSheet();
  const probeRow = Math.max(MORTS_CFG.START_ROW, findNextMortsRow(sh) - 1);
  const rule = sh.getRange(probeRow, MORTS_CFG.COL.LOT).getDataValidation();
  if (!rule) return { lots: [] };
  try {
    const cv = rule.getCriteriaValues();
    return { lots: Array.isArray(cv[0]) ? cv[0] : [] };
  } catch (e) {
    return { lots: [] };
  }
}

/**
 * Sum of Morts rows already in the sheet for this key that the engine
 * has not yet processed (E and F both empty). Mirrors
 * cmdGetPendingQty's logic (CommandesServer.js), applied to the Morts
 * columns — same reason it exists: this validates BEFORE any
 * deduction has happened, so an already-entered-but-unprocessed death
 * must count against availability or a second entry could pass
 * against fish that are already spoken for.
 */
function mortsGetPendingQty(canonKey) {
  const sh = mortsSheet();
  const lastRow = findNextMortsRow(sh) - 1;
  if (lastRow < MORTS_CFG.START_ROW) return 0;

  const data = sh.getRange(MORTS_CFG.START_ROW, 1,
    lastRow - MORTS_CFG.START_ROW + 1, 6).getValues();
  var total = 0;
  for (var i = 0; i < data.length; i++) {
    const r = data[i];
    if (cmdCanonKey(r[MORTS_CFG.COL.LOT - 1]) !== canonKey) continue;
    if (String(r[MORTS_CFG.COL.LOG - 1]   || "").trim() !== "") continue; // E
    if (String(r[MORTS_CFG.COL.ERROR - 1] || "").trim() !== "") continue; // F
    const q = cmdToNum(r[MORTS_CFG.COL.QTY - 1]);
    if (q != null && q > 0) total += q;
  }
  return total;
}

/**
 * Stock picture for one lot key, as the engine's Morts path would see
 * it. No reservation term — see file header.
 *
 * @param {string} rawKey
 * @return {Object} { key, found, source, col|row, count, pm, pending, available }
 */
function mortsGetLotAvailability(rawKey) {
  const key = cmdCanonKey(rawKey);
  const out = {
    key: key, found: false, source: null, count: null, pm: null,
    pending: 0, available: null
  };
  if (!key) return out;

  const lotNum = key.split("-")[0];
  const list = getLotFileList();
  var fileId = null;
  for (var i = 0; i < list.length; i++) {
    if (cmdCanonKey(list[i].lotNumber) === lotNum) { fileId = list[i].fileId; break; }
  }
  if (!fileId) return out; // found stays false -> caller warns

  const m = findSubLotColumnByOrderKey(SpreadsheetApp.openById(fileId), key);
  if (!m.found) return out;

  out.found = true;
  out.source = m.source;
  if (m.col) out.col = m.col;
  if (m.row) out.row = m.row;
  out.count = m.count;
  out.pm = m.pm;

  out.pending = mortsGetPendingQty(key);
  out.available = out.count - out.pending;
  return out;
}

/**
 * Validate a whole submission before saving.
 *
 * Entries are summed BY CANON KEY, not checked one at a time: a
 * batch can list the same lot on two rows, and those rows are not
 * yet in the sheet so mortsGetPendingQty cannot see them. Checking
 * row by row would let two 50-fish entries both pass against a stock
 * of 80. Same reasoning as cmdValidateOrderLines.
 *
 * @param {Array} entries  [{ lot, qty, ... }]
 * @return {Object} { ok, blocks: [msg], warnings: [msg], detail: {key: avail} }
 */
function mortsValidateEntries(entries) {
  const blocks = [];
  const warnings = [];
  const detail = {};
  if (!entries || !entries.length) return { ok: true, blocks: blocks, warnings: warnings, detail: detail };

  const wanted = {};
  entries.forEach(function (en) {
    const k = cmdCanonKey(en.lot);
    if (!k) return;
    const q = cmdToNum(en.qty);
    wanted[k] = (wanted[k] || 0) + (q != null && q > 0 ? q : 0);
  });

  Object.keys(wanted).forEach(function (k) {
    const a = mortsGetLotAvailability(k);
    detail[k] = a;

    if (!a.found) {
      warnings.push(k + " : ce lot n'est pas trouvé dans le fichier lot ; " +
                        "l'entrée sera signalée (NOT FOUND) cette nuit");
      return;
    }
    if (wanted[k] > a.available) {
      blocks.push(k + " : stock insuffisant — mortalité déclarée " + Math.round(wanted[k]) +
                      ", disponible " + a.available +
                      " (stock " + a.count +
                      (a.pending ? ", déjà en attente " + a.pending : "") + ")");
    }
  });

  return { ok: blocks.length === 0, blocks: blocks, warnings: warnings, detail: detail };
}

/**
 * Append one or more death entries to A:D. Never touches E/F — those
 * belong to the engine.
 *
 * entries: [{ lot, date, qty, cause }], date is "yyyy-MM-dd" from the browser.
 * Returns { written, startRow, endRow }.
 */
function mortsSubmit(entries) {
  if (!entries || !entries.length) throw new Error("Aucune entrée à enregistrer.");

  const sh = mortsSheet();

  const rows = entries.map(function (en) {
    if (!en.lot || !en.date || en.qty === undefined || en.qty === null || en.qty === "") {
      throw new Error("Entrée incomplète : lot, date et quantité sont obligatoires.");
    }
    const qty = cmdToNum(en.qty);
    if (qty == null || qty <= 0) {
      throw new Error("Quantité invalide (reçu: " + en.qty + ").");
    }
    return [en.lot, mortsParseDate(en.date), qty, en.cause || ""];
  });

  const startRow = findNextMortsRow(sh);
  sh.getRange(startRow, MORTS_CFG.COL.LOT, rows.length, 4).setValues(rows); // A:D
  const endRow = startRow + rows.length - 1;

  return { written: rows.length, startRow: startRow, endRow: endRow };
}

/** RUN FROM EDITOR: read-only checks of the Morts server side. */
function testMortsServer() {
  const sh = mortsSheet();
  Logger.log("Prochaine ligne libre: " + findNextMortsRow(sh));

  const opts = mortsGetOptions();
  Logger.log("Lots (" + opts.lots.length + "): " + opts.lots.slice(0, 8).join(" | "));

  if (opts.lots.length) {
    const sample = opts.lots[0];
    Logger.log("--- disponibilité pour " + sample + " ---");
    Logger.log(JSON.stringify(mortsGetLotAvailability(sample), null, 2));
  }
}
