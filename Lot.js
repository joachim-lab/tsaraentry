/***************************************************************
 * Lot.js — TSARA Entry web app
 * Screen 2: Lot (échantillonnage) — stage detection + read only.
 *
 * No writes in this file yet. Establishes, for a given lot:
 *   - which stage it's at (1-5 / 6-10 / 11-15 / 16-21 / S1 / S2-Tri /
 *     S3..S20 / Grossissement)
 *   - the exact target (tab name, or Grossissement block) to write
 *     into next
 *   - whether the lot is in a "mixed" state (Grossissement has data
 *     AND an S-tab still has active sub-lots) — flagged, not fixed;
 *     mirrors the live engine's own blind spot (engine_core.js,
 *     "if (grossFound) return" / "do not touch S-sheets if Gross
 *     matched" — verified live 2026-08-09).
 *
 * LOT FILE LOOKUP: reads the lots folder directly (file titles like
 * "Lot-21", "Lot-Mirana"), not derived from Stock Poisson's sub-lot
 * list — the folder is the source of truth for which lot files exist.
 ***************************************************************/

const LOT_CFG = {
  LOTS_FOLDER_ID: "1VeO6WrWjjb0QU6tZJ2VTLUvzYF4Fm3iJ", // 2025 lot files only

  // Grossissement layout (verified live on Lot-21, 2026-08-09)
  GROSS_SHEET: "Grossissement",
  GROSS_START_COL: 15,   // O
  GROSS_END_COL: 101,    // matches buildRecapBlock_'s own range, live-verified
  GROSS_GROUP_SIZE: 4,
  GROSS_ROW_DATE: 5,
  GROSS_ROW_TEMP: 6,
  GROSS_ROW_BASSIN: 7,
  GROSS_ROW_HAPPA: 8,
  GROSS_ROW_NOMBRE: 10,
  GROSS_ROW_PM: 11,

  // S-tab scan order — IDENTICAL to the engine's own sSheetNames
  // (engine_core.js, live-verified 2026-08-09), so "which tab is
  // active" always agrees with what the engine itself would find.
  S_SHEET_ORDER: [
    "S20","S19","S18","S17","S16","S15","S14","S13","S12","S11",
    "S10","S9","S8","S7","S6","S5","S4","S3","S2-Tri","S1",
    "16-21","11-15","6-10","1-5"
  ]
};

/** List available lot files: [{ lotNumber, fileId, fileName }]. */
function getLotFileList() {
  const folder = DriveApp.getFolderById(LOT_CFG.LOTS_FOLDER_ID);
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  const out = [];
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();
    const m = name.match(/^Lot-(.+)$/);
    if (!m) continue; // skip anything that isn't a lot file
    out.push({ lotNumber: m[1], fileId: f.getId(), fileName: name });
  }
  out.sort((a, b) => a.lotNumber.localeCompare(b.lotNumber, undefined, { numeric: true }));
  return out;
}

/**
 * Does Grossissement have any data? Mirrors the engine's own scan
 * exactly (right-to-left through row 7, first non-blank bassin wins) —
 * same rule as tt_ helpers in engine_core.js, not a reinterpretation.
 * Returns { found, lastFilledGroupStart } (0-based offset from GROSS_START_COL).
 */
function scanGrossissement_(sh) {
  const numCols = LOT_CFG.GROSS_END_COL - LOT_CFG.GROSS_START_COL + 1;
  const row7 = sh.getRange(LOT_CFG.GROSS_ROW_BASSIN, LOT_CFG.GROSS_START_COL, 1, numCols)
    .getDisplayValues()[0];

  for (let i = numCols - 1; i >= 0; i--) {
    if (String(row7[i] || "").trim() !== "") {
      const groupStart = Math.floor(i / LOT_CFG.GROSS_GROUP_SIZE) * LOT_CFG.GROSS_GROUP_SIZE;
      return { found: true, lastFilledGroupStart: groupStart };
    }
  }
  return { found: false, lastFilledGroupStart: -1 };
}

/**
 * Is THIS WEEK'S S-tab (by date, not history) showing an active
 * sub-lot? Used only for the mixed-lot warning.
 *
 * NOTE: A11:A16 on an S-tab stays non-blank forever once a sub-lot
 * has ever passed through it — B1 feeds from '1-5'!$B$1, the lot's
 * own number, set once at creation and never cleared. Every lot has
 * this kind of history on its early tabs; that is normal, not a
 * mixed-lot situation. The only meaningful check is whether the
 * CURRENT-by-date tab still has something live on it — so this
 * reuses findCurrentSTab_ rather than scanning all 24 tabs.
 */
function scanActiveSSheets_(ss) {
  const tabName = findCurrentSTab_(ss);
  if (!tabName) return [];

  const sh = ss.getSheetByName(tabName);
  const labels = sh.getRange("A11:A16").getDisplayValues();
  const ids = labels.map(r => String(r[0] || "").trim()).filter(v => v !== "");
  return ids.length ? [{ sheet: tabName, subLots: ids }] : [];
}

/**
 * Find "this week's" S-tab: the tab whose computed row-8 date range
 * contains today. Reads computed values only, never formulas or a
 * stored marker — so manual date overrides (which re-chain the whole
 * sheet) are automatically respected.
 */
function findCurrentSTab_(ss) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Walk in chronological order (reverse of the engine's scan order,
  // which is newest-first) so we can stop at the first range containing today.
  const chrono = LOT_CFG.S_SHEET_ORDER.slice().reverse();

  for (const name of chrono) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const lastCol = (name === "16-21") ? 7          // B..G
      : (name.match(/^S\d+$|^S2-Tri$/)) ? 8         // B..H
      : 6;                                           // 1-5/6-10/11-15: B..F

    const rowVals = sh.getRange(8, 2, 1, lastCol - 1).getDisplayValues()[0];
    const dates = rowVals
      .map(v => (v && v !== "-") ? new Date(v) : null)
      .filter(d => d && !isNaN(d.getTime()));

    if (!dates.length) continue;

    const minD = new Date(Math.min.apply(null, dates));
    const maxD = new Date(Math.max.apply(null, dates));
    minD.setHours(0, 0, 0, 0);
    maxD.setHours(0, 0, 0, 0);

    if (today >= minD && today <= maxD) {
      return name;
    }
  }
  return null; // no tab's range contains today — needs a fallback decision
}

/**
 * Main entry point: given a lot's fileId, determine what to show.
 * Returns:
 *   { stage: "grossissement", targetGroupStart, mixedWarning }
 *   { stage: "s-tab", tabName, mixedWarning: null }
 *   { stage: "none", reason }
 */
function getLotStage(fileId) {
  const ss = SpreadsheetApp.openById(fileId);
  const gross = ss.getSheetByName(LOT_CFG.GROSS_SHEET);
  if (!gross) return { stage: "none", reason: "Grossissement sheet not found" };

  const grossScan = scanGrossissement_(gross);

  if (grossScan.found) {
    const activeS = scanActiveSSheets_(ss);
    return {
      stage: "grossissement",
      targetGroupStart: grossScan.lastFilledGroupStart + LOT_CFG.GROSS_GROUP_SIZE,
      mixedWarning: activeS.length
        ? "This lot also has active sub-lot(s) on S-tab(s): " +
          activeS.map(a => a.sheet + " (" + a.subLots.join(", ") + ")").join("; ") +
          ". These will NOT reach Stock Poisson — same behaviour as the engine today."
        : null
    };
  }

  const tabName = findCurrentSTab_(ss);
  if (!tabName) {
    return { stage: "none", reason: "La date introduite n'existe pas pour ce lot" };
  }
  return { stage: "s-tab", tabName: tabName, mixedWarning: null };
}

/**
 * RUN FROM EDITOR: manual test for the two functions above.
 * Logs the full lot file list, then getLotStage() for the first
 * three lots in that list (enough to see both stage types without
 * needing to pick a fileId by hand).
 */
function testLotStage() {
  const lots = getLotFileList();
  Logger.log("Lot file list (" + lots.length + " lots):");
  Logger.log(JSON.stringify(lots, null, 2));

  const sample = lots.slice(0, 3);
  sample.forEach(lot => {
    const result = getLotStage(lot.fileId);
    Logger.log("--- " + lot.fileName + " (" + lot.fileId + ") ---");
    Logger.log(JSON.stringify(result, null, 2));
  });
}

/** RUN FROM EDITOR: one-line stage summary for every lot in the folder. */
function testAllLotStages() {
  const lots = getLotFileList();
  lots.forEach(lot => {
    const result = getLotStage(lot.fileId);
    Logger.log(lot.fileName + ": " + result.stage +
      (result.stage === "s-tab" ? " (" + result.tabName + ")" : "") +
      (result.stage === "none" ? " — " + result.reason : ""));
  });
}

/**
 * DIAGNOSTIC — run for a lot that returned "none" from getLotStage.
 * Dumps raw Grossissement row-7 values and, for every S-tab, the
 * raw row-8 date values, so we can see WHY neither scan matched.
 */
function debugLotRaw(fileId) {
  const ss = SpreadsheetApp.openById(fileId);
  Logger.log("=== " + ss.getName() + " ===");

  const gross = ss.getSheetByName(LOT_CFG.GROSS_SHEET);
  if (gross) {
    const numCols = LOT_CFG.GROSS_END_COL - LOT_CFG.GROSS_START_COL + 1;
    const row7 = gross.getRange(LOT_CFG.GROSS_ROW_BASSIN, LOT_CFG.GROSS_START_COL, 1, numCols)
      .getDisplayValues()[0];
    const nonBlank = row7.map((v, i) => v ? (i + LOT_CFG.GROSS_START_COL) + "=" + v : null).filter(Boolean);
    Logger.log("Grossissement row7 non-blank cols (1-indexed): " + JSON.stringify(nonBlank));
  } else {
    Logger.log("No Grossissement sheet found.");
  }

  LOT_CFG.S_SHEET_ORDER.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { Logger.log(name + ": sheet not found"); return; }
    const lastCol = (name === "16-21") ? 7 : (name.match(/^S\d+$|^S2-Tri$/)) ? 8 : 6;
    const row8 = sh.getRange(8, 2, 1, lastCol - 1).getDisplayValues()[0];
    Logger.log(name + " row8: " + JSON.stringify(row8));
  });
}
