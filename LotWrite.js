/***************************************************************
 * LotWrite.js — TSARA Entry web app
 * Screen 2: Lot (échantillonnage) — write field values.
 *
 * SAFETY MODEL (same as ROLLOUT_dateChain.gs): every submission is
 * first turned into a PLAN of individual cell changes (from -> to).
 * Nothing is written unless applyLotEntry is called explicitly.
 * submitLotEntry defaults to DRY RUN.
 *
 * NEVER WRITTEN (derived cells — the sheet owns these):
 *   - A11:A16 sub-lot labels (formula from B1/B2/B3 etc.)
 *   - S2-Tri  B11:B16 Nombre      (=IFERROR(D/C,""))
 *   - S3..S20 D11:D16 Biomasse    (=C*B)
 *   - Grossissement row 12 Biomasse
 *   - Row 8 / row 5 date chains, EXCEPT the explicit anchors a user
 *     may set: 1-5!B5, S2-Tri!B7, Grossissement block date (row 5).
 *
 * Field list is Kim's confirmed live list (2026-08-09).
 ***************************************************************/

const WRITE_CFG = {
  SUBLOT_FIRST_ROW: 11,
  SUBLOT_LAST_ROW: 16,

  // Sample column + row extent per tab (Kim's confirmed ranges).
  SAMPLE_RANGES: {
    "1-5":   { col: 10, start: 5, end: 57 },  // J
    "6-10":  { col: 10, start: 6, end: 55 },
    "11-15": { col: 10, start: 6, end: 55 },
    "16-21": { col: 10, start: 6, end: 55 },
    "S1":    { col: 11, start: 5, end: 54 }   // K
  }
};

/**
 * Build the list of intended cell changes. Pure — reads only.
 * Returns [{ sheet, row, col, a1, from, to, note }]
 */
function buildLotWritePlan(fileId, stageResult, payload) {
  const ss = SpreadsheetApp.openById(fileId);
  const plan = [];

  if (stageResult.stage === "grossissement") {
    planGrossissement(ss, stageResult.targetGroupStartCol, payload, plan);
  } else if (stageResult.stage === "s-tab") {
    const tabName = stageResult.tabName;
    const sh = ss.getSheetByName(tabName);
    if (!sh) throw new Error('Tab not found: "' + tabName + '"');

    if (tabName === "S2-Tri") planS2Tri(sh, tabName, payload, plan);
    else if (tabName === "S1") planS1(sh, tabName, payload, plan);
    else if (tabName.match(/^S\d+$/)) planS3Plus(sh, tabName, payload, plan);
    else planEarlyTab(sh, tabName, payload, plan);
  } else {
    throw new Error("Cannot write for stage: " + stageResult.stage);
  }

  return plan;
}

/** Add one intended change, recording the current value for the dry run. */
function addChange(sh, row, col, value, note, plan) {
  if (value === undefined) return;             // field not supplied -> leave alone
  const cell = sh.getRange(row, col);
  const from = cell.getValue();
  const to = (value === null) ? "" : value;
  if (String(from) === String(to)) return;     // no-op, don't clutter the plan
  plan.push({
    sheet: sh.getName(),
    row: row, col: col,
    a1: cell.getA1Notation(),
    from: from, to: to,
    note: note
  });
}

/** 1-5 / 6-10 / 11-15 / 16-21 — single sub-lot (row 11 only). */
function planEarlyTab(sh, tabName, payload, plan) {
  const f = payload.fields || {};

  if (tabName === "1-5") {
    addChange(sh, 1, 2, f.B1, "Numéro de lot", plan);
    addChange(sh, 2, 2, f.B2, "Bassin", plan);
    addChange(sh, 3, 2, f.B3, "Happa", plan);
    addChange(sh, 5, 2, f.B5, "Date de départ (ancre la chaîne de dates)", plan);
  }
  addChange(sh, 4, 2, f.B4, "Température de l'eau", plan);
  addChange(sh, 11, 4, f.D11, "Nombre", plan);
  addChange(sh, 11, 7, f.G11, "Commentaires", plan);

  planSamples(sh, tabName, f.samples, plan);
}

/** S1 — single sub-lot, samples in column K, comments in H11. */
function planS1(sh, tabName, payload, plan) {
  const f = payload.fields || {};
  addChange(sh, 4, 2, f.B4, "Température de l'eau", plan);
  addChange(sh, 11, 8, f.H11, "Commentaires", plan);
  planSamples(sh, tabName, f.samples, plan);
}

/**
 * S2-Tri — the sort event. Writes bassin/happa dropdowns for new
 * sub-lots (C2:D2 / C3:D3), the tri date (B7), and per-row Poids
 * moyen (C) + Biomasse (D) + Commentaires (H).
 * Nombre (col B) is DERIVED here (=IFERROR(D/C,"")) — never written.
 */
function planS2Tri(sh, tabName, payload, plan) {
  const f = payload.fields || {};

  addChange(sh, 2, 2, f.B2, "Bassin (sous-lot 1)", plan);
  addChange(sh, 3, 2, f.B3, "Happa (sous-lot 1)", plan);
  addChange(sh, 4, 2, f.B4, "Température de l'eau", plan);
  addChange(sh, 7, 2, f.B7, "Date de tri", plan);

  addChange(sh, 2, 3, f.C2, "Bassin (sous-lot 2)", plan);
  addChange(sh, 2, 4, f.D2, "Bassin (sous-lot 3)", plan);
  addChange(sh, 3, 3, f.C3, "Happa (sous-lot 2)", plan);
  addChange(sh, 3, 4, f.D3, "Happa (sous-lot 3)", plan);

  (f.subLots || []).forEach(sl => {
    if (!isValidSubLotRow(sl.row)) throw new Error("Invalid sub-lot row: " + sl.row);
    addChange(sh, sl.row, 3, sl.poidsMoyen, "Poids moyen", plan);
    addChange(sh, sl.row, 4, sl.biomasse, "Biomasse", plan);
    addChange(sh, sl.row, 8, sl.commentaire, "Commentaires", plan);
  });
}

/**
 * S3..S20 — per-row Nombre (B) + Poids moyen (C) + Commentaires (I).
 * Biomasse (col D) is DERIVED (=C*B) — never written.
 */
function planS3Plus(sh, tabName, payload, plan) {
  const f = payload.fields || {};

  addChange(sh, 2, 2, f.B2, "Bassin", plan);
  addChange(sh, 3, 2, f.B3, "Happa", plan);
  addChange(sh, 4, 2, f.B4, "Température de l'eau", plan);

  (f.subLots || []).forEach(sl => {
    if (!isValidSubLotRow(sl.row)) throw new Error("Invalid sub-lot row: " + sl.row);
    addChange(sh, sl.row, 2, sl.nombre, "Nombre", plan);
    addChange(sh, sl.row, 3, sl.poidsMoyen, "Poids moyen", plan);
    addChange(sh, sl.row, 9, sl.commentaire, "Commentaires", plan);
  });
}

/**
 * Grossissement — writes into the target block only.
 * date + temperature go in the block's FIRST column (shared);
 * bassin/happa/nombre/poids moyen are per column (up to 4 sub-lots).
 * Biomasse (row 12) is DERIVED — never written.
 */
function planGrossissement(ss, targetGroupStartCol, payload, plan) {
  const sh = ss.getSheetByName(LOT_CFG.GROSS_SHEET);
  if (!sh) throw new Error("Grossissement sheet not found");
  const f = payload.fields || {};

  addChange(sh, LOT_CFG.GROSS_ROW_DATE, targetGroupStartCol, f.date,
    "Date du bloc (ré-ancre les dates suivantes)", plan);
  addChange(sh, LOT_CFG.GROSS_ROW_TEMP, targetGroupStartCol, f.temperature,
    "Température de l'eau", plan);

  const lastCol = targetGroupStartCol + LOT_CFG.GROSS_GROUP_SIZE - 1;
  (f.subLots || []).forEach(sl => {
    if (sl.col < targetGroupStartCol || sl.col > lastCol) {
      throw new Error("Sub-lot column " + sl.col + " is outside the target block (" +
        targetGroupStartCol + "-" + lastCol + ")");
    }
    addChange(sh, LOT_CFG.GROSS_ROW_BASSIN, sl.col, sl.bassin, "Bassin/Cage", plan);
    addChange(sh, LOT_CFG.GROSS_ROW_HAPPA, sl.col, sl.happa, "Happa/Carré/Rond", plan);
    addChange(sh, LOT_CFG.GROSS_ROW_NOMBRE, sl.col, sl.nombre, "Nombre", plan);
    addChange(sh, LOT_CFG.GROSS_ROW_PM, sl.col, sl.poidsMoyen, "Poids moyen", plan);
  });
}

/**
 * Samples (individual weights). The submitted list is the COMPLETE
 * intended list for this tab — the form was pre-filled with whatever
 * was already there — so any remaining cells in the range are cleared.
 * That keeps one correct state rather than merging old and new.
 */
function planSamples(sh, tabName, samples, plan) {
  if (samples === undefined) return;
  const r = WRITE_CFG.SAMPLE_RANGES[tabName];
  if (!r) throw new Error("No sample range configured for tab: " + tabName);

  const capacity = r.end - r.start + 1;
  if (samples.length > capacity) {
    throw new Error("Trop d'échantillons: " + samples.length + " (maximum " + capacity + ")");
  }

  for (let i = 0; i < capacity; i++) {
    const value = (i < samples.length) ? samples[i] : null;
    addChange(sh, r.start + i, r.col, value, "Échantillon " + (i + 1), plan);
  }
}

function isValidSubLotRow(row) {
  return row >= WRITE_CFG.SUBLOT_FIRST_ROW && row <= WRITE_CFG.SUBLOT_LAST_ROW;
}

/** Render a plan as readable text (for the dry run log / UI preview). */
function renderLotWritePlan(plan) {
  if (!plan.length) return "Aucun changement.";
  const lines = plan.map(c =>
    "  " + c.sheet + "!" + c.a1 + "  [" + c.note + "]  " +
    JSON.stringify(c.from) + " -> " + JSON.stringify(c.to)
  );
  lines.push("");
  lines.push("TOTAL: " + plan.length + " cellule(s) à modifier.");
  return lines.join("\n");
}

/**
 * MAIN ENTRY POINT. Defaults to DRY RUN — writes nothing unless
 * opts.dryRun is explicitly false.
 * Returns { dryRun, changeCount, plan, rendered }.
 */
function submitLotEntry(fileId, stageResult, payload, opts) {
  const dryRun = !(opts && opts.dryRun === false);
  const plan = buildLotWritePlan(fileId, stageResult, payload);

  if (!dryRun && plan.length) {
    const ss = SpreadsheetApp.openById(fileId);
    plan.forEach(c => {
      ss.getSheetByName(c.sheet).getRange(c.row, c.col).setValue(c.to);
    });
    SpreadsheetApp.flush();
  }

  return {
    dryRun: dryRun,
    changeCount: plan.length,
    plan: plan,
    rendered: renderLotWritePlan(plan)
  };
}

/** RUN FROM EDITOR: dry-run test on Lot-24 (S7) — writes NOTHING. */
function testWriteDryRun() {
  const fileId = '1_qwbM0Hma6cVWJ3lr7jtVtcgId7hUC71kCFjYg_nCmE'; // Lot-24, S7
  const stageResult = getLotStage(fileId);
  Logger.log("Stage: " + JSON.stringify(stageResult));

  const payload = {
    fields: {
      B4: 26,
      subLots: [
        { row: 11, nombre: 12313, poidsMoyen: 0.52, commentaire: "test dry run" }
      ]
    }
  };

  const result = submitLotEntry(fileId, stageResult, payload);
  Logger.log("DRY RUN — nothing written.");
  Logger.log(result.rendered);
}
