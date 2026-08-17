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
 * SAMPLES ARE APPEND-ONLY: existing sample values are never cleared
 * or overwritten. Corrections to saved samples are made manually in
 * the sheet. See planSamples.
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

/**
 * Normalise a value for comparison only (never for writing).
 * Dates must be compared as yyyy-MM-dd: the sheet holds a Date object
 * while the browser sends the string "2026-08-09", and String(Date)
 * gives "Sat Aug 09 2026..." — so a naive compare reports a change
 * on every submission and would overwrite a live date formula with
 * static text.
 */
function normaliseForCompare(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

/** Add one intended change, recording the current value for the dry run. */
function addChange(sh, row, col, value, note, plan) {
  if (value === undefined) return;             // field not supplied -> leave alone
  const cell = sh.getRange(row, col);
  const from = cell.getValue();
  const to = (value === null) ? "" : value;
  if (normaliseForCompare(from) === normaliseForCompare(to)) return; // no-op

  const formula = cell.getFormula();
  if (formula) {
    throw new Error("Ecriture refusee : la cellule " + sh.getName() + "!" +
      cell.getA1Notation() + " contient une formule. Cette valeur est calculee " +
      "par le fichier et ne peut pas etre modifiee ici. Champ : " + (note || ""));
  }
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
    addChange(sh, 11, 4, f.D11, "Biomasse (g)", plan);
  }
  addChange(sh, 4, 2, f.B4, "Température de l'eau", plan);
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

  addChange(sh, 2, 2, f.B2, "Bassin (sous-lot 1)", plan);
  addChange(sh, 3, 2, f.B3, "Happa (sous-lot 1)", plan);
  addChange(sh, 4, 2, f.B4, "Température de l'eau", plan);

  // Sub-lots 2, 3 and 4. Four pairs exist on S3..S20 (B/C/D/E in rows 2
  // and 3), verified live on Lot-26 tab S5, 2026-08-17. Column F is not
  // green, so four is the limit.
  addChange(sh, 2, 3, f.C2, "Bassin (sous-lot 2)", plan);
  addChange(sh, 2, 4, f.D2, "Bassin (sous-lot 3)", plan);
  addChange(sh, 2, 5, f.E2, "Bassin (sous-lot 4)", plan);
  addChange(sh, 3, 3, f.C3, "Happa (sous-lot 2)", plan);
  addChange(sh, 3, 4, f.D3, "Happa (sous-lot 3)", plan);
  addChange(sh, 3, 5, f.E3, "Happa (sous-lot 4)", plan);

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
 * Samples (individual weights) — APPEND ONLY (Kim's decision,
 * 2026-08-09). New samples are written after the last non-empty cell
 * in the range; existing samples are NEVER cleared or overwritten.
 *
 * Consequence: the form must NOT pre-fill the sample list, or
 * submitting would append the existing values a second time. It shows
 * the existing count as read-only context instead. Corrections to
 * already-saved samples are done manually in the sheet.
 */
function planSamples(sh, tabName, samples, plan) {
  if (samples === undefined || !samples.length) return;
  const r = WRITE_CFG.SAMPLE_RANGES[tabName];
  if (!r) throw new Error("No sample range configured for tab: " + tabName);

  // Find the first free row: scan the whole range, remember the last
  // non-empty, start after it. (Not "first blank" — a stray gap in the
  // middle must not cause new samples to overwrite later entries.)
  const capacity = r.end - r.start + 1;
  const existing = sh.getRange(r.start, r.col, capacity, 1).getValues();
  let lastUsed = -1;
  for (let i = 0; i < capacity; i++) {
    const v = existing[i][0];
    if (v !== "" && v !== null) lastUsed = i;
  }
  const firstFree = lastUsed + 1;

  const remaining = capacity - firstFree;
  if (remaining <= 0) {
    throw new Error("L'échantillonnage de l'onglet " + tabName +
      " est déjà complet (" + capacity + "/" + capacity +
      "). Aucun échantillon supplémentaire n'est possible pour cette période.");
  }
  if (samples.length > remaining) {
    throw new Error("Plus assez de place: " + samples.length +
      " échantillon(s) à ajouter, mais seulement " + remaining +
      " emplacement(s) libre(s) sur " + tabName + ".");
  }

  for (let i = 0; i < samples.length; i++) {
    addChange(sh, r.start + firstFree + i, r.col, samples[i],
      "Échantillon (ajout " + (i + 1) + "/" + samples.length + ")", plan);
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

/**
 * RUN FROM EDITOR: dry-run test of the APPEND-ONLY sample path on
 * Lot-31, which is on tab 11-15 (an early tab with a J sample column).
 * Writes NOTHING. Confirms new samples land AFTER existing ones.
 */
function testWriteSamplesDryRun() {
  const fileId = '1DdQxI1mRzGbI_PHdrUMAjGk6l4ptyL31MaCo3il9X38'; // Lot-31
  const stageResult = getLotStage(fileId);
  Logger.log("Stage: " + JSON.stringify(stageResult));

  if (stageResult.stage !== "s-tab") {
    Logger.log("Expected an s-tab stage for this test; got: " + stageResult.stage);
    return;
  }

  // Show what's already there, so the append position can be checked.
  const current = getLotFieldValues(fileId, stageResult);
  const existing = current.fields.samples || { used: 0, capacity: 0, full: false };
  Logger.log("Échantillons: " + existing.used + "/" + existing.capacity +
    " (libres: " + existing.remaining + ")");

  if (existing.full) {
    Logger.log("Colonne d'échantillons complète (" + existing.used + "/" +
      existing.capacity + ") — l'échantillonnage de cette période est terminé. " +
      "Test d'ajout ignoré (comportement attendu, pas une erreur).");
    return;
  }

  const payload = { fields: { samples: [1.11, 2.22, 3.33] } };
  const result = submitLotEntry(fileId, stageResult, payload);
  Logger.log("DRY RUN — nothing written.");
  Logger.log(result.rendered);
}

/**
 * DIAGNOSTIC — report sample-column usage for every lot that is
 * currently on a sample-bearing tab (1-5 / 6-10 / 11-15 / 16-21 / S1).
 * Used to find a partially-filled column to test the append path
 * against. READ ONLY.
 */
function reportSampleCapacity() {
  const lots = getLotFileList();
  const sampleTabs = Object.keys(WRITE_CFG.SAMPLE_RANGES);
  let candidates = 0;

  lots.forEach(lot => {
    let stageResult;
    try {
      stageResult = getLotStage(lot.fileId);
    } catch (e) {
      Logger.log(lot.fileName + ": ERREUR " + e);
      return;
    }

    if (stageResult.stage !== "s-tab") return;
    if (sampleTabs.indexOf(stageResult.tabName) === -1) {
      Logger.log(lot.fileName + " (" + stageResult.tabName + "): pas de colonne d'échantillons");
      return;
    }

    const values = getLotFieldValues(lot.fileId, stageResult);
    const s = values.fields.samples;
    const flag = s.full ? "COMPLET" : ">>> PLACE LIBRE <<<";
    if (!s.full) candidates++;
    Logger.log(lot.fileName + " (" + stageResult.tabName + "): " +
      s.used + "/" + s.capacity + " — " + flag + "  [" + lot.fileId + "]");
  });

  Logger.log("");
  Logger.log("Lots avec de la place libre: " + candidates);
}

/**
 * RUN FROM EDITOR — PERFORMS A REAL WRITE, on the disposable test
 * copy ONLY (fileId hardcoded below, NOT a real lot). Do not repurpose
 * this function against a live fileId.
 *
 * Round-trip: read baseline -> dry run -> REAL write -> re-read ->
 * verify the 3 new samples landed in exactly the right cells and
 * nothing else moved.
 */
function testWriteSamplesLive_TESTCOPY_ONLY() {
  const TEST_FILE_ID = '1CvVR75ZeE6KqigEMMc-oFX6HtpYi8_YGFBhXlN-4lJQ'; // ZZTEST-Lot-31 copy

  const stageResult = getLotStage(TEST_FILE_ID);
  Logger.log("Stage: " + JSON.stringify(stageResult));
  if (stageResult.stage !== "s-tab" || stageResult.tabName !== "11-15") {
    Logger.log("ABORT: expected tab 11-15, got " + JSON.stringify(stageResult) +
      " — test copy may not be set up as expected.");
    return;
  }

  const before = getLotFieldValues(TEST_FILE_ID, stageResult).fields.samples;
  Logger.log("AVANT — échantillons: " + before.used + "/" + before.capacity +
    " (libres: " + before.remaining + ")");

  const testSamples = [99.1, 99.2, 99.3];
  const payload = { fields: { samples: testSamples } };

  // 1. Dry run
  const dry = submitLotEntry(TEST_FILE_ID, stageResult, payload);
  Logger.log("--- DRY RUN ---");
  Logger.log(dry.rendered);

  // 2. REAL write
  const live = submitLotEntry(TEST_FILE_ID, stageResult, payload, { dryRun: false });
  Logger.log("--- ECRITURE REELLE EFFECTUEE (" + live.changeCount + " cellule(s)) ---");

  // 3. Re-read and verify
  const stageAfter = getLotStage(TEST_FILE_ID);
  const after = getLotFieldValues(TEST_FILE_ID, stageAfter).fields.samples;
  Logger.log("APRES — échantillons: " + after.used + "/" + after.capacity +
    " (libres: " + after.remaining + ")");

  const expectUsed = before.used + testSamples.length;
  const ok1 = after.used === expectUsed;
  Logger.log((ok1 ? "OK" : "ECHEC") + " — used attendu " + expectUsed + ", obtenu " + after.used);

  // Verify exact cell placement: last 3 values in the array should be
  // our test values, in order, and nothing before them should have moved.
  const tail = after.values.slice(-3);
  const ok2 = JSON.stringify(tail) === JSON.stringify(testSamples);
  Logger.log((ok2 ? "OK" : "ECHEC") + " — 3 dernières valeurs = " + JSON.stringify(tail) +
    " (attendu " + JSON.stringify(testSamples) + ")");

  const untouched = after.values.slice(0, before.used);
  const beforeUntouched = before.values.slice(0, before.used);
  const ok3 = JSON.stringify(untouched) === JSON.stringify(beforeUntouched);
  Logger.log((ok3 ? "OK" : "ECHEC") + " — les " + before.used + " échantillons existants n'ont pas bougé");

  Logger.log(ok1 && ok2 && ok3 ? "=== TEST REUSSI ===" : "=== TEST ECHOUE — voir ci-dessus ===");
}

/** RUN FROM EDITOR: proves the formula guard. Writes NOTHING. */
function testFormulaGuard() {
  const FILE_ID = "10GhmZm67YmaTscdvIWRsYnybCPOA3gpkgXFxcKce_o8"; // Lot-31 live
  const sh = SpreadsheetApp.openById(FILE_ID).getSheetByName("16-21");
  const plan = [];
  let pass = 0, fail = 0;

  try {
    addChange(sh, 11, 4, 999, "test D11 formule", plan);
    Logger.log("FAIL 1 : pas d erreur sur D11 (cellule formule)");
    fail++;
  } catch (e) {
    Logger.log("PASS 1 : " + e.message);
    pass++;
  }

  try {
    addChange(sh, 11, 7, "guard test ok", "test G11 saisie", plan);
    Logger.log("PASS 2 : G11 accepte, plan = " + plan.length);
    pass++;
  } catch (e) {
    Logger.log("FAIL 2 : " + e.message);
    fail++;
  }

  Logger.log("pass=" + pass + " fail=" + fail + " ecritures=0");
}
