/***************************************************************
 * LotRead.js — TSARA Entry web app
 * Screen 2: Lot (échantillonnage) — read current field values.
 *
 * Given getLotStage()'s result, reads whatever is currently in the
 * editable cells for that stage — so the form shows existing entries
 * (e.g. a partial entry from earlier today) instead of blanking them.
 * READ ONLY. No writes here.
 *
 * Field ranges are Kim's confirmed live list (2026-08-09), not
 * rediscovered from formulas — verified against live lot files
 * (Lot-14, Lot-Mirana, Lot-21) but taken as given, not inferred.
 ***************************************************************/

const READ_CFG = {
  SUBLOT_FIRST_ROW: 11,
  SUBLOT_LAST_ROW: 16   // rows 11-16 = up to 6 sub-lots, S2-Tri onward only
};

/** Main entry point: fileId + the result of getLotStage() -> current field values. */
function getLotFieldValues(fileId, stageResult) {
  const ss = SpreadsheetApp.openById(fileId);

  if (stageResult.stage === "grossissement") {
    return readGrossissementFields_(ss, stageResult.targetGroupStartCol);
  }
  if (stageResult.stage === "s-tab") {
    return readSTabFields_(ss, stageResult.tabName);
  }
  throw new Error("Cannot read fields for stage: " + stageResult.stage);
}

/** Dispatches to the right reader based on which tab this is. */
function readSTabFields_(ss, tabName) {
  const sh = ss.getSheetByName(tabName);
  if (!sh) throw new Error('Tab not found: "' + tabName + '"');

  if (tabName === "S2-Tri") return readS2Tri_(sh, tabName);
  if (tabName.match(/^S\d+$/) && tabName !== "S1") return readS3Plus_(sh, tabName);
  if (tabName === "S1") return readS1_(sh, tabName);
  return readEarlyTab_(sh, tabName); // 1-5, 6-10, 11-15, 16-21
}

/**
 * 1-5 / 6-10 / 11-15 / 16-21 — always exactly one sub-lot, row 11 only
 * (confirmed: no sub-lot split is possible before S2-Tri). Sample
 * column and row extent per Kim's exact live-confirmed ranges.
 */
function readEarlyTab_(sh, tabName) {
  const ranges = {
    "1-5":   { sampleCol: "J", sampleStart: 5, sampleEnd: 57, hasB1B5: true },
    "6-10":  { sampleCol: "J", sampleStart: 6, sampleEnd: 55, hasB1B5: false },
    "11-15": { sampleCol: "J", sampleStart: 6, sampleEnd: 55, hasB1B5: false },
    "16-21": { sampleCol: "J", sampleStart: 6, sampleEnd: 55, hasB1B5: false }
  };
  const r = ranges[tabName];

  // B11 Nombre, C11 Poids moyen, D11 Biomasse — read in ONE call.
  // B11 and C11 are sheet-owned on every early tab (B11 is a formula
  // everywhere, C11 is the sample average everywhere), so they are read
  // for DISPLAY only and are never written back. See the field-ownership
  // contract, 2026-08-16 §3.4.
  const row11 = sh.getRange("B11:D11").getValues()[0];

  const fields = {
    B4: sh.getRange("B4").getValue(),
    B11: row11[0],
    C11: row11[1],
    D11: row11[2],
    G11: sh.getRange("G11").getValue(),
    samples: readSampleColumn_(sh, r.sampleCol, r.sampleStart, r.sampleEnd)
  };
  if (r.hasB1B5) {
    fields.B1 = sh.getRange("B1").getValue();
    fields.B2 = sh.getRange("B2").getValue();
    fields.B3 = sh.getRange("B3").getValue();
    fields.B5 = sh.getRange("B5").getValue();
  }

  return { stage: "s-tab", tabName: tabName, kind: "early", fields: fields };
}

/** S1 — one sub-lot, row 11 only. Samples in column K. */
function readS1_(sh, tabName) {
  // B11 Nombre (='1-5'!B11), C11 Poids moyen (=K55), D11 Biomasse.
  // All three are sheet-owned on S1 — display only, never written.
  const row11 = sh.getRange("B11:D11").getValues()[0];

  const fields = {
    B4: sh.getRange("B4").getValue(),
    B11: row11[0],
    C11: row11[1],
    D11: row11[2],
    H11: sh.getRange("H11").getValue(),
    samples: readSampleColumn_(sh, "K", 5, 54)
  };
  return { stage: "s-tab", tabName: tabName, kind: "s1", fields: fields };
}

/**
 * S2-Tri — the sort/split event. B2:B4 = original bassin/happa/temp,
 * B7 = date de tri, C2:D3 = bassin/happa dropdowns for new sub-lots
 * (per Kim, 2026-08-09), C11:D16 = Poids moyen/Biomasse per sub-lot
 * row, H11:H16 = comments per row. B11:B16 (Nombre) is derived
 * (=IFERROR(D/C,"")) — read for display, never written.
 */
function readS2Tri_(sh, tabName) {
  const fields = {
    B2: sh.getRange("B2").getValue(),
    B3: sh.getRange("B3").getValue(),
    B4: sh.getRange("B4").getValue(),
    B7: sh.getRange("B7").getValue(),
    C2: sh.getRange("C2").getValue(),
    D2: sh.getRange("D2").getValue(),
    C3: sh.getRange("C3").getValue(),
    D3: sh.getRange("D3").getValue(),
    subLots: []
  };

  const n = READ_CFG.SUBLOT_LAST_ROW - READ_CFG.SUBLOT_FIRST_ROW + 1;
  const labels = sh.getRange(READ_CFG.SUBLOT_FIRST_ROW, 1, n, 1).getDisplayValues();
  const bVals  = sh.getRange(READ_CFG.SUBLOT_FIRST_ROW, 2, n, 1).getValues(); // Nombre, derived
  const cVals  = sh.getRange(READ_CFG.SUBLOT_FIRST_ROW, 3, n, 1).getValues(); // Poids moyen
  const dVals  = sh.getRange(READ_CFG.SUBLOT_FIRST_ROW, 4, n, 1).getValues(); // Biomasse
  const hVals  = sh.getRange(READ_CFG.SUBLOT_FIRST_ROW, 8, n, 1).getValues(); // Commentaires

  for (let i = 0; i < n; i++) {
    fields.subLots.push({
      row: READ_CFG.SUBLOT_FIRST_ROW + i,
      label: labels[i][0],
      nombre: bVals[i][0],   // derived, display only
      poidsMoyen: cVals[i][0],
      biomasse: dVals[i][0],
      commentaire: hVals[i][0]
    });
  }
  return { stage: "s-tab", tabName: tabName, kind: "s2-tri", fields: fields };
}

/**
 * S3..S20 — B2:B4 shared (bassin/happa/temp), then per sub-lot row
 * (11-16): B=Nombre (typed), C=Poids moyen (typed), I=Commentaires.
 * D (Biomasse) is derived (=C*B), read for display, never written.
 */
function readS3Plus_(sh, tabName) {
  // Bassin/Happa for FOUR sub-lots: B2/B3, C2/C3, D2/D3, E2/E3.
  // Read live from Lot-26 tab S5 on 2026-08-17: all eight cells carry
  // the green fill and a dropdown; column F is empty and not green, so
  // four is the template limit. The green-cell contract of 2026-08-16 §5
  // listed only three pairs — it was incomplete.
  // The screen used to show only B2/B3, so sub-lots 2, 3 and 4 could not
  // be recorded at all.
  const row2 = sh.getRange("B2:E2").getValues()[0];
  const row3 = sh.getRange("B3:E3").getValues()[0];

  const fields = {
    B4: sh.getRange("B4").getValue(),
    B2: row2[0],
    C2: row2[1],
    D2: row2[2],
    E2: row2[3],
    B3: row3[0],
    C3: row3[1],
    D3: row3[2],
    E3: row3[3],
    subLots: []
  };

  const n = READ_CFG.SUBLOT_LAST_ROW - READ_CFG.SUBLOT_FIRST_ROW + 1;
  const labels = sh.getRange(READ_CFG.SUBLOT_FIRST_ROW, 1, n, 1).getDisplayValues();
  const bVals  = sh.getRange(READ_CFG.SUBLOT_FIRST_ROW, 2, n, 1).getValues(); // Nombre
  const cVals  = sh.getRange(READ_CFG.SUBLOT_FIRST_ROW, 3, n, 1).getValues(); // Poids moyen
  const dVals  = sh.getRange(READ_CFG.SUBLOT_FIRST_ROW, 4, n, 1).getValues(); // Biomasse, derived
  const iVals  = sh.getRange(READ_CFG.SUBLOT_FIRST_ROW, 9, n, 1).getValues(); // Commentaires

  for (let i = 0; i < n; i++) {
    fields.subLots.push({
      row: READ_CFG.SUBLOT_FIRST_ROW + i,
      label: labels[i][0],
      nombre: bVals[i][0],
      poidsMoyen: cVals[i][0],
      biomasse: dVals[i][0],   // derived, display only
      commentaire: iVals[i][0]
    });
  }
  return { stage: "s-tab", tabName: tabName, kind: "s3plus", fields: fields };
}

/**
 * Grossissement — reads the target block only (targetGroupStartCol from
 * getLotStage: the next empty block, or a partially-filled one if
 * staff already started today). date/temp are the block's shared
 * first column; bassin/happa/nombre/PM are per column (up to 4
 * sub-lots per block). Biomasse (row 12) is derived, display only.
 */
function readGrossissementFields_(ss, targetGroupStartCol) {
  const gross = ss.getSheetByName(LOT_CFG.GROSS_SHEET);
  const g = LOT_CFG.GROSS_GROUP_SIZE;

  const date = gross.getRange(LOT_CFG.GROSS_ROW_DATE, targetGroupStartCol).getValue();
  const temp = gross.getRange(LOT_CFG.GROSS_ROW_TEMP, targetGroupStartCol).getValue();

  const bassinRow  = gross.getRange(LOT_CFG.GROSS_ROW_BASSIN, targetGroupStartCol, 1, g).getValues()[0];
  const happaRow   = gross.getRange(LOT_CFG.GROSS_ROW_HAPPA, targetGroupStartCol, 1, g).getValues()[0];
  const nombreRow  = gross.getRange(LOT_CFG.GROSS_ROW_NOMBRE, targetGroupStartCol, 1, g).getValues()[0];
  const pmRow      = gross.getRange(LOT_CFG.GROSS_ROW_PM, targetGroupStartCol, 1, g).getValues()[0];

  const subLots = [];
  for (let i = 0; i < g; i++) {
    subLots.push({
      col: targetGroupStartCol + i,
      bassin: bassinRow[i],
      happa: happaRow[i],
      nombre: nombreRow[i],
      poidsMoyen: pmRow[i]
    });
  }

  return {
    stage: "grossissement",
    targetGroupStartCol: targetGroupStartCol,
    fields: { date: date, temperature: temp, subLots: subLots }
  };
}

/**
 * Reads a sample column (individual weights) between two rows and
 * reports its capacity.
 *
 * Samples are APPEND-ONLY (see LotWrite.js), so the UI needs to know
 * at load time whether there is room left — a full column means the
 * échantillonnage for this period is complete and the entry box
 * should be hidden, rather than failing at submit time.
 *
 * `used` counts up to the LAST non-empty cell (not the number of
 * non-blank values), because that is where appending actually starts;
 * a stray gap mid-column must not be reported as free space.
 *
 * Returns { values, used, capacity, remaining, full }.
 */
function readSampleColumn_(sh, colLetter, startRow, endRow) {
  const capacity = endRow - startRow + 1;
  const col = sh.getRange(colLetter + startRow + ":" + colLetter + endRow).getColumn();
  const raw = sh.getRange(startRow, col, capacity, 1).getValues();

  let lastUsed = -1;
  for (let i = 0; i < capacity; i++) {
    const v = raw[i][0];
    if (v !== "" && v !== null) lastUsed = i;
  }
  const used = lastUsed + 1;

  return {
    values: raw.map(r => r[0]).filter(v => v !== "" && v !== null),
    used: used,
    capacity: capacity,
    remaining: capacity - used,
    full: used >= capacity
  };
}

/** RUN FROM EDITOR: manual test — reads fields for one lot at whatever stage it's at. */
function testReadFields(fileId) {
  const stageResult = getLotStage(fileId);
  Logger.log("Stage: " + JSON.stringify(stageResult, null, 2));
  const values = getLotFieldValues(fileId, stageResult);
  Logger.log("Values: " + JSON.stringify(values, null, 2));
}

/** RUN FROM EDITOR: convenience wrapper — tests one Grossissement lot and one S-tab lot. */
function runReadTest() {
  testReadFields('1A0Cz7M3m3cRdadfcmwUGam-IPdTLn8hPFwMmbXU0DcQ'); // Lot-21, grossissement
  testReadFields('1_qwbM0Hma6cVWJ3lr7jtVtcgId7hUC71kCFjYg_nCmE');  // Lot-24, s-tab
}
