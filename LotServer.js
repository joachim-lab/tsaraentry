/***************************************************************
 * LotServer.js — TSARA Entry web app
 * Screen 2: server endpoints the Lot UI calls.
 *
 * Thin layer over Lot.js / LotRead.js / LotWrite.js. Its jobs:
 *   - hand the UI everything it needs in ONE call (lot list, stage,
 *     current values) instead of several round trips
 *   - format dates as yyyy-MM-dd strings in the script's timezone,
 *     so the browser never does timezone maths (a Date crossing to
 *     the browser as UTC can render one day early — Africa/Nairobi
 *     is UTC+3)
 *   - classify a submission as "new" vs "overwrite" so the UI only
 *     asks for confirmation when existing data would change
 ***************************************************************/

/** Format a value as yyyy-MM-dd in the script timezone; pass through anything else. */
function formatDateForUI(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return v;
}

/** Recursively format any Date found in a plain object/array. */
function formatDatesDeep(obj) {
  if (obj instanceof Date) return formatDateForUI(obj);
  if (Array.isArray(obj)) return obj.map(formatDatesDeep);
  if (obj && typeof obj === "object") {
    const out = {};
    Object.keys(obj).forEach(k => { out[k] = formatDatesDeep(obj[k]); });
    return out;
  }
  return obj;
}

/** UI: the lot dropdown. */
function uiGetLotList() {
  return getLotFileList();
}

/**
 * UI: everything needed to render the form for one lot.
 * Returns { lotName, stage, tabName, targetGroupStartCol, kind,
 *           mixedWarning, fields } or { stage: "none", reason }.
 */
function uiLoadLot(fileId) {
  const stageResult = getLotStage(fileId);

  if (stageResult.stage === "none") {
    return { stage: "none", reason: stageResult.reason };
  }

  const values = getLotFieldValues(fileId, stageResult);

  return formatDatesDeep({
    fileId: fileId,
    stage: stageResult.stage,
    tabName: stageResult.tabName || null,
    targetGroupStartCol: stageResult.targetGroupStartCol || null,
    kind: values.kind || "grossissement",
    mixedWarning: stageResult.mixedWarning || null,
    fields: values.fields
  });
}

/**
 * UI: dry run. Returns the plan plus a classification the UI uses to
 * decide whether to ask for confirmation:
 *   overwrites = changes to cells that already held a value
 *   additions  = changes to cells that were empty
 * Only overwrites require confirmation (Kim's decision, 2026-08-09).
 */
/**
 * Refuse a save whose target moved while the form was open.
 *
 * The screen decides which tab (or which Grossissement block) it is on
 * when it LOADS. uiPreviewSubmit and uiSubmit both re-detect it at SAVE
 * time. If a form is left open across a date boundary the two differ,
 * and the payload built for one tab would be written into another
 * tab's cells — G11 on 16-21 is a comment, on S1 it is a formula.
 *
 * The client sends what it displayed as payload.stageStamp. Editor test
 * functions send no stamp, so they are unaffected.
 */
function assertStageUnchanged(stageResult, payload) {
  const stamp = payload && payload.stageStamp;
  if (!stamp) return;

  const nowTab = stageResult.tabName || null;
  const nowCol = stageResult.targetGroupStartCol || null;

  if (stamp.stage === stageResult.stage && stamp.tabName === nowTab &&
      stamp.targetGroupStartCol === nowCol) {
    return;
  }

  const was = stamp.tabName || ("Grossissement, colonne " + stamp.targetGroupStartCol);
  const is = nowTab || ("Grossissement, colonne " + nowCol);

  throw new Error(
    "Enregistrement refuse : la periode a change pendant que le formulaire " +
    "etait ouvert. Le formulaire a ete ouvert sur " + was + ", mais le lot " +
    "est maintenant sur " + is + ". Rien n'a ete ecrit. Rechargez la page et " +
    "ressaisissez.");
}

function uiPreviewSubmit(fileId, payload) {
  const stageResult = getLotStage(fileId);
  assertStageUnchanged(stageResult, payload);
  const result = submitLotEntry(fileId, stageResult, payload); // dry run by default

  const overwrites = [];
  const additions = [];
  result.plan.forEach(c => {
    const wasEmpty = (c.from === "" || c.from === null);
    const entry = {
      a1: c.sheet + "!" + c.a1,
      note: c.note,
      from: formatDateForUI(c.from),
      to: formatDateForUI(c.to)
    };
    if (wasEmpty) additions.push(entry); else overwrites.push(entry);
  });

  return {
    changeCount: result.changeCount,
    additions: additions,
    overwrites: overwrites,
    needsConfirmation: overwrites.length > 0
  };
}

/** UI: perform the write. Called only after the UI has handled any confirmation. */
function uiSubmit(fileId, payload) {
  const stageResult = getLotStage(fileId);
  assertStageUnchanged(stageResult, payload);
  const result = submitLotEntry(fileId, stageResult, payload, { dryRun: false });
  return { changeCount: result.changeCount };
}

/**
 * RUN FROM EDITOR: exercises the three UI endpoints without writing.
 * Uses Lot-24 (S7) — an S3+ tab with two live sub-lots.
 */
function testUiEndpoints() {
  const fileId = '1_qwbM0Hma6cVWJ3lr7jtVtcgId7hUC71kCFjYg_nCmE'; // Lot-24

  Logger.log("=== uiGetLotList (count only) ===");
  Logger.log(uiGetLotList().length + " lots");

  Logger.log("=== uiLoadLot ===");
  const loaded = uiLoadLot(fileId);
  Logger.log(JSON.stringify(loaded, null, 2));

  Logger.log("=== uiPreviewSubmit — mix of new + overwrite ===");
  // C11 already holds 0.45 (overwrite); I11 is empty (addition).
  const payload = {
    fields: {
      subLots: [
        { row: 11, poidsMoyen: 0.99, commentaire: "test preview" }
      ]
    }
  };
  Logger.log(JSON.stringify(uiPreviewSubmit(fileId, payload), null, 2));
}
