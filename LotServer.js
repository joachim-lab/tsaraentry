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
function uiPreviewSubmit(fileId, payload) {
  const stageResult = getLotStage(fileId);
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
  const result = submitLotEntry(fileId, stageResult, payload, { dryRun: false });
  return { changeCount: result.changeCount };
}
