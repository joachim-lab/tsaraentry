/***************************************************************
 * TzRepair.js — one-time repair, 2026-09-15 (Kim).
 *
 * WHY. From 2026-09-09 to 2026-09-15 the script time zone was
 * Asia/Dubai (UTC+4) while the Commandes file is Africa/Nairobi
 * (UTC+3). cmdParseDate / recParseDate build new Date(y, m, d) =
 * midnight DUBAI = 20:00 UTC. In the sheet that shows as 23:00 on
 * the day BEFORE. Every date saved by the app in that window sits
 * one day early, with a hidden 23:00.
 *
 * SIGNATURE. A Date whose instant is exactly 20:00:00.000 UTC.
 * A date typed in the sheet, or saved before 09-09 or after the fix,
 * is midnight Nairobi = 21:00 UTC. So the signature matches only
 * the damaged cells. The repair adds one hour: 23:00 -> 00:00 of the
 * next day, the date that was typed.
 *
 * LIMIT. A date re-saved in that window lost one day PER SAVE
 * (15 -> 14 -> 13). The repair restores only one day. Check every
 * listed row against the paper order.
 *
 * RUN. TSARA Entry -> TzRepair.js -> tzRepairDryRun first. Read the
 * log. Then tzRepairApply. Delete this file after use (editor UI:
 * clasp push never deletes a remote file).
 ***************************************************************/

const TZR_BAD_MS_OF_DAY = 20 * 3600 * 1000;   // 20:00 UTC = Dubai midnight
const TZR_HOUR_MS = 3600 * 1000;

/** Tabs and columns that hold app-written dates. */
function tzRepairTargets() {
  const C = CMD_CFG.COL;
  return [
    { tab: CMD_CFG.SHEET, firstRow: CMD_CFG.START_ROW,
      cols: [C.DATE_CMD, C.PAIEMENT, C.DATE_LIVRAISON, C.RECU] },
    { tab: DEM_SHEET, firstRow: DEM_START, cols: [1] },
    { tab: REC_SHEET, firstRow: 2, cols: [5, 6, 10] }
  ];
}

function tzRepairScan() {
  const ss = SpreadsheetApp.openById(CMD_CFG.SS_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const hits = [];
  tzRepairTargets().forEach(function (t) {
    const sh = ss.getSheetByName(t.tab);
    if (!sh) throw new Error('Onglet introuvable: "' + t.tab + '"');
    const last = sh.getLastRow();
    if (last < t.firstRow) return;
    const n = last - t.firstRow + 1;
    const orderNo = (t.tab === CMD_CFG.SHEET)
      ? sh.getRange(t.firstRow, CMD_CFG.COL.ORDER_NO, n, 1).getValues() : null;
    t.cols.forEach(function (col) {
      const rng = sh.getRange(t.firstRow, col, n, 1);
      const vals = rng.getValues();
      const fx = rng.getFormulas();
      for (var i = 0; i < n; i++) {
        const v = vals[i][0];
        if (!(v instanceof Date) || isNaN(v.getTime())) continue;
        if (fx[i][0]) continue;
        const ms = ((v.getTime() % 86400000) + 86400000) % 86400000;
        if (ms !== TZR_BAD_MS_OF_DAY) continue;
        const fixed = new Date(v.getTime() + TZR_HOUR_MS);
        hits.push({
          sh: sh, row: t.firstRow + i, col: col, fixed: fixed,
          a1: t.tab + "!" + sh.getRange(t.firstRow + i, col).getA1Notation(),
          ref: orderNo ? String(orderNo[i][0] || "") : "",
          before: Utilities.formatDate(v, tz, "dd/MM/yyyy HH:mm"),
          after: Utilities.formatDate(fixed, tz, "dd/MM/yyyy HH:mm")
        });
      }
    });
  });
  return hits;
}

function tzRepairLog(hits, verb) {
  Logger.log(hits.length + " cellule(s) " + verb + ".");
  hits.forEach(function (h) {
    Logger.log(h.a1 + (h.ref ? " " + h.ref : "") + " : " + h.before + " -> " + h.after);
  });
}

/** Lists the damaged cells. Writes nothing. */
function tzRepairDryRun() {
  tzRepairLog(tzRepairScan(), "à corriger (DRY RUN, rien écrit)");
}

/** Rescans, then writes the corrected dates. Safe to run twice. */
function tzRepairApply() {
  const hits = tzRepairScan();
  hits.forEach(function (h) { h.sh.getRange(h.row, h.col).setValue(h.fixed); });
  SpreadsheetApp.flush();
  tzRepairLog(hits, "corrigée(s)");
  const left = tzRepairScan().length;
  Logger.log("Contrôle après écriture : " + left + " cellule(s) restante(s).");
}
