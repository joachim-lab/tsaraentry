/***************************************************************
 * TzFix2.js — one-time repair, 2026-09-15 (Kim). Follows TzScan.js.
 *
 * TARGET. Nourrissage & inventaire -> "snapshots stock" -> column E
 * ("Date comptage"). WARNINGSYSTEM wrote countDay = new Date(y, m, d)
 * under Asia/Dubai: 20:00 UTC = 23:00 the day before in this
 * Africa/Nairobi sheet. TzScan found 14 cells, rows 58-71.
 * The fix adds one hour -> 00:00 of the right day.
 *
 * NOT HERE, on purpose: "FCR reel (par cohorte)" and Production
 * Model V3 "Fry needs". Both tabs are cleared and rebuilt by their
 * own script, so a rebuild fixes them.
 *
 * RUN. TSARA Entry -> TzFix2.js -> tzFix2DryRun, then tzFix2Apply.
 ***************************************************************/

const TZF2_SS_ID = "1JBoH5c7BqZc2V5czcDAnEt-2hvkNKlJAxuDSupttTfs";
const TZF2_TAB = "snapshots stock";
const TZF2_COL = 5;                               // E
const TZF2_BAD_MS_OF_DAY = 20 * 3600 * 1000;      // 20:00 UTC = Dubai midnight

function tzFix2Scan() {
  const ss = SpreadsheetApp.openById(TZF2_SS_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const sh = ss.getSheetByName(TZF2_TAB);
  if (!sh) throw new Error('Onglet introuvable: "' + TZF2_TAB + '"');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rng = sh.getRange(2, TZF2_COL, last - 1, 1);
  const vals = rng.getValues();
  const fx = rng.getFormulas();
  const hits = [];
  for (var i = 0; i < vals.length; i++) {
    const v = vals[i][0];
    if (!(v instanceof Date) || isNaN(v.getTime()) || fx[i][0]) continue;
    const ms = ((v.getTime() % 86400000) + 86400000) % 86400000;
    if (ms !== TZF2_BAD_MS_OF_DAY) continue;
    const fixed = new Date(v.getTime() + 3600 * 1000);
    hits.push({ sh: sh, row: 2 + i, fixed: fixed,
      before: Utilities.formatDate(v, tz, "dd/MM/yyyy HH:mm"),
      after: Utilities.formatDate(fixed, tz, "dd/MM/yyyy HH:mm") });
  }
  return hits;
}

function tzFix2Log(hits, verb) {
  Logger.log(hits.length + " cellule(s) " + verb + ".");
  hits.forEach(function (h) { Logger.log("E" + h.row + " : " + h.before + " -> " + h.after); });
}

function tzFix2DryRun() { tzFix2Log(tzFix2Scan(), "à corriger (DRY RUN, rien écrit)"); }

function tzFix2Apply() {
  const hits = tzFix2Scan();
  hits.forEach(function (h) { h.sh.getRange(h.row, TZF2_COL).setValue(h.fixed); });
  SpreadsheetApp.flush();
  tzFix2Log(hits, "corrigée(s)");
  Logger.log("Contrôle après écriture : " + tzFix2Scan().length + " cellule(s) restante(s).");
}
