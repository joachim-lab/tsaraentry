/***************************************************************
 * TzScan.js — READ-ONLY scan, 2026-09-15 (Kim). Writes no cell.
 *
 * WHY. From 2026-09-09 to 2026-09-15 all 18 projects ran on
 * Asia/Dubai (UTC+4). The sheets are Africa/Nairobi (UTC+3). A
 * date built as midnight by any script (new Date(y, m, d),
 * setHours(0,0,0,0)) was stored as 20:00:00.000 UTC = 23:00 the
 * day before in the sheet. TzRepair.js fixed the Commandes file.
 * This scan finds the same signature everywhere else.
 *
 * SCOPE. Every spreadsheet in Drive modified since 2026-09-09.
 * Damage cannot sit in a file not modified since then. Formula
 * cells are skipped: they follow their inputs.
 *
 * RUN. TSARA Entry -> TzScan.js -> tzScanRun. It stops after
 * about 4.5 minutes and saves its place. Run it again until the
 * log says TERMINÉ. tzScanReset starts over. Delete this file
 * after use (editor UI).
 ***************************************************************/

const TZS_KEY = "TZSCAN_STATE_20260915";
const TZS_BAD_MS_OF_DAY = 20 * 3600 * 1000;   // 20:00 UTC = Dubai midnight
const TZS_BUDGET_MS = 270 * 1000;
const TZS_QUERY = "mimeType = 'application/vnd.google-apps.spreadsheet'" +
                  " and modifiedDate > '2026-09-09T00:00:00' and trashed = false";

function tzScanReset() {
  PropertiesService.getScriptProperties().deleteProperty(TZS_KEY);
  Logger.log("État effacé. Le prochain tzScanRun repart du début.");
}

function tzScanColLetter(n) {
  var s = "";
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** Hits in one spreadsheet: [{tab, col, count, first, last, sample}]. */
function tzScanFile(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const out = [];
  ss.getSheets().forEach(function (sh) {
    const rng = sh.getDataRange();
    const vals = rng.getValues();
    const fx = rng.getFormulas();
    const byCol = {};
    for (var r = 0; r < vals.length; r++) {
      for (var c = 0; c < vals[r].length; c++) {
        const v = vals[r][c];
        if (!(v instanceof Date) || isNaN(v.getTime()) || fx[r][c]) continue;
        const ms = ((v.getTime() % 86400000) + 86400000) % 86400000;
        if (ms !== TZS_BAD_MS_OF_DAY) continue;
        const g = byCol[c] || (byCol[c] = { count: 0, first: r + 1, last: r + 1,
          sample: Utilities.formatDate(v, tz, "dd/MM/yyyy HH:mm") });
        g.count++; g.last = r + 1;
      }
    }
    Object.keys(byCol).forEach(function (c) {
      const g = byCol[c];
      out.push({ tab: sh.getName(), col: tzScanColLetter(Number(c) + 1),
                 count: g.count, first: g.first, last: g.last, sample: g.sample });
    });
  });
  return out;
}

function tzScanRun() {
  const t0 = Date.now();
  const props = PropertiesService.getScriptProperties();
  const st = JSON.parse(props.getProperty(TZS_KEY) || '{"last":"","files":0,"hits":0,"dirty":0,"errors":0}');

  const ids = [];
  const it = DriveApp.searchFiles(TZS_QUERY);
  while (it.hasNext()) ids.push(it.next().getId());
  ids.sort();
  const todo = ids.filter(function (id) { return id > st.last; });
  Logger.log(ids.length + " classeur(s) modifié(s) depuis le 09/09 ; " +
             todo.length + " restant(s) à scanner.");

  for (var i = 0; i < todo.length; i++) {
    if (Date.now() - t0 > TZS_BUDGET_MS) {
      props.setProperty(TZS_KEY, JSON.stringify(st));
      Logger.log("PAS FINI : " + st.files + " scanné(s), " + (todo.length - i) +
                 " restant(s). Relancer tzScanRun.");
      return;
    }
    const id = todo[i];
    try {
      const ss = SpreadsheetApp.openById(id);
      const hits = tzScanFile(ss);
      if (hits.length) {
        st.dirty++;
        hits.forEach(function (h) {
          st.hits += h.count;
          Logger.log("HIT | " + ss.getName() + " | " + id + " | " + h.tab + "!" + h.col +
                     " | " + h.count + " cellule(s) | lignes " + h.first + "-" + h.last +
                     " | ex. " + h.sample);
        });
      }
    } catch (e) {
      st.errors++;
      Logger.log("ERREUR | " + id + " | " + e.message);
    }
    st.files++;
    st.last = id;
  }
  props.deleteProperty(TZS_KEY);
  Logger.log("TERMINÉ : " + st.files + " classeur(s) scanné(s), " + st.dirty +
             " avec dégâts, " + st.hits + " cellule(s), " + st.errors + " erreur(s).");
}
