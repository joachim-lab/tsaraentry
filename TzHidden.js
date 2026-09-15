/***************************************************************
 * TzHidden.js — READ-ONLY, 2026-09-16 (Kim). Step 3 of the sheet
 * time zone alignment. Writes no cell.
 *
 * WHY. Before 2026-09-16 many Tsara sheets were not on Nairobi.
 * A date a script wrote into such a sheet kept a hidden time of
 * day (01:00 in a Dubai sheet, 03:00 in CohortRegistry, 14:00 in a
 * Los Angeles sheet). The time zone change cannot remove it.
 *
 * WHAT IT FINDS. A non-formula cell holding a date, whose number
 * format shows NO time (no h / s), but whose value is not
 * midnight. Grouped per file / tab / column, with the times seen.
 * A column of real timestamps shown as dates also appears here:
 * that is why nothing is fixed before the list is read.
 *
 * SCOPE. Same files as TzAlign.js (tzAlignIds).
 * RUN. tzHiddenRun until TERMINÉ. tzHiddenReset starts over.
 ***************************************************************/

const TZH_KEY = "TZHIDDEN_STATE_20260916";
const TZH_BUDGET_MS = 270 * 1000;

function tzHiddenReset() {
  PropertiesService.getScriptProperties().deleteProperty(TZH_KEY);
  Logger.log("État effacé.");
}

function tzHiddenFile(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const out = [];
  ss.getSheets().forEach(function (sh) {
    const rng = sh.getDataRange();
    const vals = rng.getValues(), fx = rng.getFormulas(), nf = rng.getNumberFormats();
    const byCol = {};
    for (var r = 0; r < vals.length; r++) {
      for (var c = 0; c < vals[r].length; c++) {
        const v = vals[r][c];
        if (!(v instanceof Date) || isNaN(v.getTime()) || fx[r][c]) continue;
        if (/[hHs]/.test(String(nf[r][c]))) continue;          // format shows a time: a real timestamp
        const hm = Utilities.formatDate(v, tz, "HH:mm");
        if (hm === "00:00") continue;
        const g = byCol[c] || (byCol[c] = { n: 0, first: r + 1, last: r + 1, times: {},
          ex: Utilities.formatDate(v, tz, "dd/MM/yyyy HH:mm") });
        g.n++; g.last = r + 1; g.times[hm] = (g.times[hm] || 0) + 1;
      }
    }
    Object.keys(byCol).forEach(function (c) {
      const g = byCol[c];
      const top = Object.keys(g.times).sort(function (a, b) { return g.times[b] - g.times[a]; })
        .slice(0, 4).map(function (k) { return k + "×" + g.times[k]; }).join(" ");
      out.push(sh.getName() + "!" + tzScanColLetterH(Number(c) + 1) + " | " + g.n + " cellule(s) | lignes " +
               g.first + "-" + g.last + " | heures " + top + (Object.keys(g.times).length > 4 ? " …" : "") +
               " | ex. " + g.ex);
    });
  });
  return out;
}

function tzScanColLetterH(n) {
  var s = "";
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function tzHiddenRun() {
  const t0 = Date.now();
  const props = PropertiesService.getScriptProperties();
  const st = JSON.parse(props.getProperty(TZH_KEY) || '{"last":"","seen":0,"dirty":0,"errors":0}');
  const ids = tzAlignIds();
  const todo = ids.filter(function (id) { return id > st.last; });
  Logger.log(ids.length + " classeur(s) ; " + todo.length + " restant(s).");
  for (var i = 0; i < todo.length; i++) {
    if (Date.now() - t0 > TZH_BUDGET_MS) {
      props.setProperty(TZH_KEY, JSON.stringify(st));
      Logger.log("PAS FINI : " + (todo.length - i) + " restant(s). Relancer tzHiddenRun.");
      return;
    }
    const id = todo[i];
    try {
      const ss = SpreadsheetApp.openById(id);
      const hits = tzHiddenFile(ss);
      if (hits.length) st.dirty++;
      hits.forEach(function (h) { Logger.log("HID | " + ss.getName() + " | " + id + " | " + h); });
    } catch (e) {
      st.errors++;
      Logger.log("ERREUR | " + id + " | " + e.message);
    }
    st.seen++;
    st.last = id;
  }
  props.deleteProperty(TZH_KEY);
  Logger.log("TERMINÉ : " + st.seen + " lu(s), " + st.dirty + " avec heures cachées, " + st.errors + " erreur(s).");
}
