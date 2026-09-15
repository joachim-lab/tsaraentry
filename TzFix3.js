/***************************************************************
 * TzFix3.js — one-time repair, 2026-09-16 (Kim). Step 4 of the
 * sheet time zone alignment.
 *
 * TARGET. CohortRegistry -> Movements -> column D (EventDate).
 * Until 2026-09-16 this sheet was on UTC+7. Scripts wrote midnight
 * dates into it: from Nairobi (UTC+3) they landed as 04:00, from
 * Dubai (UTC+4, 09-09..09-15) as 03:00. Before the change, scripts
 * read 04:00 back as exact Nairobi midnight. Now they read 04:00.
 * cr_resolveCohortForLotId_ drops a movement whose date is later
 * than asOf: a same-day movement at 04:00 is now dropped. Setting
 * those cells back to 00:00 restores the old reading; the 03:00
 * cells also get their correct day back in the scripts.
 *
 * ONLY cells at exactly 04:00:00 or 03:00:00 (Nairobi). Other
 * times are "now" timestamps written by recordMovement; their
 * behaviour did not change and they are left as they are.
 *
 * RUN. TSARA Entry -> TzFix3.js -> tzFix3DryRun, then tzFix3Apply.
 ***************************************************************/

const TZF3_SS_ID = "10nwW_3pJ9ineIHKL1nY3AXBdb8cCcF2dGGsENQ7qQTE";
const TZF3_TAB = "Movements";
const TZF3_COL = 4;                                   // D = EventDate
const TZF3_HOURS = [3, 4];
const TZF3_OFFSET_MS = 3 * 3600 * 1000;               // Africa/Nairobi, no DST

function tzFix3Plan() {
  const ss = SpreadsheetApp.openById(TZF3_SS_ID);
  const tz = ss.getSpreadsheetTimeZone();
  if (tz !== "Africa/Nairobi") throw new Error("CohortRegistry n'est pas sur Africa/Nairobi (" + tz + "). Arrêt.");
  const sh = ss.getSheetByName(TZF3_TAB);
  if (!sh) throw new Error('Onglet introuvable: "' + TZF3_TAB + '"');
  const hdr = String(sh.getRange(1, TZF3_COL).getValue());
  if (hdr !== "EventDate") throw new Error("D1 = " + hdr + ", attendu EventDate. Arrêt.");
  const last = sh.getLastRow();
  const rng = sh.getRange(2, TZF3_COL, last - 1, 1);
  if (rng.getFormulas().some(function (r) { return r[0]; })) throw new Error("Formule dans la colonne D. Arrêt.");
  const vals = rng.getValues();
  const byHour = {}, sample = {};
  var n = 0;
  const out = vals.map(function (r, i) {
    const v = r[0];
    if (!(v instanceof Date) || isNaN(v.getTime())) return [v];
    const msDay = (((v.getTime() + TZF3_OFFSET_MS) % 86400000) + 86400000) % 86400000;
    const h = msDay / 3600000;
    if (TZF3_HOURS.indexOf(h) < 0) return [v];
    n++; byHour[h] = (byHour[h] || 0) + 1;
    const fixed = new Date(v.getTime() - msDay);
    if (!sample[h]) sample[h] = "D" + (i + 2) + " " + Utilities.formatDate(v, tz, "dd/MM/yyyy HH:mm") +
                               " -> " + Utilities.formatDate(fixed, tz, "dd/MM/yyyy HH:mm");
    return [fixed];
  });
  return { rng: rng, out: out, n: n, byHour: byHour, sample: sample, rows: last - 1 };
}

function tzFix3Log(p, verb) {
  Logger.log(p.n + " cellule(s) " + verb + " sur " + p.rows + " ligne(s). Par heure : " + JSON.stringify(p.byHour));
  Object.keys(p.sample).forEach(function (h) { Logger.log("ex. " + p.sample[h]); });
}

function tzFix3DryRun() { tzFix3Log(tzFix3Plan(), "à corriger (DRY RUN, rien écrit)"); }

function tzFix3Apply() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const p = tzFix3Plan();
    if (p.n) p.rng.setValues(p.out);
    SpreadsheetApp.flush();
    tzFix3Log(p, "corrigée(s)");
    Logger.log("Contrôle après écriture : " + tzFix3Plan().n + " cellule(s) restante(s).");
  } finally {
    lock.releaseLock();
  }
}
