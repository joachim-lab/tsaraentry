/***************************************************************
 * TzAlign.js — 2026-09-15 (Kim). Step 2: set every Tsara
 * spreadsheet to Africa/Nairobi, the time zone of the scripts.
 *
 * PROVED FIRST (tzAuditTest, 2026-09-15): changing a spreadsheet's
 * time zone does NOT move any displayed date. Sheets stores a day
 * number. Only the instant Apps Script reads changes, to Nairobi
 * midnight — the goal. NOW() and TODAY() follow Nairobi after it.
 *
 * SCOPE. Every spreadsheet under My Drive/Tsara Tilapia (all
 * subfolders), except the "old lots" archive, plus the TSARA ENGINE
 * spreadsheet in My Drive root. Includes the lot template Lot-x,
 * so new lot files are born on Nairobi. Tama Ferme, Claude General
 * and other projects are not touched.
 *
 * RUN. TSARA Entry -> TzAlign.js -> tzAlignDryRun (writes nothing),
 * then tzAlignApply. Both stop after about 4.5 min and save their
 * place: run again until TERMINÉ. tzAlignReset starts over.
 ***************************************************************/

const TZL_TARGET = "Africa/Nairobi";
const TZL_ROOT = "1Aef_ggEw_0av7qosGbf3Abjusjv8cbnY";          // My Drive/Tsara Tilapia
const TZL_OLD_LOTS = "1ei75vejg3_CUtY4QdpQG9EmHHi1ujGYw";       // never touched
const TZL_EXTRA = ["1ZsAbW9fY-HYXoSW1pp55LeACcURYAH30VxfCTnQI_Z0"]; // TSARA ENGINE (root)
const TZL_KEY = "TZALIGN_STATE_20260915";
const TZL_BUDGET_MS = 270 * 1000;

function tzAlignReset() {
  PropertiesService.getScriptProperties().deleteProperty(TZL_KEY);
  Logger.log("État effacé.");
}

function tzAlignIds() {
  const ids = TZL_EXTRA.slice();
  const stack = [DriveApp.getFolderById(TZL_ROOT)];
  while (stack.length) {
    const f = stack.pop();
    if (f.getId() === TZL_OLD_LOTS) continue;
    const files = f.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) ids.push(files.next().getId());
    const subs = f.getFolders();
    while (subs.hasNext()) stack.push(subs.next());
  }
  return ids.filter(function (id, i) { return ids.indexOf(id) === i; }).sort();
}

function tzAlignDryRun() { tzAlignWork(false); }
function tzAlignApply() { tzAlignWork(true); }

function tzAlignWork(write) {
  const t0 = Date.now();
  const key = TZL_KEY + (write ? "_APPLY" : "_DRY");
  const props = PropertiesService.getScriptProperties();
  const st = JSON.parse(props.getProperty(key) || '{"last":"","seen":0,"off":0,"done":0,"errors":0}');
  const ids = tzAlignIds();
  const todo = ids.filter(function (id) { return id > st.last; });
  Logger.log((write ? "APPLY" : "DRY RUN") + " — racine : " + DriveApp.getFolderById(TZL_ROOT).getName() +
             " ; " + ids.length + " classeur(s) dans le périmètre ; " + todo.length + " restant(s).");

  for (var i = 0; i < todo.length; i++) {
    if (Date.now() - t0 > TZL_BUDGET_MS) {
      props.setProperty(key, JSON.stringify(st));
      Logger.log("PAS FINI : " + (todo.length - i) + " restant(s). Relancer.");
      return;
    }
    const id = todo[i];
    try {
      const ss = SpreadsheetApp.openById(id);
      const tz = ss.getSpreadsheetTimeZone();
      if (tz !== TZL_TARGET) {
        st.off++;
        if (write) {
          ss.setSpreadsheetTimeZone(TZL_TARGET);
          const now = SpreadsheetApp.openById(id).getSpreadsheetTimeZone();
          if (now === TZL_TARGET) st.done++;
          Logger.log((now === TZL_TARGET ? "CHANGÉ | " : "ÉCHEC | ") + tz + " -> " + now + " | " + ss.getName() + " | " + id);
        } else {
          Logger.log("À CHANGER | " + tz + " | " + ss.getName() + " | " + id);
        }
      }
    } catch (e) {
      st.errors++;
      Logger.log("ERREUR | " + id + " | " + e.message);
    }
    st.seen++;
    st.last = id;
  }
  props.deleteProperty(key);
  Logger.log("TERMINÉ : " + st.seen + " lu(s), " + st.off + " pas sur " + TZL_TARGET +
             (write ? ", " + st.done + " changé(s)" : "") + ", " + st.errors + " erreur(s).");
}
