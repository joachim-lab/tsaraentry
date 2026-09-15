/***************************************************************
 * TzAudit.js — 2026-09-15 (Kim). Step 1 of aligning the sheet
 * time zones to Africa/Nairobi. Changes no existing file.
 *
 * tzAuditTest — proves what a time zone change does to a stored
 *   date, on a TEMPORARY spreadsheet made in Tsara Tilapia/Claude/.
 *   Expected: the displayed date stays the same (Sheets stores a
 *   day number, not an instant); only the instant Apps Script
 *   reads moves by the offset difference. The temp file is sent
 *   to the Drive trash at the end.
 *
 * tzAuditRun — lists every spreadsheet you own whose time zone is
 *   not Africa/Nairobi. Skips the "old lots" archive. Stops after
 *   about 4.5 minutes and saves its place: run again until TERMINÉ.
 *   tzAuditReset starts over.
 ***************************************************************/

const TZA_TARGET = "Africa/Nairobi";
const TZA_CLAUDE_FOLDER = "1lhbm6Mzqj2X9ViBSvgY8cAksK216uZaZ";   // Tsara Tilapia/Claude/
const TZA_OLD_LOTS = "1ei75vejg3_CUtY4QdpQG9EmHHi1ujGYw";        // never touched
const TZA_KEY = "TZAUDIT_STATE_20260915";
const TZA_BUDGET_MS = 270 * 1000;

function tzAuditTest() {
  const ss = SpreadsheetApp.create("TZ TEST temp (safe to trash)");
  const file = DriveApp.getFileById(ss.getId());
  file.moveTo(DriveApp.getFolderById(TZA_CLAUDE_FOLDER));
  try {
    ss.setSpreadsheetTimeZone("Indian/Reunion");                  // UTC+4, like the lot files
    const serial = (Date.UTC(2026, 8, 10) - Date.UTC(1899, 11, 30)) / 86400000;
    const cell = ss.getSheets()[0].getRange("A1");
    cell.setValue(serial).setNumberFormat("dd/MM/yyyy HH:mm");    // typed-style date: 10/09/2026 00:00
    SpreadsheetApp.flush();
    const b = SpreadsheetApp.openById(ss.getId()).getSheets()[0].getRange("A1");
    const before = { disp: b.getDisplayValue(), iso: b.getValue().toISOString() };

    ss.setSpreadsheetTimeZone(TZA_TARGET);
    SpreadsheetApp.flush();
    const ss2 = SpreadsheetApp.openById(ss.getId());
    const a = ss2.getSheets()[0].getRange("A1");
    const after = { tz: ss2.getSpreadsheetTimeZone(), disp: a.getDisplayValue(), iso: a.getValue().toISOString() };

    Logger.log("AVANT (Indian/Reunion) : affiché " + before.disp + " | instant " + before.iso);
    Logger.log("APRÈS (" + after.tz + ") : affiché " + after.disp + " | instant " + after.iso);
    const ok = before.disp === after.disp && before.disp === "10/09/2026 00:00" &&
               after.iso === "2026-09-09T21:00:00.000Z";
    Logger.log(ok ? "TEST OK : l'affichage ne bouge pas, seul l'instant lu passe à minuit Nairobi."
                  : "TEST INATTENDU : ne rien changer, prévenir Claude.");
  } finally {
    file.setTrashed(true);
    Logger.log("Fichier test mis à la corbeille.");
  }
}

function tzAuditReset() {
  PropertiesService.getScriptProperties().deleteProperty(TZA_KEY);
  Logger.log("État effacé.");
}

function tzAuditPath(file) {
  const names = [];
  var it = file.getParents();
  var guard = 0;
  while (it.hasNext() && guard++ < 8) {
    const f = it.next();
    if (f.getId() === TZA_OLD_LOTS) return null;
    names.unshift(f.getName());
    it = f.getParents();
  }
  return names.join("/");
}

function tzAuditRun() {
  const t0 = Date.now();
  const props = PropertiesService.getScriptProperties();
  const st = JSON.parse(props.getProperty(TZA_KEY) || '{"last":"","files":0,"off":0,"errors":0}');
  const ids = [];
  const it = DriveApp.searchFiles("mimeType = 'application/vnd.google-apps.spreadsheet'" +
                                  " and 'me' in owners and trashed = false");
  while (it.hasNext()) ids.push(it.next().getId());
  ids.sort();
  const todo = ids.filter(function (id) { return id > st.last; });
  Logger.log(ids.length + " classeur(s) à vous ; " + todo.length + " restant(s).");

  for (var i = 0; i < todo.length; i++) {
    if (Date.now() - t0 > TZA_BUDGET_MS) {
      props.setProperty(TZA_KEY, JSON.stringify(st));
      Logger.log("PAS FINI : " + st.files + " lu(s), " + (todo.length - i) + " restant(s). Relancer tzAuditRun.");
      return;
    }
    const id = todo[i];
    try {
      const file = DriveApp.getFileById(id);
      const path = tzAuditPath(file);
      if (path !== null) {
        const tz = SpreadsheetApp.openById(id).getSpreadsheetTimeZone();
        if (tz !== TZA_TARGET) {
          st.off++;
          Logger.log("TZ | " + tz + " | " + file.getName() + " | " + id + " | " + path);
        }
      }
    } catch (e) {
      st.errors++;
      Logger.log("ERREUR | " + id + " | " + e.message);
    }
    st.files++;
    st.last = id;
  }
  props.deleteProperty(TZA_KEY);
  Logger.log("TERMINÉ : " + st.files + " classeur(s) lu(s), " + st.off +
             " pas sur " + TZA_TARGET + ", " + st.errors + " erreur(s).");
}
