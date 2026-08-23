/***************************************************************
 * DataBassinsServer.js — TSARA Entry web app, screen 9
 *
 * Weekly turbidity per bassin (1-10), "noeud" scale 1-4.
 *
 * SCALE (ruled by Kim 2026-08-23 — matches WARNINGSYSTEM and the
 * actionsurstock visualisation):
 *   noeud 1 = eau très trouble  -> changer l'eau urgemment
 *   noeud 2 = eau trouble       -> changer l'eau
 *   noeud 3 = OK
 *   noeud 4 = parfait (eau claire)
 * LOW noeud = DIRTY water. The instruction block in the legacy
 * sheet said the opposite; that text is wrong.
 *
 * Storage: "Data Bassins ERP" (CFG.DATA_BASSINS_SS_ID in Code.js),
 * tab "journal", append-only:
 *   timestamp | date_mesure | bassin | parametre | valeur | operateur
 * A correction is a NEW row. The LAST row wins for the same
 * (date_mesure, bassin, parametre). Nothing edits or deletes rows.
 *
 * Readers of the same journal (patched the same day):
 *   WARNINGSYSTEM  ws_checkTurbidity_     alert at noeud <= 2
 *   actionsurstock fetchTurbiditeAlerts_  red pond title at <= 2
 *
 * dabImportProbe / dabImportLegacy: ONE-TIME migration of the
 * legacy grid (the old "Data Bassins" file). Run dabImportProbe,
 * read the log, then run dabImportLegacy ONCE. Both functions are
 * removed by a follow-up patch after the import.
 ***************************************************************/

const DAB_CFG = {
  SHEET: "journal",
  BASSIN_COUNT: 10,
  PARAM: "turbidite",
  NOEUD_MIN: 1,
  NOEUD_MAX: 4,
  ALERT_MAX: 2,            // noeud 1-2 = eau trouble
  HISTORY_WEEKS: 12,
  HEADERS: ["timestamp", "date_mesure", "bassin", "parametre", "valeur", "operateur"],

  // Legacy grid — used ONLY by the one-time import below.
  LEGACY_SS_ID: "1PcBCz8MpVNswEoh14IO3t9Xz9aaYK_LDhrRNX78yHnE",
  LEGACY_SHEET: "turbidité bassins",
  // If the week labels in row 1 are TEXT ("20 avril", no year), the
  // anchor pins column B to Monday 2026-04-20; each next column adds
  // 7 days, and every computed date is verified against its label.
  // Real Date cells are used as-is.
  LEGACY_ANCHOR: "2026-04-20"
};

const DAB_FR_MONTHS = ["janvier", "février", "mars", "avril", "mai", "juin",
  "juillet", "août", "septembre", "octobre", "novembre", "décembre"];

/** The journal sheet. Throws a French message if it is missing. */
function dabSheet() {
  const sh = SpreadsheetApp.openById(CFG.DATA_BASSINS_SS_ID)
    .getSheetByName(DAB_CFG.SHEET);
  if (!sh) {
    throw new Error('Onglet "' + DAB_CFG.SHEET + '" introuvable dans Data Bassins ERP.');
  }
  return sh;
}

/** Today in the script timezone, "yyyy-MM-dd". */
function dabTodayIso() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/**
 * The signed-in worker's e-mail. Local copy of the Tracabilité
 * logic ON PURPOSE: that machinery is slated for removal when the
 * tracform route is rebuilt, and this screen must not die with it.
 */
function dabOperatorEmail() {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error(
      "Impossible d'identifier l'opérateur. Vérifiez que vous êtes connecté " +
      "avec votre compte @jdsresearch.com, puis rechargez la page.");
  }
  return email;
}

/**
 * Saves the selected measures. entries = [{bassin, valeur}, ...].
 * All rows are checked before any row is written. Returns
 * { saved, dateIso, alerts: [{bassin, valeur}] } — alerts are the
 * saved entries at noeud <= 2 (eau trouble).
 */
function dabSave(dateIso, entries) {
  const iso = String(dateIso || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) {
    throw new Error("Date invalide : " + iso + ".");
  }
  if (iso > dabTodayIso()) {
    throw new Error("La date " + iso + " est dans le futur.");
  }
  const p = iso.split("-");
  const dateMesure = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  if (Utilities.formatDate(dateMesure, Session.getScriptTimeZone(), "yyyy-MM-dd") !== iso) {
    throw new Error("Date invalide : " + iso + ".");
  }

  if (!entries || !entries.length) {
    throw new Error("Aucune mesure sélectionnée.");
  }
  const seen = {};
  const clean = [];
  for (let i = 0; i < entries.length; i++) {
    const b = Number(entries[i] && entries[i].bassin);
    const v = Number(entries[i] && entries[i].valeur);
    if (!(b >= 1 && b <= DAB_CFG.BASSIN_COUNT) || b !== Math.floor(b)) {
      throw new Error("Bassin invalide : " + String(entries[i] && entries[i].bassin) + ".");
    }
    if (!(v >= DAB_CFG.NOEUD_MIN && v <= DAB_CFG.NOEUD_MAX) || v !== Math.floor(v)) {
      throw new Error("Bassin " + b + " : noeud invalide (" +
        String(entries[i] && entries[i].valeur) + "). Valeurs permises : 1 à 4.");
    }
    if (seen[b]) {
      throw new Error("Bassin " + b + " envoyé deux fois.");
    }
    seen[b] = true;
    clean.push({ bassin: b, valeur: v });
  }

  const email = dabOperatorEmail();
  const now = new Date();
  const rows = clean.map(function (e) {
    return [now, dateMesure, e.bassin, DAB_CFG.PARAM, e.valeur, email];
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = dabSheet();
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  const alerts = clean
    .filter(function (e) { return e.valeur <= DAB_CFG.ALERT_MAX; })
    .sort(function (a, b) { return a.bassin - b.bassin; });
  return { saved: rows.length, dateIso: iso, alerts: alerts };
}

/**
 * History matrix for the Historique view. Weeks are grouped on the
 * MONDAY of each week; the last entry in the journal wins for a
 * (bassin, week) cell. Returns
 * { weeks: ["yyyy-MM-dd" x 12, ascending, current week last],
 *   rows: [10 arrays of (null | {v, d, op})] }.
 */
function dabGetHistory() {
  const sh = dabSheet();
  const lastRow = sh.getLastRow();
  const tz = Session.getScriptTimeZone();

  const today = new Date();
  const monday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  monday.setDate(monday.getDate() - ((monday.getDay() + 6) % 7));

  const weeks = [];
  const index = {};
  for (let i = DAB_CFG.HISTORY_WEEKS - 1; i >= 0; i--) {
    const d = new Date(monday);
    d.setDate(d.getDate() - 7 * i);
    const iso = Utilities.formatDate(d, tz, "yyyy-MM-dd");
    index[iso] = weeks.length;
    weeks.push(iso);
  }

  const rows = [];
  for (let b = 0; b < DAB_CFG.BASSIN_COUNT; b++) {
    rows.push(new Array(weeks.length).fill(null));
  }

  if (lastRow >= 2) {
    const vals = sh.getRange(2, 1, lastRow - 1, 6).getValues();
    for (let r = 0; r < vals.length; r++) {
      if (String(vals[r][3]).trim() !== DAB_CFG.PARAM) continue;
      const b = Number(vals[r][2]);
      if (!(b >= 1 && b <= DAB_CFG.BASSIN_COUNT)) continue;
      const d = vals[r][1];
      if (!(d instanceof Date)) continue;
      const m = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      m.setDate(m.getDate() - ((m.getDay() + 6) % 7));
      const iso = Utilities.formatDate(m, tz, "yyyy-MM-dd");
      if (!(iso in index)) continue;
      rows[b - 1][index[iso]] = {
        v: Number(vals[r][4]),
        d: Utilities.formatDate(d, tz, "yyyy-MM-dd"),
        op: String(vals[r][5] || "")
      };
    }
  }
  return { weeks: weeks, rows: rows };
}

/* ---------------- one-time import of the legacy grid ---------------- */

/** Read-only. Logs what the import WOULD write. Writes nothing. */
function dabImportProbe() {
  dabImportScan(false);
}

/** Writes the scan result into the journal. Refuses a non-empty journal. */
function dabImportLegacy() {
  dabImportScan(true);
}

function dabImportScan(write) {
  const tz = Session.getScriptTimeZone();
  const legacy = SpreadsheetApp.openById(DAB_CFG.LEGACY_SS_ID)
    .getSheetByName(DAB_CFG.LEGACY_SHEET);
  if (!legacy) {
    throw new Error('Onglet "' + DAB_CFG.LEGACY_SHEET + '" introuvable dans le fichier legacy.');
  }
  const lastRow = legacy.getLastRow();
  const lastCol = legacy.getLastColumn();
  const grid = legacy.getRange(1, 1, lastRow, lastCol).getValues();

  const ap = DAB_CFG.LEGACY_ANCHOR.split("-");
  const anchor = new Date(Number(ap[0]), Number(ap[1]) - 1, Number(ap[2]));

  // Resolve the dated columns of row 1 (column B onward).
  const todayIso = dabTodayIso();
  const cols = [];
  let realDates = 0;
  let labelDates = 0;
  for (let c = 1; c < lastCol; c++) {
    const v = grid[0][c];
    let d = null;
    if (v instanceof Date) {
      d = new Date(v.getFullYear(), v.getMonth(), v.getDate());
      realDates++;
    } else {
      const label = String(v || "").trim().toLowerCase();
      if (!label) continue;
      d = new Date(anchor);
      d.setDate(d.getDate() + 7 * (c - 1));
      const expect = d.getDate() + " " + DAB_FR_MONTHS[d.getMonth()];
      if (label !== expect) {
        throw new Error("Colonne " + (c + 1) + ' : étiquette "' + label +
          '" ne correspond pas à la date calculée "' + expect + '". Import arrêté.');
      }
      labelDates++;
    }
    const iso = Utilities.formatDate(d, tz, "yyyy-MM-dd");
    if (iso <= todayIso) cols.push({ c: c, iso: iso, date: d });
  }
  Logger.log("Colonnes datées retenues (jusqu'à aujourd'hui) : " + cols.length +
    " — " + realDates + " dates réelles, " + labelDates +
    " étiquettes texte ancrées sur " + DAB_CFG.LEGACY_ANCHOR + ".");

  // Walk the bassin blocks: "Bassin N" label, "Turbidité" row beneath.
  const now = new Date();
  const out = [];
  const anomalies = [];
  let plusCount = 0;
  for (let r = 0; r < lastRow; r++) {
    const lab = String(grid[r][0] || "").trim();
    const mB = lab.toLowerCase().match(/^bassin\s+(\d+)$/);
    if (!mB) continue;
    const bassin = Number(mB[1]);
    let tr = -1;
    for (let rr = r + 1; rr < Math.min(lastRow, r + 4); rr++) {
      if (String(grid[rr][0] || "").trim().toLowerCase().indexOf("turbidit") === 0) {
        tr = rr;
        break;
      }
    }
    if (tr < 0) {
      anomalies.push("Bassin " + bassin + " : pas de ligne Turbidité — ignoré.");
      continue;
    }
    if (bassin < 1 || bassin > DAB_CFG.BASSIN_COUNT) {
      anomalies.push("Bassin " + bassin + " hors plage 1-" + DAB_CFG.BASSIN_COUNT + " — ignoré.");
      continue;
    }

    let count = 0;
    let plus = 0;
    let first = "";
    let last = "";
    for (let k = 0; k < cols.length; k++) {
      const cell = String(grid[tr][cols[k].c] || "").trim();
      if (!cell) continue;
      const m = cell.toLowerCase().match(/noeud\s*(\d+)(\+?)/);
      if (!m) {
        anomalies.push("Bassin " + bassin + " · " + cols[k].iso +
          ' : valeur illisible "' + cell + '" — ignorée.');
        continue;
      }
      let n = parseInt(m[1], 10);
      if (m[2] === "+") { plus++; plusCount++; }
      if (n > DAB_CFG.NOEUD_MAX) {
        anomalies.push("Bassin " + bassin + " · " + cols[k].iso +
          ' : "' + cell + '" ramené à ' + DAB_CFG.NOEUD_MAX + ".");
        n = DAB_CFG.NOEUD_MAX;
      }
      if (n < DAB_CFG.NOEUD_MIN) {
        anomalies.push("Bassin " + bassin + " · " + cols[k].iso +
          ' : "' + cell + '" ignoré (noeud < 1).');
        continue;
      }
      out.push([now, cols[k].date, bassin, DAB_CFG.PARAM, n, "import"]);
      count++;
      if (!first) first = cols[k].iso;
      last = cols[k].iso;
    }
    Logger.log("Bassin " + bassin + " : " + count + " mesures (" + first + " → " + last + ")" +
      (plus ? ' · ' + plus + ' valeurs "+"' : ""));
  }

  for (let i = 0; i < anomalies.length; i++) {
    Logger.log("ANOMALIE — " + anomalies[i]);
  }
  Logger.log("TOTAL : " + out.length + " mesures · " + plusCount +
    ' valeurs "noeud N+" (le "+" est abandonné, N est conservé).');

  if (!write) {
    Logger.log("PROBE — rien n'a été écrit.");
    return;
  }

  const sh = dabSheet();
  if (sh.getLastRow() !== 1) {
    throw new Error("Le journal n'est pas vide (dernière ligne " + sh.getLastRow() +
      "). Import refusé — il a peut-être déjà été fait.");
  }
  if (!out.length) {
    throw new Error("Aucune mesure à importer.");
  }
  sh.getRange(2, 1, out.length, 6).setValues(out);
  SpreadsheetApp.flush();
  Logger.log("IMPORT : " + out.length + " lignes écrites dans le journal.");
}

/**
 * Read-only check. Run from the editor after the push and BEFORE
 * the redeployment. Project: TSARA Entry -> DataBassinsServer.js.
 */
function testDataBassinsServer() {
  const sh = dabSheet();
  Logger.log("Journal : onglet " + sh.getName() + " · gid " + sh.getSheetId() +
    " · dernière ligne " + sh.getLastRow());
  const head = sh.getRange(1, 1, 1, 6).getValues()[0].map(function (v) {
    return String(v).trim();
  });
  Logger.log("En-têtes : " + head.join(" | ") +
    (head.join("|") === DAB_CFG.HEADERS.join("|")
      ? " — OK"
      : " — ÉCART, attendu : " + DAB_CFG.HEADERS.join(" | ")));
  const h = dabGetHistory();
  Logger.log("Historique : " + h.weeks.length + " semaines · " + h.rows.length +
    " bassins · semaine courante " + h.weeks[h.weeks.length - 1]);
  const legacy = SpreadsheetApp.openById(DAB_CFG.LEGACY_SS_ID)
    .getSheetByName(DAB_CFG.LEGACY_SHEET);
  Logger.log("Legacy : " + (legacy
    ? 'onglet "' + legacy.getName() + '" · ' + legacy.getLastRow() + " lignes × " +
      legacy.getLastColumn() + " colonnes"
    : "ONGLET INTROUVABLE"));
}
