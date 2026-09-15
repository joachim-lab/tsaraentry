/***************************************************************
 * AccessServer.js — TSARA Entry: per-user screen access (2026-09-11)
 *
 * WHAT IT DOES. Each account sees and opens only the screens listed
 * for its e-mail in the tab "acces" of the spreadsheet
 * "Accès Tsara Entry" (script property ACCESS_SS_ID).
 *   A = email   B = nom   C = écrans
 * C holds screen keys separated by commas, or * for every screen:
 *   nourrissage, inventaires, impressions, databassins, creerlot,
 *   commandes, morts, lot, tracabilite
 * Fine-grained rights INSIDE the Commandes screen (tabs, fields,
 * buttons) live in a second tab of the same file: "commandes",
 * one row per restricted account, one checkbox per element.
 * No row = the full screen. See accessCommandesRights().
 * The first row with the e-mail wins. Unknown words in C are ignored
 * (testAccess lists them).
 *
 * FAIL CLOSED. Blank e-mail or e-mail not in the tab: no screen.
 * Missing property, file or tab: an error, so no screen either.
 *
 * WHAT IT DOES NOT DO. It gates doGet (Router.js) and the menu tiles
 * (Menu.html). It does NOT guard server functions: any public
 * function stays callable from the browser console. This is workflow
 * control, not security (Kim, 2026-09-11, option 1 of 3).
 *
 * CACHE. The keys are cached 5 min in the SCRIPT cache, keyed by
 * e-mail. Not the user cache: the app runs as Kim. An edit of the tab
 * applies within 5 min; testAccess applies it at once.
 *
 * FIRST RUN. tsaraentry -> AccessServer.js -> testAccess creates the
 * file, seeds today's 4 accounts with *, and stores the property.
 * File deleted -> delete the property, run testAccess again.
 * Kim owns the file. Staff do not need it shared.
 ***************************************************************/

var ACCESS_KEYS = ["nourrissage", "inventaires", "impressions", "databassins",
  "creerlot", "commandes", "morts", "lot", "tracabilite"];

/** ?screen= value -> access key. A sub-screen takes its tile's key. */
var ACCESS_SCREEN_KEY = {
  nourrissage: "nourrissage",
  gestioninventaires: "inventaires", inventaire: "inventaires",
  checkstock: "inventaires", achats: "inventaires",
  impressions: "impressions",
  databassins: "databassins",
  creerlot: "creerlot",
  commandes: "commandes",
  morts: "morts",
  lot: "lot",
  tracabilite: "tracabilite", tracform: "tracabilite"
};

var ACCESS_PROP = "ACCESS_SS_ID";
var ACCESS_TAB = "acces";
var ACCESS_CACHE_SEC = 300;
var ACCESS_SEED = [
  ["joachim@jdsresearch.com", "Joachim", "*"],
  ["hasina@jdsresearch.com", "Hasina", "*"],
  ["audry@jdsresearch.com", "Audry", "*"],
  ["charles@jdsresearch.com", "Charles", "*"]
];

/**
 * Pure. The access keys of one e-mail, from the tab values
 * (row 0 = header). Returns [] when no row matches.
 */
function accessParse(values, email) {
  const e = String(email || "").trim().toLowerCase();
  if (!e) return [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || "").trim().toLowerCase() !== e) continue;
    const c = String(values[i][2] || "").trim().toLowerCase();
    if (c === "*") return ACCESS_KEYS.slice();
    const want = c.split(",").map(function (w) { return w.trim(); });
    return ACCESS_KEYS.filter(function (k) { return want.indexOf(k) >= 0; });
  }
  return [];
}

/** The current account's e-mail, lower case. "" when Google gives none. */
function accessEmail() {
  return String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
}

/** Values of the acces tab. Throws when the property, file or tab is missing. */
function accessReadTab() {
  const id = PropertiesService.getScriptProperties().getProperty(ACCESS_PROP);
  if (!id) {
    throw new Error("Table d'accès absente : lancer tsaraentry -> AccessServer.js -> testAccess.");
  }
  const sh = SpreadsheetApp.openById(id).getSheetByName(ACCESS_TAB);
  if (!sh) throw new Error("Onglet « " + ACCESS_TAB + " » introuvable dans « Accès Tsara Entry ».");
  return sh.getDataRange().getValues();
}

/** The keys the current account may open. Cached ACCESS_CACHE_SEC. */
function accessMenuKeys() {
  const email = accessEmail();
  if (!email) return [];
  const cache = CacheService.getScriptCache();
  const hit = cache.get("acc:" + email);
  if (hit !== null) return JSON.parse(hit);
  const keys = accessParse(accessReadTab(), email);
  cache.put("acc:" + email, JSON.stringify(keys), ACCESS_CACHE_SEC);
  return keys;
}

/** True when the current account may open this ?screen= value. */
function accessAllows(screen) {
  if (!Object.prototype.hasOwnProperty.call(ACCESS_SCREEN_KEY, screen)) return false;
  return accessMenuKeys().indexOf(ACCESS_SCREEN_KEY[screen]) >= 0;
}

/***** Droits fins dans l'écran Commandes (onglet "commandes") *****/

var ACCESS_CMD_TAB = "commandes";

/**
 * Column header (row 1 of the tab, exact text) -> right key.
 * The first two columns identify the account; every later column is a
 * checkbox: checked = allowed. To make a NEW element controllable, add
 * a line here, add the column in the sheet, and test the element
 * against RIGHTS.<key> in CommandesIndex.html. A sheet without the
 * column defaults to allowed.
 */
var ACCESS_CMD_COLS = [
  ["email", "email"],
  ["nom", "nom"],
  ["Pré-commandes", "dem"],
  ["Nouvelle cmd", "new"],
  ["Récurrentes", "rec"],
  ["Paiement & livraison", "ful"],
  ["Historique", "hist"],
  ["Prix", "prix"],
  ["Clients", "cli"],
  ["Résas", "res"],
  ["Champ date paiement", "paiement"],
  ["Champ date livraison", "livraison"],
  ["Champ reçue le", "reception"],
  ["Champ téléphone", "telephone"],
  ["Champ N° facture", "nfacture"],
  ["Bouton Modifier", "modifier"],
  ["Bouton Annuler", "annuler"],
  ["Bouton Facture", "facture"]
];

/**
 * Pure. Rights of one e-mail in the Commandes screen, from the grid
 * values (row 0 = header). No row for the e-mail, or no grid at all:
 * every right true - the screen itself is already gated by the tab
 * "acces". A row that exists: checked = true, anything else = false,
 * so an all-blank row hides everything. A header column missing from
 * the sheet defaults to true.
 */
function accessCmdParse(values, email) {
  const out = {};
  for (let i = 2; i < ACCESS_CMD_COLS.length; i++) out[ACCESS_CMD_COLS[i][1]] = true;
  const e = String(email || "").trim().toLowerCase();
  if (!e || !values || values.length < 2) return out;
  const head = values[0].map(function (h) { return String(h || "").trim().toLowerCase(); });
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0] || "").trim().toLowerCase() !== e) continue;
    for (let i = 2; i < ACCESS_CMD_COLS.length; i++) {
      const c = head.indexOf(ACCESS_CMD_COLS[i][0].toLowerCase());
      out[ACCESS_CMD_COLS[i][1]] = c < 0 ? true : values[r][c] === true;
    }
    return out;
  }
  return out;
}

/**
 * The rights of the current account in the Commandes screen. Injected
 * into CommandesIndex.html at page build as RIGHTS. Cached like
 * accessMenuKeys; testAccess clears the cache at once. Tab missing
 * from the file: every right true (the tab appears at the next
 * testAccess run).
 */
function accessCommandesRights() {
  const email = accessEmail();
  const cache = CacheService.getScriptCache();
  const hit = cache.get("accCmd:" + email);
  if (hit !== null) return JSON.parse(hit);
  const id = PropertiesService.getScriptProperties().getProperty(ACCESS_PROP);
  const sh = id ? SpreadsheetApp.openById(id).getSheetByName(ACCESS_CMD_TAB) : null;
  const rights = accessCmdParse(sh ? sh.getDataRange().getValues() : null, email);
  cache.put("accCmd:" + email, JSON.stringify(rights), ACCESS_CACHE_SEC);
  return rights;
}

function accessEscape(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** The page served instead of a refused screen. */
function accessDeniedOutput() {
  const email = accessEmail();
  const who = email
    ? "Compte : " + accessEscape(email)
    : "Compte non identifié : connectez-vous avec votre compte @jdsresearch.com.";
  const html = '<!DOCTYPE html><html><head><base target="_top"></head>' +
    '<body style="font-family:Arial,sans-serif;max-width:600px;margin:64px auto;' +
    'padding:0 16px;text-align:center;color:#202124">' +
    '<h1 style="color:#146E3C;font-size:24px">Accès refusé</h1>' +
    '<p>Cet écran n’est pas autorisé pour ce compte.</p>' +
    '<p style="color:#5f6368">' + who + '</p>' +
    '<p><a href="' + accessEscape(getWebAppUrl()) + '">&larr; Retour au menu</a></p>' +
    '</body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle("Interface Tsara Tilapia")
    .setFaviconUrl(TSARA_FAVICON)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Editor run: tsaraentry -> AccessServer.js -> testAccess.
 * First run: creates "Accès Tsara Entry", seeds 4 accounts with *.
 * Every run: logs the file URL, the keys of each row, unknown words,
 * your own keys, and clears the cache of every e-mail in the tab.
 */
function testAccess() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(ACCESS_PROP)) {
    const ss = SpreadsheetApp.create("Accès Tsara Entry");
    const sh = ss.getSheets()[0].setName(ACCESS_TAB);
    sh.getRange(1, 1, 1, 3).setValues([["email", "nom", "écrans"]]).setFontWeight("bold");
    sh.getRange(2, 1, ACCESS_SEED.length, 3).setValues(ACCESS_SEED);
    sh.setFrozenRows(1);
    props.setProperty(ACCESS_PROP, ss.getId());
    console.log("Créé : Accès Tsara Entry");
  }
  const ssA = SpreadsheetApp.openById(props.getProperty(ACCESS_PROP));
  if (!ssA.getSheetByName(ACCESS_CMD_TAB)) {
    const shC = ssA.insertSheet(ACCESS_CMD_TAB);
    const heads = ACCESS_CMD_COLS.map(function (c) { return c[0]; });
    shC.getRange(1, 1, 1, heads.length).setValues([heads]).setFontWeight("bold");
    shC.getRange(2, 3, 50, heads.length - 2).insertCheckboxes();
    shC.setFrozenRows(1);
    shC.setFrozenColumns(2);
    console.log("Créé : onglet « " + ACCESS_CMD_TAB + " » (droits par compte, écran Commandes)");
  }
  // A grid column added to ACCESS_CMD_COLS after the tab was created is
  // appended at the right, with checkboxes. Rows that already hold an
  // e-mail get it TICKED: a missing column meant "allowed", so nothing
  // changes for anyone until Kim unticks a box.
  const shX = ssA.getSheetByName(ACCESS_CMD_TAB);
  const headX = shX.getRange(1, 1, 1, Math.max(1, shX.getLastColumn())).getValues()[0]
    .map(function (h) { return String(h || "").trim().toLowerCase(); });
  const rowsX = Math.max(shX.getLastRow(), 51) - 1;   // the seeded 50 rows, or more if used
  const emailsX = shX.getRange(2, 1, rowsX, 1).getValues();
  ACCESS_CMD_COLS.slice(2).forEach(function (col) {
    if (headX.indexOf(col[0].toLowerCase()) >= 0) return;
    const c = headX.length + 1;
    shX.getRange(1, c).setValue(col[0]).setFontWeight("bold");
    shX.getRange(2, c, rowsX, 1).insertCheckboxes()
      .setValues(emailsX.map(function (r) { return [String(r[0] || "").trim() !== ""]; }));
    headX.push(col[0].toLowerCase());
    console.log("Colonne ajoutée : « " + col[0] + " » (cochée pour les comptes existants)");
  });
  const values = accessReadTab();
  const cacheKeys = [];
  for (let i = 1; i < values.length; i++) {
    const e = String(values[i][0] || "").trim().toLowerCase();
    if (!e) continue;
    cacheKeys.push("acc:" + e);
    const c = String(values[i][2] || "").trim().toLowerCase();
    const unknown = c === "*" ? [] : c.split(",").map(function (w) { return w.trim(); })
      .filter(function (w) { return w && ACCESS_KEYS.indexOf(w) < 0; });
    console.log("Ligne " + (i + 1) + " " + e + " -> " + JSON.stringify(accessParse(values, e)) +
      (unknown.length ? "  MOTS INCONNUS : " + unknown.join(", ") : ""));
  }
  const shG = ssA.getSheetByName(ACCESS_CMD_TAB);
  const gv = shG ? shG.getDataRange().getValues() : null;
  if (gv) {
    for (let g = 1; g < gv.length; g++) {
      const ge = String(gv[g][0] || "").trim().toLowerCase();
      if (!ge) continue;
      cacheKeys.push("accCmd:" + ge);
      console.log("commandes " + ge + " -> " + JSON.stringify(accessCmdParse(gv, ge)));
    }
  }
  if (cacheKeys.length) CacheService.getScriptCache().removeAll(cacheKeys);
  console.log("Fichier : " + SpreadsheetApp.openById(props.getProperty(ACCESS_PROP)).getUrl());
  console.log("Vous (" + accessEmail() + ") -> " + JSON.stringify(accessMenuKeys()));
  console.log("Clés valides : " + ACCESS_KEYS.join(", "));
}
