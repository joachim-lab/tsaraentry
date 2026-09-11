/***************************************************************
 * FactureServer.js - TSARA Entry web app
 * FACTURE PDF (Kim, 2026-09-11): the "Imprimer" button in the
 * "Facture" column of Commandes > Historique.
 *
 * READS tab "2026" and tab "CRM" of the Commandes file.
 * WRITES ONLY the PDF, in the Drive folder "Factures".
 *
 * ONE INVOICE = ONE N° FACTURE (col C), NOT ONE ORDER. Finance puts
 * several orders on one invoice (FA26095 covers AL-26-127 and
 * GR-26-58, see automatismescommande 05_facture_bl.js). The invoice
 * therefore takes every non-cancelled row that carries the number of
 * the clicked order. Both orders print the same invoice.
 *
 * REFUSED, NEVER GUESSED:
 *   - the order has no N° facture (it is minted at AE, "reçue")
 *   - the order is cancelled (AA)
 *   - the rows of the invoice name two different clients (R)
 *   - the line total differs from K + Q of the same rows by more
 *     than 1 Ar. The sheet then holds an amount this code cannot
 *     explain, and an invoice must not show a different total.
 *
 * LINES, one per row of tab "2026":
 *   F > 0   "Alevins <G> g"          PU = I   Qté = F    F x I
 *   L > 0   "Poisson grossi <M> g"   PU = O   Qté = L    L x O
 *   J and P are summed into one "Transport" line.
 *   AF > 0  one "Remise" line = -(F x I + L x O) x AF / 100, per row.
 *           Same arithmetic as cmdRemiseFormulas: the remise is on the
 *           fish / fry price only, never on transport.
 *   AF is read RAW (getValues): the column has carried a date format.
 *
 * DATE FACTURE: the latest AE (date reçue) of the rows. An invoice
 * minted before AE existed was minted at delivery and has no AE:
 * the latest V (date livraison) is its minting date.
 *
 * CLIENT BLOCK: the name from R. Localisation (C), téléphone (B),
 * NIF (K) and STAT (L) from the CRM row with the same canonical name.
 * A blank field is left off the invoice.
 *
 * THE FILE: "<N° facture> - <client>.pdf" in the Drive folder whose id
 * is the script property FACT_FOLDER_ID. The first testFacture run
 * creates the folder "Factures" in Kim's My Drive. A reprint puts the
 * old file of the same name in the bin and writes a new one.
 * The web app runs as Kim, so Kim owns the folder: staff need it
 * SHARED (lecteur is enough), or they get "Accès refusé".
 * Folder deleted -> delete the property, run testFacture, share the
 * new folder again.
 ***************************************************************/

const FACT_CFG = {
  PROP: "FACT_FOLDER_ID",
  FOLDER: "Factures",
  // Printed before the number of column C, as on the paper model
  // (TSRLP- FA26023). Not added when column C already starts with it.
  PREFIX: "TSRLP-",
  SOCIETE: ["TSARATILAPIA SA", "Enceinte Activo Andaona Ankeniheny",
            "Arivonimamo I", "Madagascar"],
  // Conditions de règlement: one entry per printed line.
  REGLEMENT: ["Espèces",
              "Chèque à l'ordre de « TSARATILAPIA SA »",
              "Virement bancaire BMOI 00001-05079520101 34",
              "Mobile money : Mvola 038 85 402 65",
              "Mobile money : Orange money 037 78 745 15"],
  TOLERANCE_AR: 1
};

/* ---------- numbers ---------- */

/** 17600 -> "17 600,00" (dec 2), 60 -> "60" (dec 0). Non-breaking spaces. */
function factNum(n, dec) {
  const s = Math.abs(n).toFixed(dec);
  const p = s.split(".");
  const i = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  return (n < 0 ? "-" : "") + i + (dec ? "," + p[1] : "");
}

/** Kg: "60,00". Alevins: "1 000". */
function factQty(n, isKg) {
  return isKg ? factNum(n, 2) : factNum(Math.round(n), 0);
}

/**
 * Whole number -> French words, traditional spelling.
 * 1056000 -> "un million cinquante-six mille".
 * "cents" and "quatre-vingts" take the s only at the end of the
 * number or before million / milliard, never before mille.
 */
function factWords(n) {
  n = Math.round(Math.abs(Number(n) || 0));
  if (n === 0) return "zéro";
  const U = ["", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit",
             "neuf", "dix", "onze", "douze", "treize", "quatorze", "quinze",
             "seize", "dix-sept", "dix-huit", "dix-neuf"];
  const T = ["", "", "vingt", "trente", "quarante", "cinquante", "soixante",
             "soixante", "quatre-vingt", "quatre-vingt"];

  function below100(x, plural) {
    if (x < 20) return U[x];
    const t = Math.floor(x / 10), u = x % 10;
    if (t === 7 || t === 9) {                       // 70-79, 90-99
      return T[t] + (t === 7 && u === 1 ? " et " : "-") + U[10 + u];
    }
    if (u === 0) return T[t] + (t === 8 && plural ? "s" : "");
    if (u === 1 && t !== 8) return T[t] + " et un";
    return T[t] + "-" + U[u];
  }
  function below1000(x, plural) {                  // 1..999
    const h = Math.floor(x / 100), r = x % 100;
    var s = "";
    if (h > 0) s = (h > 1 ? U[h] + " " : "") + "cent" + (h > 1 && r === 0 && plural ? "s" : "");
    if (r > 0) s += (s ? " " : "") + below100(r, plural);
    return s;
  }

  const parts = [];
  const mrd = Math.floor(n / 1e9), mio = Math.floor(n / 1e6) % 1000,
        mil = Math.floor(n / 1e3) % 1000, rest = n % 1000;
  if (mrd) parts.push(below1000(mrd, true) + " milliard" + (mrd > 1 ? "s" : ""));
  if (mio) parts.push(below1000(mio, true) + " million" + (mio > 1 ? "s" : ""));
  if (mil) parts.push(mil === 1 ? "mille" : below1000(mil, false) + " mille");
  if (rest) parts.push(below1000(rest, true));
  return parts.join(" ");
}

/** Cell -> Date or null. Accepts a Date or a "dd/mm/yyyy" text. */
function factDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const m = String(v == null ? "" : v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (!m) return null;
  const y = Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]);
  const d = new Date(y, Number(m[2]) - 1, Number(m[1]));
  return isNaN(d.getTime()) ? null : d;
}

function factHtmlEsc(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/* ---------- data ---------- */

/**
 * Everything the invoice prints, for the order orderNumber (col B).
 * Read-only. Throws a French message on every refusal.
 */
function factData(orderNumber) {
  const want = String(orderNumber == null ? "" : orderNumber).trim();
  if (!want) throw new Error("Numéro de commande manquant.");

  const sh = cmdSheet();
  const C = CMD_CFG.COL;
  const tz = sh.getParent().getSpreadsheetTimeZone();
  const lastRow = findNextCommandeRow(sh) - 1;
  if (lastRow < CMD_CFG.START_ROW) throw new Error("Aucune commande dans l'onglet " + CMD_CFG.SHEET + ".");
  const n = lastRow - CMD_CFG.START_ROW + 1;
  const v = sh.getRange(CMD_CFG.START_ROW, 1, n, C.REMISE).getValues();
  const cell = function (r, col) { return v[r][col - 1]; };
  const txt = function (r, col) { return String(cell(r, col) == null ? "" : cell(r, col)).trim(); };

  // 1. The clicked order: its number of invoice, and not cancelled.
  var facture = "";
  var seen = false;
  for (var i = 0; i < n; i++) {
    if (txt(i, C.ORDER_NO) !== want) continue;
    seen = true;
    if (txt(i, C.ANNULE)) throw new Error("La commande " + want + " est annulée : pas de facture.");
    if (!facture && txt(i, C.FACTURE)) facture = txt(i, C.FACTURE);
  }
  if (!seen) throw new Error("Commande " + want + " introuvable dans l'onglet " + CMD_CFG.SHEET + ".");
  if (!facture) throw new Error("La commande " + want + " n'a pas de numéro de facture. " +
                                "Le numéro est créé quand la commande est reçue.");

  // 2. Every non-cancelled row of that invoice.
  const lines = [];
  const orders = [];                 // [{ no, date }], first-seen order
  const orderSeen = {};
  const clients = {};
  var client = "";
  var transport = 0, remise = 0, sheetTotal = 0;
  var remisePct = {};
  var dateRecu = null, dateLivr = null;

  for (var j = 0; j < n; j++) {
    if (txt(j, C.FACTURE) !== facture || txt(j, C.ANNULE)) continue;
    const rowNum = CMD_CFG.START_ROW + j;

    const name = txt(j, C.CLIENT).replace(/\s+/g, " ");
    if (name) {
      clients[crmCanonName(name)] = true;
      if (!client) client = name;
    }
    const no = txt(j, C.ORDER_NO);
    if (no && !orderSeen[no]) {
      orderSeen[no] = true;
      orders.push({ no: no, date: factDate(cell(j, C.DATE_CMD)) });
    }

    const nbAl = cmdToNum(cell(j, C.ALEVINS_NB)) || 0;
    const prixAl = cmdToNum(cell(j, C.ALEVINS_PRIX)) || 0;
    const kg = cmdToNum(cell(j, C.POISSON_KG)) || 0;
    const prixKg = cmdToNum(cell(j, C.PRIX_KG)) || 0;
    const pct = cmdToNum(cell(j, C.REMISE)) || 0;
    var brut = 0;

    if (nbAl > 0) {
      const pmAl = cmdToNum(cell(j, C.ALEVINS_PM));
      lines.push({ label: "Alevins" + (pmAl ? " " + factNum(pmAl, pmAl % 1 ? 1 : 0) + " g" : ""),
                   pu: prixAl, qty: nbAl, isKg: false, montant: nbAl * prixAl, row: rowNum });
      brut += nbAl * prixAl;
    }
    if (kg > 0) {
      const pmGr = cmdToNum(cell(j, C.POISSON_PM));
      lines.push({ label: "Poisson grossi" + (pmGr ? " " + factNum(Math.round(pmGr), 0) + " g" : ""),
                   pu: prixKg, qty: kg, isKg: true, montant: kg * prixKg, row: rowNum });
      brut += kg * prixKg;
    }
    if (pct > 0) {
      remise += brut * pct / 100;
      remisePct[pct] = true;
    }
    transport += (cmdToNum(cell(j, C.TRANSPORT)) || 0) + (cmdToNum(cell(j, C.FRAIS)) || 0);
    sheetTotal += (cmdToNum(cell(j, C.ARGENT_ALEVINS)) || 0) + (cmdToNum(cell(j, C.ARGENT_POISSON)) || 0);

    const dr = factDate(cell(j, C.RECU));
    if (dr && (!dateRecu || dr > dateRecu)) dateRecu = dr;
    const dl = factDate(cell(j, C.DATE_LIVRAISON));
    if (dl && (!dateLivr || dl > dateLivr)) dateLivr = dl;
  }

  if (!lines.length) throw new Error("Facture " + facture + " : aucune ligne alevins ou poisson.");
  if (Object.keys(clients).length > 1) {
    throw new Error("Facture " + facture + " : les lignes portent plusieurs clients. " +
                    "Corriger la colonne R ou le numéro de facture.");
  }

  if (remise > 0) {
    const pcts = Object.keys(remisePct);
    lines.push({ label: "Remise" + (pcts.length === 1 ? " " + factNum(Number(pcts[0]), Number(pcts[0]) % 1 ? 1 : 0) + " %" : ""),
                 pu: null, qty: null, isKg: false, montant: -remise });
  }
  if (transport > 0) {
    lines.push({ label: "Transport", pu: null, qty: null, isKg: false, montant: transport });
  }

  var total = 0;
  lines.forEach(function (l) { total += l.montant; });
  if (Math.abs(total - sheetTotal) > FACT_CFG.TOLERANCE_AR) {
    throw new Error("Facture " + facture + " : le total calculé (" + factNum(total, 2) +
                    " Ar) ne correspond pas aux colonnes K + Q (" + factNum(sheetTotal, 2) +
                    " Ar). Vérifier les formules de ces lignes. Prévenir Kim.");
  }

  // 3. The client, from the CRM tab.
  const info = { tel: "", loc: "", nif: "", stat: "" };
  const crm = crmEntrySheet();
  const crmLast = crm.getLastRow();
  if (client && crmLast >= CRM_START) {
    const cv = crm.getRange(CRM_START, 1, crmLast - CRM_START + 1, CRM_COLS).getValues();
    const canon = crmCanonName(client);
    for (var k = 0; k < cv.length; k++) {
      if (crmCanonName(cv[k][0]) !== canon) continue;
      const s = function (col) { return String(cv[k][col - 1] == null ? "" : cv[k][col - 1]).trim(); };
      info.tel = s(CRM_COL_TEL);
      info.loc = s(CRM_COL_LOC);
      info.nif = s(CRM_COL_NIF);
      info.stat = s(CRM_COL_STAT);
      break;
    }
  }

  const fmtD = function (d) { return d ? Utilities.formatDate(d, tz, "dd/MM/yyyy") : ""; };
  const dateFacture = dateRecu || dateLivr;
  const numero = facture.toUpperCase().indexOf(FACT_CFG.PREFIX.toUpperCase()) === 0
    ? facture : FACT_CFG.PREFIX + facture;

  return {
    facture: facture,
    numero: numero,
    dateFacture: fmtD(dateFacture),
    commandes: orders.map(function (o) { return o.no + (o.date ? " du " + fmtD(o.date) : ""); }),
    client: client,
    tel: info.tel, loc: info.loc, nif: info.nif, stat: info.stat,
    lines: lines,
    total: total,
    enLettres: factWords(total)
  };
}

/* ---------- HTML ---------- */

/** The invoice page. Tables only: the Apps Script PDF converter
 *  handles tables and borders, not flexbox or grid. */
function factHtml(d) {
  const e = factHtmlEsc;
  const B = "border:1.5px solid #000;";
  const TD = "padding:4px 8px;";
  const cap = d.enLettres.charAt(0).toUpperCase() + d.enLettres.slice(1);

  var h = '<html><head><meta charset="utf-8"><style>' +
    '@page { size: A4; margin: 15mm; }' +
    'body { font-family: Arial, Helvetica, sans-serif; font-size: 10.5pt; color: #000; }' +
    'table { border-collapse: collapse; }' +
    '</style></head><body>';

  // Header: logo + invoice box.
  h += '<table style="width:100%"><tr>' +
    '<td style="width:52%;vertical-align:top"><img src="data:image/jpeg;base64,' + FACT_LOGO_B64 +
      '" style="width:170px"></td>' +
    '<td style="vertical-align:top"><table>' +
      '<tr><td style="' + TD + 'font-weight:bold;white-space:nowrap">N° FACTURE :</td>' +
        '<td style="' + TD + B + 'font-weight:bold;font-size:12pt">' + e(d.numero) + '</td></tr>' +
      '<tr><td style="' + TD + 'vertical-align:top;white-space:nowrap">N° Commande :</td><td style="' + TD + 'white-space:nowrap">' +
        d.commandes.map(e).join('<br>') + '</td></tr>' +
      '<tr><td style="' + TD + 'white-space:nowrap">Date facture :</td><td style="' + TD + '">' + e(d.dateFacture) + '</td></tr>' +
    '</table></td></tr>';

  // Seller + client.
  const cli = [ '<b>' + e(d.client) + '</b>' ];
  if (d.loc) cli.push(e(d.loc));
  if (d.tel) cli.push('Tél : ' + e(d.tel));
  if (d.nif) cli.push('NIF : ' + e(d.nif));
  if (d.stat) cli.push('STAT : ' + e(d.stat));
  h += '<tr><td style="vertical-align:bottom;padding-top:18px">' +
      FACT_CFG.SOCIETE.map(e).join('<br>') + '</td>' +
    '<td style="vertical-align:top;padding-top:18px">Adressé à' +
      '<div style="' + B + 'padding:6px 8px;margin-top:2px">' + cli.join('<br>') + '</div></td>' +
    '</tr></table>';

  // Lines.
  const TH = TD + B + 'text-align:center;font-weight:bold;';
  const COL = TD + 'border-left:1.5px solid #000;border-right:1.5px solid #000;';
  h += '<table style="width:100%;margin-top:36px">' +
    '<tr><th style="' + TH + 'width:46%">Désignation</th><th style="' + TH + 'width:17%">PU</th>' +
      '<th style="' + TH + 'width:14%">Qté/kg</th><th style="' + TH + 'width:23%">Montant Total</th></tr>';
  d.lines.forEach(function (l) {
    h += '<tr><td style="' + COL + '">' + e(l.label) + '</td>' +
      '<td style="' + COL + 'text-align:right">' + (l.pu == null ? '' : factNum(l.pu, 2)) + '</td>' +
      '<td style="' + COL + 'text-align:right">' + (l.qty == null ? '' : factQty(l.qty, l.isKg)) + '</td>' +
      '<td style="' + COL + 'text-align:right">' + factNum(l.montant, 2) + '</td></tr>';
  });
  // Blank space under the lines, closed by the bottom border.
  h += '<tr><td style="' + COL + 'border-bottom:1.5px solid #000;height:50px"></td>' +
    '<td style="' + COL + 'border-bottom:1.5px solid #000"></td>' +
    '<td style="' + COL + 'border-bottom:1.5px solid #000"></td>' +
    '<td style="' + COL + 'border-bottom:1.5px solid #000"></td></tr>' +
    '<tr><td></td><td colspan="2" style="' + TD + B + 'text-align:center;font-weight:bold">TOTAL à payer</td>' +
    '<td style="' + TD + B + 'text-align:right;font-weight:bold">' + factNum(d.total, 2) + '</td></tr>' +
    '</table>';

  h += '<p style="margin-top:20px">Arrêté la présente facture à la somme de : <b>' +
    e(cap) + ' ariary</b></p>';

  h += '<table style="margin-top:24px"><tr><td style="vertical-align:top;padding-right:8px;white-space:nowrap">' +
    'Conditions de règlement :</td><td>' + FACT_CFG.REGLEMENT.map(e).join('<br>') + '</td></tr></table>';

  return h + '</body></html>';
}

/* ---------- Drive ---------- */

/** The "Factures" folder: created once, then always the same folder. */
function factFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(FACT_CFG.PROP);
  if (!id) {
    const created = DriveApp.createFolder(FACT_CFG.FOLDER);
    props.setProperty(FACT_CFG.PROP, created.getId());
    return created;
  }
  var folder;
  try {
    folder = DriveApp.getFolderById(id);
  } catch (e) {
    folder = null;
  }
  if (!folder || folder.isTrashed()) {
    throw new Error("Dossier « " + FACT_CFG.FOLDER + " » introuvable (" + id + "). " +
      "Kim : tsaraentry -> Paramètres du projet -> supprimer la propriété " +
      FACT_CFG.PROP + ", puis lancer testFacture.");
  }
  return folder;
}

/**
 * Build the PDF of the invoice of one order and save it in "Factures".
 * Called by the "Imprimer" button of the Facture column.
 * Returns { url, name, facture }.
 */
function factPdf(orderNumber) {
  const d = factData(orderNumber);
  const name = (d.facture + " - " + d.client).replace(/[\\\/:*?"<>|]+/g, " ").trim() + ".pdf";
  const blob = HtmlService.createHtmlOutput(factHtml(d)).getAs(MimeType.PDF).setName(name);

  // Two clicks on one invoice at the same moment would trash each
  // other's file.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const folder = factFolder();
    const old = folder.getFilesByName(name);
    while (old.hasNext()) old.next().setTrashed(true);
    const file = folder.createFile(blob);
    return { url: file.getUrl(), name: name, facture: d.facture };
  } finally {
    lock.releaseLock();
  }
}

/** RUN FROM EDITOR: tsaraentry -> FactureServer.js -> testFacture
 *  Takes the most recent order that has an N° facture and is not
 *  cancelled. Writes ONLY the Factures folder, and creates it on the
 *  first run. Logs the lines, the total and the link. */
function testFacture() {
  const h = histList();
  const o = h.orders.filter(function (x) {
    return x.orderNumber && String(x.facture || "").trim() && !x.annule;
  })[0];
  if (!o) { Logger.log("Aucune commande avec un numéro de facture."); return; }
  const d = factData(o.orderNumber);
  Logger.log("Commande " + o.orderNumber + " -> facture " + d.numero + " du " + d.dateFacture);
  Logger.log("Client : " + d.client + " | loc=" + (d.loc || "-") + " | tel=" + (d.tel || "-") +
             " | NIF=" + (d.nif || "-") + " | STAT=" + (d.stat || "-"));
  d.lines.forEach(function (l) {
    Logger.log("  " + l.label + " | PU " + (l.pu == null ? "-" : l.pu) + " | Qté " +
               (l.qty == null ? "-" : l.qty) + " | " + l.montant);
  });
  Logger.log("Total : " + d.total + " Ar - " + d.enLettres + " ariary");
  const r = factPdf(o.orderNumber);
  Logger.log("Fichier : " + r.name + " -> " + r.url);
}

/** The farm logo, JPEG, taken from the paper invoice model FA26023. */
const FACT_LOGO_B64 = "/9j/4AAQSkZJRgABAQEA3ADcAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCACPAR8DASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD3eiiigAooooAaeKz59asIHKvcAsDghQTj8q0SK5rU/Djz3TT2rKN5yyt2PqK4cbUrwheirs3w8acpWqOyLL+KbNThUlcfTFU7nxUXjZYbcgkcFj0pkXhSYj99cIv+6M1m6rYJp06QrIXYjcxIr5/FYrMYU3KeiPTo0cJKajHVlaO6uIpPMSZ1fOSQetaS+Jb9UC/umx3KnP8AOs20gN1dxQD+NgPw712Mfh/TkABg3H1LHn9a5sto4usnKlKyN8ZUoU2lON2YP/CT3/8Adh/75P8AjW/pGrpqUZBGyZPvLn9RT10TTlORbL+JNSrHY2ILqkMXHJAAr3cNSxNGd6tRNHm1p0pxtCOpczRmsyXXtPjJHnhsf3eaZbeILK5uBCrMpb7pYYBrv+u0Obl5lc5vq9S17GvRTS2Bk9KzbrXrK1YoZN7jqFGa1q4inSV5uxEKcpuyRqUlYC+KrUtgxSgeuB/jV+01qzvGCRy4c9Fbg1jTx+Hm+WMtTSWGqxV2jRopM/nUNzdRWkJlmfag7muqUoxXM9jFJt2RPRVG31eyunCRTqWPQHgmm6hq1tp5VZiS7dFUVi8VSUefm0L9jO/LbU0KKyrLXbS9mEKFlkPQMOtamaulXhVjzQd0KdOUHaSClqhd6taWR2yyjd/dHJrPbxVaZ4ilI9cCsKmPw9N2lLUuOGqSV0jepayLbxBY3DBC5jY8AOMVqggjPatqOIp1leDuTOlKGkkOooorYzCiiigAooooAKKKKACiiigAooooATvRQTVa4vbe1UtNIqj3NROcYL3mNRbehYPSuE1m4+06rOwPCnYPw/8Ar12TXaHT2ulPybCwJHavP2YsxY9Scmvm+IK6dOMF1PWyun7zk+hr+GYPN1QyY4jUn8Tx/jXVX07WtnLOibyi521j+FYdtpNNjl3wPoK3pApQh8bSMHNduV0XTwemjZz4ypzV32RxFxrt9cE/vfLHogqtHbXl6QVSSXPc81v3Xhq3lLG1m2Mf4TyKzWGpaI45Pl/mteJWw9dT5q7bj5HoQq03G1JJPzIJtGvoIw7QEg8cckVc07w/cSyLLcAxRgg47n/CtXTfEMN0VinAjkPHsa072cW9lNMP4VJFd2Hy/CNe3jK9jmq4quv3cluc9r2sOHNnbsRt++w/lWFbwNc3EcKD5nIFRszOzOxyzHJJrofC1oHkkumH3flX+teXGU8fi1FvQ7ZKOFoXW4mr6baafpaBR++LDDevrXPoxR1deGBBBrrPEsEBtBPIT5i/KgzxmuT7Us1p+xxCUNAwMvaUm5anZ6JqsuoiQPEF8sD5gepNUPFc5zb24PBy5H6D+taXh+2+z6VGcfNJ85/Hp+lYXiZt2qgD+GMD9TXsYypUjly53qzgoRhLFe6tEUdKjD6pbg8Ybd+XP9KNVuzeahJL/CDtX6CoIZ3t2Zk+8VK59M1Zk0yaPTY7sqcMeR6Dsa8CnKpUoOnDXuenNRhW55fI0PDNiZblrtvux8L9TWnrurfYoxDCf3zjr/dFZFhrq2Gn+QsJMgyc9jWTPPJczNNKxLsea9F46GGwipUviZyxw8q1dznshjFpHLMSzMep6k10n9kW1nozy3K5l25yex7Cs/w/a/adTVmHyRDef6V1GrQQz6e/nkhEG7g9wKeX4TnoTrz1YsXX5akacdEcFiuo0DV5ZpIrJ48hV+/noB/nFcvXUeFbbEc1yw5Y7VP865solU+tcsXob49Q9jdo6WikpksqQxtJIcKoyT6CvuG0ldnztrvQfS1BBdQ3CBopFYH0NTZ4ojOMldMHFrcWikB4paoQUUUUAFFGaQmk2lqwCq91ewWib5pAorL1XX47TdDDh5f0FcuzXWpXPVpZG7eleNjc2jTfs6Wsjvw+Cc1zT0RrX/iWWXclquxT/EetZsNpe6lLkB5M/wATdBW7p/hyOIebeEMf7vYVLB4k0x9di0WzZZZyrM/l/dQAZ5/lXLSwOIxMlUxErLsXVxlCh7lND9dYWmhiFeN2Ixj0/wD1CuOzXReK5908NuDwoLH8eK56vIzeXNiVBbI9HARtRcurO70aDyNKgTuV3H8eawvEd/Mbv7LG5VEGWwcZJre0lp306NrgAMRwAOg7Vma/pD3Dfa7cbmAwy+uO9e9i4VHgUqPY8yhKCxD5zK0JbmbU49kjhEO5+e3pXZTwR3ELRyKGUjoa4/R9UXS2kSaI/Mc57itK58TxeURAhLkcZFc2BxNGlhnGq9fM2xNKpUrXgtDm7mLyLqWMH7jEA/StuK+kvPDlzG5zJFgE+ozWCztI7O3LMSSfU10/h7Tz9hnaZcLNwAfSvMy9VKlacafwu514vlhTi5bo5eu20CNY9Ihx/Flj+dcvqOlT2EzAqWjJ+Vh6VJba3c2tkbZFGRkBvSqwEvqVeTqoWKX1imuRlnxLe+feLbocpF19zWKo3MF9Tip7ezub6chEZixyWP8AOlvbKawuDG6njkMOhrlxHtq9X28ou1zel7OnD2SetjvYQsduijgKorhdVuRd6nNKv3d2F+gq4/iC6ltPIWPDlcFhWYbS4VdxhfHriu/MsW8RSjTpJ2W5y4Oj7KblNi2lubq8igH8bAH6d69BWJFhEe0bQMYNcl4bhJ1Quyn5EJ5FdiK9HIsNy0XKS3OXMq3NUSXQoz6XaPG+IEDEHkCuCIKsVPUGvSiMjFcdrWjywXDzwoWic54/hrLO8DeCnTjsXl2ItJxm9y94UQeTcSdywGfwo8T3+2JbND8z8v8AT0rJ03VZdMWRBHuDEHnjBqq/2jULsvsZ5HPQCuP641g1Qpp8xv8AV74h1JvQr13OhoE0eD3GfzrlL7SrmyRGkXIYckdjVmz164tbMW4jDEDCk9qyyyf1Oq3WWpeMX1iC9mzZvfECWmofZvL3KMBmB6Zq5q1tNe6e0UDBS3XPcelc5pel3GoXgubgER7tzEj7xrrLi5htLdpp5FjjQZZmOABX0GElVxUJe10T2PLxHs6DXLutzhXivdMk5EkR9R0rYsfEzKQl2vH99aTSvGWheIppLNZFDhiFWUY3j1FP1Hw2MGWzP1Q1wVMHicI+ehK67G9LGYfEq09zobe5iuYw8Thge4NTZrz2G5utNnOwsjA8qe9dZpetw3yhG+SYdVPf6V24LNY1nyVNJGWIwUqfvR1Rr0UgOaWvYTucI0niud1zWzGTa2rfN0dx29qu65qP2G0wh/eycL7Vx0MUl1cLGuWeQ9T/ADr5/NsfJP2FLdnqYLDJr2s9kS2NjNqNxsTJ5yzHtXXwW1nolk0shVFRcvI1TWFjFp1qEXGQMsx71498QPGUmsXr6bZyYsYWwxU/6xh/St8vy+OHhzz1kzgzTM7Ky2JfGHxEuNUeSy0t2htASDKOGk/wFc34Z8Qy+HNYF+sQm3KUcE8kE81iilr0OZ7nyE8TUlPnbPaIvHHhfWUVrsmGY8HeOR+NTLe+FI5Vl/tKMqOdpbrXiFXdM0m91i7W2soGlcnsOB9axnh6VSXNKN2enRzjEpckT2bVfiLoVnYzfZ5/PlCkIiDqccVB8PfFk+tWYsr2NzcQrgTY4cD1PrVTw78LrS0Cz6q/ny/88l+6P8a761sbaxhEVtAkSDsgxXZGLOqiq85KctCK50u0uwTLEMnuODWe3hi0JyHcD0rd6UZrCpgaFR3lE9OOIqQ0UjKttAsrdw2ze3+1WooCjAGB7UtcVrfxDg0TVJLCaxnLqcKcYDe4rSnQpUVaCsY18S7XqM7NkVxhgCPQ1AbC1ZsmBM/SvLtQ+LN0FdbXTxG3RWkNc3P8QvEVw2Dd7AT0RaU405O7VzheZwhpFs95SOKEYRVX6VBdT2CR7rqWEIO7kV4S+r+Jr18R3d0yE8E/LVabTbsrnUNTjRScEPKWOfpRaFrW0Mnmcr3ij2GfxV4YspggnhZ/RBmnweOdAuJFjFzgsONy14qttpcLZfUZGYdDGlTmXTmyUmvD7hKmMIR2Rm8yrN6nvtleWF2vmWksMgPBKEVdr5zW6FlMZ7G5vIX6jCkCum0P4pX1myw6kguYh/GOGFaxlGOiRtDMYt++ez00gHqOK53R/G+iay4jguQkxH+rk4NdGCCM9qvRo7qdWM9YsrtY2zkloUJ+lOjt4YR+7jVfoKmoxUKhTTukbOpJq1zhPEHjLUdKulEmhSvaD77Nzke1VYviP4Y273tJUk7r5fSvQpIklQo6KynqCM5rj9f+HWk6sjSW6fZLjruQcH6ipnRjJ3aucVR4iF3TkZV78WdNijK2NnLI2ONw2gVwPiHxjqniJtk8nlW4ziJOn4+tQ+IPCmpeHZ9t1EWhP3ZVGVNYg5qW2tDxcRia8naYqM0bh0YqynIIPIr0rwZ8R5IXj0/Wn3R/dS4PUfWvNaMUk2jChiJ0pXR9JX2m22q26yIV3MMrIveuOubeewudjgq69CKxvhx4ye3uU0XUJcwPxA7H7p/u/SvTNW01NQtWGAJFGVavMzHLo1o+1paSR9nlmZKUUpbFTRNaF0ot5yBMBwf71bo6V5tmS2n4ykiH8jXdaTfi/skk4Djhh70sozB1P3NT4kdWNwqh+8hszlNcujc6nIM5WP5VH860vC1mGMl0w6fKv9a56Zi08jHqWP8AOu08PKF0eHHfJP515WWf7TjXUkdeM/c4dRRi/ETW20bwxIsLbZ7pvJQjsO5/Lj8a8Hr0/wCMUzfadKg/hCyPj3yteY19XU3sfAZhUcqtuwUUUGszzzR0HRLnxBq0VjbDluXY9FXua9+0Hw9Y+H7BLa0jGf45COWNch8JNNii0a51EqDNNMYwfRVxx+ZNejV0U42Vz6DAYeMYKb3YUtFFaHpHOeONSm0nwneXVvIY5xtVGHYlgK8usvH2vCEs17ll7Mowa7f4q3NrH4chtp5HV5ZtyKn8W0Hr7civH4Ii8ZDPgbSQCcCsZydzxMbWnGraLPRbD4qz2s6w6nbLKmB+9i6/lWvrF14b8c6YYoruOO9UEws52kNjp9K8hQRHSbhmP74TRBc9du18/wDstVwSrZUkEdwannfU5/rk7cs9UXNStLvTp2s7ojcDnAbP41LYXBj+a3ihVwOXl5x71TjaJ2LXDSMccYNWVvbaFAIbJC4/jkOTUo5bq9zQA+0ENe6pKxxjZCpP4Va/syzhGTYykHo9zIF/SsNtWuznZII89kUCq0k80xzLI7/7xzTuX7WKOnWfTrbJIsoxjlVUyHNU7rXtsRjtJAMntCFrBpKXMS6rexoNreoMCDPn8BVSedpypdUBAx8q4qKnRr5kqJ/eYCluRdtiKxjcMrFWHQg4IrufDPxKvtK8u21HddWw43fxqP61y2v2Tadr15aMhRY5SqAjGV/hP5Yqq1syWMV3/C8jR/8AfIU/+zU02jWnUqUpaM+j9K1yx1m3E1pMGB6qeCPwrR4r51vfFF1NcRz2o+ySqckxHGeP5V6l8O/FV14htLiG+ZWuICMMByy+preM76HtYfGqo+V7nc0GkFLVnoFa8sre/tXt7mJZInGCrCvC/G/hCTw1fiSEFrGY/u2/un+6a99xXP8AjTTo9T8KahE6gskRlQ+jLyKicU0cWMw8akG+qPnelpKWuY+aegKzI6uhIZTkEdjX0P4O1k674ZtLyQjzgPLl/wB4cH8+v4188HpXr3wfnZtG1CA/dScMPxX/AOtWlN9D0stqONTlNnxNZiG8SdRhZBz9ai8O3htr1o8/JIp4PqK2PFCg6ch7q4x+tcnDIY5ldeor5HHv6pjeeJ+gYVe3w3Kx95EYb2aMjo5/Kur8MziTTPLzzGxH9azPE1l5dwt2g+V/lbHrVXQb/wCx321z+7l+U+3vV4b/AGLHcstmTV/2jDJrdGP8YLJmtdNvQPljd4mP1wR/I15RX0h4h0eLXtDubCTH7xco391hyDXzreWc+n3strcoUmiYqyntX1c9dUfBZlRcanN0IaDRRWZ5h6v8JNYjNpdaRI2JVczR57qQAfyI/WvT818xadqFxpWoRXlrIUmibII7+30r3Xwp4ysvEdqi7livAPniJ6+49a3pyVrHu4DExceR7nUUtIDmlrU9U8x+MULNY6XPj5UldD9SAf8A2U15pGd0Qhgt5Zzt6le/4V9H3un2uoxLFdwJMituAYZwcYz+tMt9KsbXAgtIU+iCs5Qu7nm18C6tTmueIaR8PNc1Mqzw/ZoieWl/wrtYvhVpEFrm6u5SwHzOW2ivQZn8iFpAjNtGQq9TXkviW68R6/qDRvbXNnYrwsYQnd9cUnFRRE8NSox1V2cxrOl6LBdSRaZfyTBDtJK5H4GsWa1kgUMcMh6Mp4rrF8M3XlgRW0jFRwPJIP55pkfgbxLqC4+xrEmf+WjYOKzabPOlQlJ6ROPor0RPhHqjRgtfQKx7bTVG7+F2vQN+6MMyj+INijkZDwlZa2MbRPCOra8y/Z7cpCTzK4wK1vEPhm28OJHbtY3V1KQC1wvCEnsMelUoJtes3MAvbiFE4JVGKitRdd1Q4STxEJRjhTHmmkjWMKajZrU4mRfmysTIPQ1HkqwYcEc10cuo6xqF19mtlaaQkjCxDketa9h8L9bvv3l3JFbAjODyaXK+hmsPOT9xE41TQvGNjFFrEn2LVI12C5A4bHc1Y0rwVcz2F1pEskFxaysJoLmJgTG+MZI9CK2dN+E+nQFXvbiScjqo4BruNO0my0uIR2cCxLjHFaKN9z06OElLWojy3QvhrqqzTJePFDATtLY3MQPT0r0nQvD1h4ftPIsYgpP33PVjWtS1aikdlLDU6WyCiijNUdIlcv491eLSfCl3ub97cIYYl7knjP4CtXWtdsdCs2uL2ZUGCVTux9AK8H8U+JrnxNqZuJcpAnyxRZ4Uf41nOSSODGYmMIOK3MOlpKWuc+cYGvaPhNZG38MTXLDH2iclT6gAD/GvJdH0q41rVYLC2Ul5GAJ/ujuTX0Xp9lBpOlwWkICwwRhR+HetIaLmZ62WUHKfMZPiqcLbwwZ+Zm3fgKwNPgNzepHjjBJ/KpNWvft1+8g+4vyr9K2PDFjtR7uQcsNq59K+SqR+v4522R93B/VsNrub91bR3du8MgyrCuDv7KSwuTFIDj+FvUV6HiqeoafDfwGOQc9m7ivfzPL1iY80fiR5uExTouz2MfQtaEiC1uWww4Vj39qyfHngldehN/YqFv415H/PQDt9ahvtPuNOmw4O3PyuO9auleITEBDdnK9n9PrXBgcxdN+wxOjReOy+GIhz09UzwaWGS3meGaNo5FOGVhgim17v4m8Gab4pg+0QlYbvHyzoPvex9a8c1rw7qWgT+XfW5VScLIvKt+Ne3ZNXWx8VisFOjLbQy/pT4Lia2mWaCRo5FOQynBFR59KKRxptO6PR/DvxTubULBq8ZmjHAlX7w+vrXpWleJNK1mJXs7yNif4GOGH4V83cVNatILmMRS+UzMBv3YA+taRqNHoUMwqQ0lqfUNFeS6PrGrLdGz0XWxqDR5LRXKYGB1Ib0z/OtqD4hXdqgk1PSWMAOGuLVw6rx14rVSR60MXBrU9ApCo9BXL2nxC8O3fAvfLPpIuK27bWtMvEDQX9u4PQCQZ/KndM2VWnLZl3YB2FLikDAjg/rS54oLTiLSEZozRmmPQytc06e+0qe1smihmmXYZGXO0Hrj3rjNK+FFvbOJLy+kkbuqDAr0jNJUuKbMZ0Kc3dmdpehafpEWy0tkQ92xyfxrRxS5pM01Y0ioxWgtLVSfUbO2z593BHjqHkArHuPG+gQcfb1lbssQLE0XQnVgt2dFRXnOpfFrT4W22FpJOcfec7Rn6VlaZ461TxLrUenGX7HFKr7fIxuJAJAyfpilzo55Y2knZO7PT73U7LTojJd3UUKjrvbH6V5/r/AMVbeJHh0eIyyHjzZBwPoK4vVNGvdRL3dhfPqkIG5lJ/ep16r/hXMFSrEMCpBwQRyKzlNnn4jHVNoqxe1PVr3WLpri9naVz0yeB9KpUcUmaybueXKUpO7Fqazs7jULqO1tYmklkIVVUVq6B4U1TxFKotoSsGcNO4woHf617J4f8ADGleErMuNrTn7879T7D0o0S5paI7cLgalaW2hD4L8IQ+GbHzZ9r30o/eP/dH90U7XtaEu60t2+Xo7Dv7VFquvvc5htyUj7t3aqWm6VNqEgIBWIHlzXh43HyxD9hhvvPtcFgYYaCnU0E0vTZNQuQgBEY+839K7qGJYIljQYCjAFRWdnFZwLFEuAP1q1Xp5bgFhoa/EzDF4l1pabBSGlor1DkIJ7eK5iMcqBlPYiuX1Lw3JETJafOnXZ3FdfSEZrhxWApYle8tToo4mdJ6M8/tdQu9OlwjMuDyjVvLqum6xAbW/hTDcFZBlTWre6Xa3y/vYxnsw6iucvfDVxDlrZvMT0PWvH9hjcE/c96J2ylhsUrTVmYOufCu2ut1xolwIiefKflT9D2rzrVPDuq6LKyXtnIgU/6wDKn3zXqkV3faZLt3PGf7rdDWrD4jjmXy762V1PUgZ/Q10Us1pSdqi5WeRisg5vepHgVT29ndXjqttbySksFGxSeT0FeyX3hDwnrjF4QtrM3JMR2fp0rn7j4a63pU3n6HqYfnPUow/EcH9K9GE4T1i7nh1MsrUpe8tDF1SaPwlo7aJaOrancqDfTp/AD/AMswf51zdvqt5babc2EUpWC5K+YPXHQfrWtr+h+JTctdalp0rP0aWOMNu9yV/ma51gVbawKkdQR0rR36HLV509rI3PB+nNqfiOG3KI0GC05dchYx1Pt6Z96i1y/sZb+VdLtVtoEkIR0Y5dffn8a2rC603SfCmqRWGoJLqN5hCzIYysWOQOuSeR+PtXG0N2VglLkgktze8O3OoXuqQWCaldQtO4RCjE8n1rS1TxDq2hX81na6/c3EkMhjk3DABHXHPPNP+HunSrrT6vLby/Z7O2kmRthw7YxgHv1P5VzAtb7UNT8vyZWu7mQnay4LMee9PWxfPNU009WbifEHxKgx/aDH6gGnj4i+JRj/AE3J91Fcqw2sVOMg44rovCGlw3d/JqF8ypp9gBLKz9GbOFX8TSTk2TTrVZSsmdFqWu+NdN0lNQnvY9hKh0QDdEWGRuGOOKwj8Q/ErDH288+iitbR7yGTWr46tq1lNaamCk6Bm4J+6RleMHj2FclrekT6Fq89hPyYz8rf31PRqcm7XNatSolzRZ1Gia9rPiK+i0+bX57eaViEAXg4GeuaqXGqwLdSWd9qur7kcxyNvGAQcdMe1cxZXMlnqFtdxE+ZDKrrj1BBrvPGMKWetDUrXQVnjuIlummZHZVJGTuHQdM8007oIVZSp3bOW8S6Tc6JqQtprprhJI1ljkyeVNbXhDU0v2uNEuPKge6i2W1wqAFJB0/OuX1PU7vV71ru8l8yUgD0AA6ADsKrwMFuIm8xogHH7xRyvPUVN9TBVf3l1sOu7WaxvJrW4UrLCxRh7g1PpF+dL1izvgCRBMrkDuAeR+VbfibV9L166Wa2s7h70xqjTDCiRgPvFRnn8araX4R8RXksU1rp0ybWDK8o2AEfWi2tx+ylz+6rm7rvleFtdu7qwsJXdpBIlxIx8sbgG+UDGeuOc1nXGv6Rqull9Ys/M1LecSW6iM4z3Pf8q6W3+Gmq6jJ5+vatjnlVJYkfU9K3rHw14T0Bg3lpczr/ABSHef8AAVNStCGsnY9KngsRWekdDyzTfC+ra7dH7BYSJAzcPJwqj6nrXouh/DHTtOxcaxMLiReQnRB/jW3P4l2rss4FjUdCR/Sst5b7U5SMySn0HQV5lXNqcXy0lzM9bC5Coe9VN2fXbOwhFtYRLhRgBRhRWBNdXmpS4YtIx6KOgrVsvDEsmGun2L/dXrXRWmn21km2GMD37mudYbGY13qvlj2PVVTD4ZWpq7MHTfDTMVlvDx18sf1rpYokiQIihVHAAFSAUtezhcDSw6tBHDWxE6rvJiUtFFdpgFFFFABRRRQAUmKWigCCa1hnXbLGrD3FZF14YtZSWhZoj6DkVvUlctbB0aq9+JrTr1IfCzi7jw5ewndFtlHbacGqwl1OwOMzxgdjnFd5ikKKw5AP1FebPJYp3pSaOuOPk9JxucdB4lvEGJVjlHfIwaWXUdIvv+P7SYpMjBbaG/nXSz6TY3H37dCfUDBqjL4YsnOVLp9G/wAay+q4+l8EroHLB1fihY5qXw34JvgQbQ27noVZlx+XFZsnwz8PzDFrrUyN23lT/QV1UvhMhiYrrj0Zaqv4avE+68TD6kVLxWNp/HC5jLL8DU2djno/htewKUsPEuxOyjI/kaQ+BvF1s4a016Nz6tIykfoa2pNFvoxnapx6OKiFtfxcAOv0kH+NS83nD44Ef2JQl8Mjjn+FPiJfuvZt9JT/AIVqReEPG0Nh9hjksBbcZj+XBIGMn5eT71uGbUI+s8w/7af/AF6ab29H/L1P/wB/DU/27TW8RLhyN7xkcpb/AA08T2s4mi+xiQZwS+R+RWr994G8ZayoF/eWT4wBuc5+nC1t/bb08fapv+/hp4kv5Ok0xPvJ/wDXoWeweiiH+rkUrORzdt8J9YR1kk1G1hdTlSm5iD+QrQk+GlzPgX/iTeBxg8/zNa32O/m6q7fVx/jUseh38oB2oP8AeeqWbVJfBTBZHhofFIx4/hr4cg/4+dYmdh1Cso/oa0o9A8E2OPLsPPYf3izfzNXo/DF4+N0kSj6k1ai8JgHMt0Svoq4NWsTjqnwQSNI5fgae+pWi1bTLLP2HSoY+2QoFRT+JL2TiMpEP9kc/rW5F4ZsYz84d/q3+FXYNMs7fmO3jBHfGTT+qY+r8c7I2U8JT+CFzjj/aV+ek8gP1xVu38M3kuGlZIh6dTXYhFXoAKWtaeTQbvVk2EsfK1oKxiWvhq0hOZcyt/tdK1ooIoUCxoqj0AqWlr06OEo0laETknXnP4mJRS0V0IyCiiimAUUUUAf/Z";
