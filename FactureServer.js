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
 * CLIENT BLOCK: the name from R. Adresse (M), téléphone (B), NIF (K)
 * and STAT (L) from the CRM row with the same canonical name.
 * Lieu de livraison (C) is NOT printed: it is where the fish go,
 * the Adresse is the invoice address (Kim, 2026-09-11).
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
  const info = { tel: "", adresse: "", nif: "", stat: "" };
  const crm = crmEntrySheet();
  const crmLast = crm.getLastRow();
  if (client && crmLast >= CRM_START) {
    const cv = crm.getRange(CRM_START, 1, crmLast - CRM_START + 1, CRM_COLS).getValues();
    const canon = crmCanonName(client);
    for (var k = 0; k < cv.length; k++) {
      if (crmCanonName(cv[k][0]) !== canon) continue;
      const s = function (col) { return String(cv[k][col - 1] == null ? "" : cv[k][col - 1]).trim(); };
      info.tel = s(CRM_COL_TEL);
      info.adresse = s(CRM_COL_ADR);
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
    tel: info.tel, adresse: info.adresse, nif: info.nif, stat: info.stat,
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
    'body { font-family: Arial, Helvetica, sans-serif; font-size: 12pt; color: #000; }' +
    'table { border-collapse: collapse; font-size: inherit; }' +  // quirks mode: see patch 20260911d
    '</style></head><body>';

  // Header: logo + invoice box.
  h += '<table style="width:100%"><tr>' +
    '<td style="width:52%;vertical-align:top"><img src="data:image/jpeg;base64,' + FACT_LOGO_B64 +
      '" style="width:147px;height:120px"></td>' +
    '<td style="vertical-align:top"><table>' +
      '<tr><td style="' + TD + 'font-weight:bold;white-space:nowrap">N° FACTURE :</td>' +
        '<td style="' + TD + B + 'font-weight:bold;font-size:12pt">' + e(d.numero) + '</td></tr>' +
      '<tr><td style="' + TD + 'vertical-align:top;white-space:nowrap">N° Commande :</td><td style="' + TD + 'white-space:nowrap">' +
        d.commandes.map(e).join('<br>') + '</td></tr>' +
      '<tr><td style="' + TD + 'white-space:nowrap">Date facture :</td><td style="' + TD + '">' + e(d.dateFacture) + '</td></tr>' +
    '</table></td></tr>';

  // Seller + client.
  const cli = [ '<b>' + e(d.client) + '</b>' ];
  if (d.adresse) cli.push(e(d.adresse));
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
  Logger.log("Client : " + d.client + " | adresse=" + (d.adresse || "-") + " | tel=" + (d.tel || "-") +
             " | NIF=" + (d.nif || "-") + " | STAT=" + (d.stat || "-"));
  d.lines.forEach(function (l) {
    Logger.log("  " + l.label + " | PU " + (l.pu == null ? "-" : l.pu) + " | Qté " +
               (l.qty == null ? "-" : l.qty) + " | " + l.montant);
  });
  Logger.log("Total : " + d.total + " Ar - " + d.enLettres + " ariary");
  const r = factPdf(o.orderNumber);
  Logger.log("Fichier : " + r.name + " -> " + r.url);
}

/** The farm logo, JPEG, 360 x 293 px: Kim's logo file (589 x 480),
 *  resized. Printed at 147 x 120 px, the same proportions. The logo
 *  first taken from the Excel model FA26023 was squashed (287 x 143). */
const FACT_LOGO_B64 = "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";
