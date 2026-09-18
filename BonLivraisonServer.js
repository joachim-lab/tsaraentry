/***************************************************************
 * BonLivraisonServer.js - TSARA Entry web app
 * BON DE LIVRAISON EXTERNE (Kim, 2026-09-16): the button
 * "Imprimer le bon de livraison externe" on the order card of
 * Commandes > Paiement & livraison.
 *
 * DECISIONS (Kim, 2026-09-16):
 *   - GOODS ONLY. Lot, product, PM, kg, number, delivery mode and km.
 *     No price: the facture carries the money. A price-only edit is
 *     therefore NOT an amendment.
 *   - ONE ORDER = ONE BL. The BL number is column D, typed on the card.
 *   - KEEP ALL. Every PDF is saved in the Drive folder "Bons de
 *     livraison" (inside Tsara Tilapia). A reprint puts the old file of
 *     the same name in the bin, so there is one file per BL.
 *   - RIGHTS: the grid column "Bouton BL" (AccessServer.js).
 *
 * THE ORIGINAL ORDER. Tab "2026" is overwritten in place by every edit
 * (cmdModifyOrder, cmdUpdateDelivery). So the original is COPIED at
 * the moment of delivery: the save that records the FIRST Date
 * livraison (V empty -> set, cmdRecordFulfilmentPrbBody) calls
 * blSnapshot, which appends one row per order row to the tab
 * "BL origine" of the Commandes file.
 *   - A snapshot row matches an order row on: sheet row + N° commande
 *     + client. A cleared test row reused by a new order cannot match.
 *   - Written ONCE per order row. If V is later cleared by hand and set
 *     again, the first snapshot stays: it is the first delivery.
 *   - A failed snapshot never undoes the delivery: the save returns
 *     blWarn and the screen shows it beside the invoice line.
 *   - Orders delivered before 2026-09-16 have no snapshot. Their BL
 *     prints the current state and says so. They cannot be rebuilt:
 *     price edits leave no trace, and the Y "Ajusté" note holds a
 *     deduction count, not F / L / PM.
 *
 * AMENDED = any difference between snapshot and current rows on
 * lot (A), alevins (F), PM alevins (G), à livrer (H), kg (L),
 * PM poisson (M), livraison (AB) or km (AC). N is a formula of L and M
 * and is not compared. Amended -> the PDF prints the original table,
 * then the amended table with the changed cells highlighted.
 *
 * NUMBER COLUMN: alevins print H ("à livrer", the fish handed over,
 * +5 % included); H empty (old rows) -> F. Grossis print N.
 *
 * PRINTING, the Stock poisson way: a throw-away Sheet in "Rapports
 * Commande", a Sheets PDF export, the Sheet to the bin, the base64 to
 * the browser, which prints from a hidden iframe.
 *
 * REFUSED, NEVER GUESSED: cancelled order (AA), no Date livraison (V),
 * no N° BL (D), rows of two different orders.
 *
 * ONE-OFF SETUP: tsaraentry -> BonLivraisonServer.js -> blSetup
 * creates the tab "BL origine". Until then the delivery save still
 * works and says the snapshot was not saved.
 ***************************************************************/

const BL_CFG = {
  FOLDER_ID: "12ptgD7ifqWM5r4Ck7KmobV2RXiYZdQ44",   // Drive: Tsara Tilapia > "Bons de livraison"
  FOLDER_NAME: "Bons de livraison",
  SNAP_SHEET: "BL origine",
  SNAP_HEADERS: ["Ligne", "N° commande", "Client", "N° BL", "Date livraison",
                 "Enregistré le", "Lot", "Alevins", "PM alevins (g)", "À livrer",
                 "Kg", "PM poisson (g)", "Nombre poissons", "Livraison", "Km"],
  LOCK_MS: 10000,
  TITLE: "BON DE LIVRAISON EXTERNE",
  // Same codes and words as LIV_LABELS in CommandesIndex.html.
  LIV_LABELS: { enlevement: "Enlèvement à la ferme",
                environs: "Livraison environs de la ferme",
                ambohim: "Livraison Ambohimangakely" },
  TABLE_HEADERS: ["Lot", "Produit", "PM (g)", "Kg", "Nombre"],
  COL_WIDTHS: [130, 150, 90, 90, 110],
  CHANGED_BG: "#fff2cc",
  HEAD_BG: "#e8eaed"
};

/* ---------- goods of one row ---------- */

/** Goods of one row of tab "2026". v = values A..AC. */
function blGoods(v) {
  const C = CMD_CFG.COL;
  return {
    lot: String(v[C.LOT - 1] == null ? "" : v[C.LOT - 1]).trim(),
    alevins: cmdToNum(v[C.ALEVINS_NB - 1]),
    pmAl: cmdToNum(v[C.ALEVINS_PM - 1]),
    livrer: cmdToNum(v[C.ALEVINS_LIVRER - 1]),
    kg: cmdToNum(v[C.POISSON_KG - 1]),
    pmGr: cmdToNum(v[C.POISSON_PM - 1]),
    nombre: cmdToNum(v[C.POISSON_NB - 1]),
    livraison: String(v[C.LIVRAISON - 1] == null ? "" : v[C.LIVRAISON - 1]).trim(),
    km: cmdToNum(v[C.KM - 1])
  };
}

/** Goods of one row of tab "BL origine" (SNAP_HEADERS order). */
function blGoodsFromSnap(s) {
  return {
    lot: String(s[6] == null ? "" : s[6]).trim(),
    alevins: cmdToNum(s[7]), pmAl: cmdToNum(s[8]), livrer: cmdToNum(s[9]),
    kg: cmdToNum(s[10]), pmGr: cmdToNum(s[11]), nombre: cmdToNum(s[12]),
    livraison: String(s[13] == null ? "" : s[13]).trim(),
    km: cmdToNum(s[14])
  };
}

/** Null-safe equality: numbers to 1e-9, strings exact. */
function blSame(a, b) {
  const na = (a === "" || a == null) ? null : a;
  const nb = (b === "" || b == null) ? null : b;
  if (na === null || nb === null) return na === nb;
  if (typeof na === "number" && typeof nb === "number") return Math.abs(na - nb) < 1e-9;
  return String(na) === String(nb);
}

/** Every data row of tab "BL origine". */
function blSnapRows(snap) {
  const last = snap.getLastRow();
  if (last < 2) return [];
  return snap.getRange(2, 1, last - 1, BL_CFG.SNAP_HEADERS.length).getValues();
}

/* ---------- snapshot at delivery ---------- */

/**
 * Called by cmdRecordFulfilmentPrbBody on the save that records the
 * first Date livraison, AFTER its flush. Appends one row per order row
 * not yet in "BL origine". Throws a short French reason; the caller
 * turns it into blWarn.
 * @return {number} rows written
 */
function blSnapshot(sh, targets) {
  const C = CMD_CFG.COL;
  const snap = sh.getParent().getSheetByName(BL_CFG.SNAP_SHEET);
  if (!snap) {
    throw new Error("onglet « " + BL_CFG.SNAP_SHEET + " » absent " +
      "(Kim : tsaraentry -> BonLivraisonServer.js -> blSetup)");
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(BL_CFG.LOCK_MS)) throw new Error("une autre sauvegarde est en cours");
  try {
    const existing = blSnapRows(snap);
    const at = new Date();
    const out = [];
    targets.slice().sort(function (a, b) { return a - b; }).forEach(function (r) {
      const rng = sh.getRange(r, 1, 1, C.KM);
      const v = rng.getValues()[0];
      const dv = rng.getDisplayValues()[0];
      const no = String(dv[C.ORDER_NO - 1] || "").trim();
      const client = String(dv[C.CLIENT - 1] || "").trim().replace(/\s+/g, " ");
      const seen = existing.some(function (s) {
        return Number(s[0]) === r && String(s[1] || "").trim() === no &&
               crmCanonName(s[2]) === crmCanonName(client);
      });
      if (seen) return;
      const g = blGoods(v);
      const cell = function (x) { return x == null ? "" : x; };
      out.push([r, no, client, String(dv[C.BL - 1] || "").trim(),
                String(dv[C.DATE_LIVRAISON - 1] || "").trim(), at,
                g.lot, cell(g.alevins), cell(g.pmAl), cell(g.livrer),
                cell(g.kg), cell(g.pmGr), cell(g.nombre), g.livraison, cell(g.km)]);
    });
    if (!out.length) return 0;
    const first = snap.getLastRow() + 1;
    const n = BL_CFG.SNAP_HEADERS.length;
    // B, D, E and G as TEXT: an order number, BL number, date string or
    // lot name must not be re-parsed by the sheet.
    [2, 4, 5, 7].forEach(function (col) {
      snap.getRange(first, col, out.length, 1).setNumberFormat("@");
    });
    snap.getRange(first, 6, out.length, 1).setNumberFormat("dd/MM/yyyy HH:mm");
    snap.getRange(first, 1, out.length, n).setValues(out);
    SpreadsheetApp.flush();
    return out.length;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- data of one BL ---------- */

/** The table lines of a goods list. showChanges: mark changed cells. */
function blLines(list, showChanges) {
  const lines = [];
  var kg = 0, nb = 0;
  list.forEach(function (g) {
    const ch = showChanges ? (g.changed || {}) : {};
    const al = g.livrer != null ? g.livrer : g.alevins;
    if ((al || 0) > 0) {
      lines.push({
        cells: [g.lot, "Alevins", g.pmAl == null ? "" : g.pmAl, "", al],
        changed: [ch.lot ? 0 : -1, ch.pmAl ? 2 : -1,
                  (ch.livrer || ch.alevins) ? 4 : -1].filter(function (i) { return i >= 0; })
      });
      nb += al;
    }
    if ((g.kg || 0) > 0) {
      lines.push({
        cells: [g.lot, "Poisson grossi", g.pmGr == null ? "" : g.pmGr, g.kg,
                g.nombre == null ? "" : g.nombre],
        changed: [ch.lot ? 0 : -1, ch.pmGr ? 2 : -1, ch.kg ? 3 : -1,
                  (ch.kg || ch.pmGr) ? 4 : -1].filter(function (i) { return i >= 0; })
      });
      kg += g.kg;
      nb += g.nombre || 0;
    }
  });
  return { lines: lines, kg: kg, nombre: nb };
}

/**
 * Everything the BL prints, read NOW from the sheet.
 * rows: the order rows the card holds (cmdFindOrders).
 */
function blData(rows) {
  const sh = cmdSheet();
  const C = CMD_CFG.COL;
  const ss = sh.getParent();
  const tz = ss.getSpreadsheetTimeZone();
  const lastRow = findNextCommandeRow(sh) - 1;
  const targets = (rows || []).map(Number).filter(function (r) {
    return isFinite(r) && r >= CMD_CFG.START_ROW && r <= lastRow;
  }).sort(function (a, b) { return a - b; });
  if (!targets.length) throw new Error("Aucune ligne de commande valide.");

  const fmtD = function (d) { return d ? Utilities.formatDate(d, tz, "dd/MM/yyyy") : ""; };
  const cur = [];
  var orderNo = null, client = null, bl = "", dateCmd = "", dateLiv = "",
      tel = "", remarques = "";

  targets.forEach(function (r) {
    const rng = sh.getRange(r, 1, 1, C.KM);
    const v = rng.getValues()[0];
    const dv = rng.getDisplayValues()[0];
    const t = function (col) { return String(dv[col - 1] || "").trim(); };
    if (t(C.ANNULE)) throw new Error("Commande annulée : pas de bon de livraison.");
    const no = t(C.ORDER_NO);
    const cl = t(C.CLIENT).replace(/\s+/g, " ");
    if (orderNo === null) {
      orderNo = no; client = cl;
    } else if (no !== orderNo || crmCanonName(cl) !== crmCanonName(client)) {
      throw new Error("Les lignes reçues appartiennent à plusieurs commandes. Relancer la recherche.");
    }
    if (!bl) bl = t(C.BL);
    if (!dateCmd) { const d = factDate(v[C.DATE_CMD - 1]); dateCmd = d ? fmtD(d) : t(C.DATE_CMD); }
    if (!dateLiv) { const d = factDate(v[C.DATE_LIVRAISON - 1]); dateLiv = d ? fmtD(d) : t(C.DATE_LIVRAISON); }
    if (!tel) tel = t(C.CONTACT);
    if (!remarques) remarques = t(C.REMARQUES);
    const g = blGoods(v);
    g.row = r;
    cur.push(g);
  });

  if (!dateLiv) throw new Error("Commande non livrée : pas de bon de livraison.");
  if (!bl) {
    throw new Error("Pas de numéro de bon de livraison externe. Le saisir dans le champ " +
                    "« Bon de livraison externe », puis Enregistrer.");
  }

  // Lieu de livraison: CRM column C, same canonical name as the facture.
  var lieu = "";
  const crm = crmEntrySheet();
  const crmLast = crm.getLastRow();
  if (client && crmLast >= CRM_START) {
    const cv = crm.getRange(CRM_START, 1, crmLast - CRM_START + 1, CRM_COL_LOC).getValues();
    const canon = crmCanonName(client);
    for (var k = 0; k < cv.length; k++) {
      if (crmCanonName(cv[k][0]) !== canon) continue;
      lieu = String(cv[k][CRM_COL_LOC - 1] == null ? "" : cv[k][CRM_COL_LOC - 1]).trim();
      break;
    }
  }

  // The original, from "BL origine".
  const snap = ss.getSheetByName(BL_CFG.SNAP_SHEET);
  if (!snap) {
    throw new Error("Onglet « " + BL_CFG.SNAP_SHEET + " » absent du fichier Commandes. " +
                    "Prévenir Kim (blSetup).");
  }
  const want = {};
  targets.forEach(function (r) { want[r] = true; });
  const origin = [];
  var snapAt = null;
  blSnapRows(snap).forEach(function (s) {
    const r = Number(s[0]);
    if (!want[r]) return;
    if (String(s[1] || "").trim() !== orderNo) return;
    if (crmCanonName(s[2]) !== crmCanonName(client)) return;
    if (origin.some(function (o) { return o.row === r; })) return;
    const g = blGoodsFromSnap(s);
    g.row = r;
    origin.push(g);
    if (s[5] instanceof Date && (!snapAt || s[5] < snapAt)) snapAt = s[5];
  });
  const hasOrigin = origin.length > 0;

  const FIELDS = ["lot", "alevins", "pmAl", "livrer", "kg", "pmGr"];
  var goodsChanged = false;
  cur.forEach(function (g) {
    g.changed = {};
    if (!hasOrigin) return;
    const o = origin.filter(function (x) { return x.row === g.row; })[0];
    FIELDS.forEach(function (f) {
      if (!o || !blSame(g[f], o[f])) g.changed[f] = true;
    });
    if (Object.keys(g.changed).length) goodsChanged = true;
  });

  const livLabel = function (g) {
    const code = g.livraison;
    return (BL_CFG.LIV_LABELS[code] || code || "—") + (g.km ? " · " + g.km + " km" : "");
  };
  const livTxt = livLabel(cur[0]);
  const livTxtOrigin = hasOrigin ? livLabel(origin[0]) : "";
  const livChanged = hasOrigin && livTxt !== livTxtOrigin;

  return {
    bl: bl, orderNumber: orderNo || "(sans numéro)", dateCmd: dateCmd, dateLiv: dateLiv,
    client: client, tel: tel, lieu: lieu, remarques: remarques,
    livTxt: livTxt, livTxtOrigin: livTxtOrigin, livChanged: livChanged,
    hasOrigin: hasOrigin,
    amended: goodsChanged || livChanged,
    snapAt: snapAt ? Utilities.formatDate(snapAt, tz, "dd/MM/yyyy HH:mm") : "",
    current: blLines(cur, true),
    origin: hasOrigin ? blLines(origin, false) : null
  };
}

/* ---------- layout ---------- */

/** Builds the throw-away Sheet. Returns { ss, sh, lastRow, ncol }. */
function blBuildSheet(d) {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const W = BL_CFG.TABLE_HEADERS.length;
  const L = [];     // { vals, style, changed, table }
  const push = function (vals, style, changed, table) {
    L.push({ vals: vals, style: style || "plain", changed: changed || [], table: !!table });
  };
  const text = function (s, style) { push([s, "", "", "", ""], style); };
  const table = function (t, showChanges) {
    push(BL_CFG.TABLE_HEADERS.slice(), "head", [], true);
    t.lines.forEach(function (l) { push(l.cells, "line", showChanges ? l.changed : [], true); });
    push(["Total", "", "", t.kg || "", t.nombre || ""], "total", [], true);
  };

  FACT_CFG.SOCIETE.forEach(function (s, i) { text(s, i === 0 ? "bold" : "plain"); });
  text("");
  text(BL_CFG.TITLE + " N° " + d.bl, "title");
  text("");
  text("Commande : " + d.orderNumber + (d.dateCmd ? " du " + d.dateCmd : ""));
  text("Date de livraison : " + d.dateLiv);
  text("Client : " + d.client);
  if (d.tel) text("Téléphone : " + d.tel);
  if (d.lieu) text("Lieu de livraison : " + d.lieu);
  text("Mode de livraison : " + d.livTxt +
       (d.livChanged ? " (d'origine : " + d.livTxtOrigin + ")" : ""));
  if (d.remarques) text("Remarques : " + d.remarques);
  text("");

  if (d.amended) {
    text("1. Commande livrée — version d'origine (enregistrée le " + d.snapAt + ")", "bold");
    table(d.origin, false);
    text("");
    text("2. Commande modifiée après livraison — état au " +
         Utilities.formatDate(now, tz, "dd/MM/yyyy HH:mm"), "bold");
    table(d.current, true);
    text("Les cases surlignées ont changé depuis la livraison.", "note");
  } else {
    text("Marchandises livrées", "bold");
    table(d.current, false);
    if (!d.hasOrigin) {
      text("Version d'origine non enregistrée : commande livrée avant l'enregistrement " +
           "automatique (16/09/2026).", "note");
    }
  }
  text("");
  text("");
  push(["Livré par :", "", "Reçu par (nom et signature) :", "", ""], "plain");
  text("");
  text("");
  text("");
  text("Imprimé le " + Utilities.formatDate(now, tz, "dd/MM/yyyy HH:mm"), "note");

  // Folder FIRST: an unreachable folder fails before any file exists.
  const folder = DriveApp.getFolderById(HIST_REPORT_CFG.FOLDER_ID);
  const name = "BL " + d.bl + " (généré " +
    Utilities.formatDate(now, tz, "dd-MM-yy HH:mm").replace(":", "h") + ")";
  const ss = SpreadsheetApp.create(name);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  ss.setSpreadsheetTimeZone(tz);
  const sh = ss.getSheets()[0].setName("BL");
  prbMark("create+move");

  const n = L.length;
  if (sh.getMaxRows() < n) sh.insertRowsAfter(sh.getMaxRows(), n - sh.getMaxRows());
  if (sh.getMaxRows() > n) sh.deleteRows(n + 1, sh.getMaxRows() - n);
  if (sh.getMaxColumns() > W) sh.deleteColumns(W + 1, sh.getMaxColumns() - W);

  const values = [], weights = [], bgs = [], sizes = [], styles = [], formats = [];
  L.forEach(function (row) {
    const isTable = row.style === "line" || row.style === "total";
    values.push(row.vals);
    weights.push(row.vals.map(function (x, c) {
      return (row.style === "title" || row.style === "bold" || row.style === "head" ||
              row.style === "total" || row.changed.indexOf(c) >= 0) ? "bold" : "normal";
    }));
    bgs.push(row.vals.map(function (x, c) {
      return row.style === "head" ? BL_CFG.HEAD_BG
           : row.changed.indexOf(c) >= 0 ? BL_CFG.CHANGED_BG : null;
    }));
    sizes.push(row.vals.map(function () { return row.style === "title" ? 14 : 11; }));
    styles.push(row.vals.map(function () { return row.style === "note" ? "italic" : "normal"; }));
    // A, B text (a lot name must not become a date); C PM as typed;
    // D kg one decimal; E whole number.
    formats.push(["@", "@",
                  isTable ? "General" : "@",
                  isTable ? "#,##0.0" : "@",
                  isTable ? "#,##0" : "@"]);
  });
  const all = sh.getRange(1, 1, n, W);
  all.setNumberFormats(formats);
  all.setValues(values);
  all.setFontWeights(weights).setBackgrounds(bgs).setFontSizes(sizes).setFontStyles(styles);

  // Borders around each table block (head .. total).
  var start = -1;
  for (var i = 0; i <= n; i++) {
    const inTable = i < n && L[i].table;
    if (inTable && start < 0) start = i;
    if (!inTable && start >= 0) {
      sh.getRange(start + 1, 1, i - start, W).setBorder(true, true, true, true, true, true);
      start = -1;
    }
  }
  BL_CFG.COL_WIDTHS.forEach(function (w, c) { sh.setColumnWidth(c + 1, w); });
  SpreadsheetApp.flush();
  prbMark("format+flush");

  return { ss: ss, sh: sh, lastRow: n, ncol: W };
}

/* ---------- archive ---------- */

/** "BL <n°> - <commande> - <client>.pdf", safe for Drive. */
function blFileName(d) {
  return ("BL " + d.bl + " - " + d.orderNumber + " - " + d.client)
    .replace(/[\\\/:*?"<>|]/g, "-") + ".pdf";
}

/** Saves the PDF in "Bons de livraison"; a same-name file goes to the bin. */
function blArchive(pdf) {
  const blob = Utilities.newBlob(Utilities.base64Decode(pdf.base64), MimeType.PDF, pdf.filename);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(BL_CFG.LOCK_MS)) {
    throw new Error("Une autre impression est en cours : copie non enregistrée.");
  }
  try {
    var folder;
    try {
      folder = DriveApp.getFolderById(BL_CFG.FOLDER_ID);
    } catch (e) {
      throw new Error("Dossier « " + BL_CFG.FOLDER_NAME + " » introuvable (" + BL_CFG.FOLDER_ID +
        "). Kim : tsaraentry -> BonLivraisonServer.js -> BL_CFG.");
    }
    if (folder.isTrashed()) {
      throw new Error("Dossier « " + BL_CFG.FOLDER_NAME + " » est à la corbeille.");
    }
    const old = folder.getFilesByName(pdf.filename);
    while (old.hasNext()) old.next().setTrashed(true);
    return folder.createFile(blob).getUrl();
  } finally {
    lock.releaseLock();
  }
}

/* ---------- the button ---------- */

/**
 * "Imprimer le bon de livraison externe". rows = the order rows of the card.
 * Returns { base64, filename, size, bl, amended, hasOrigin,
 *           archiveUrl | archiveError }.
 */
function blPdfPrbBody(rows) {
  const d = blData(rows);
  prbMark("data");
  const b = blBuildSheet(d);
  const file = DriveApp.getFileById(b.ss.getId());
  var pdf;
  try {
    const url = "https://docs.google.com/spreadsheets/d/" + b.ss.getId() + "/export"
      + "?format=pdf"
      + "&gid=" + b.sh.getSheetId()
      + "&size=A4"
      + "&portrait=true"
      + "&fitw=true"
      + "&gridlines=false"
      + "&printtitle=false"
      + "&sheetnames=false"
      + "&pagenum=UNDEFINED"
      + "&attachment=false"
      + "&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5"
      + "&r1=0"
      + "&c1=0"
      + "&r2=" + b.lastRow
      + "&c2=" + b.ncol;
    pdf = tempFetchPdf(url, "BL");
    prbMark("export");
  } finally {
    // Export done or failed, the Sheet has served: bin it either way.
    file.setTrashed(true);
  }
  pdf.filename = blFileName(d);
  try {
    pdf.archiveUrl = blArchive(pdf);
  } catch (e) {
    pdf.archiveError = e.message;
    console.error("Archivage BL : " + e.message);
  }
  prbMark("archive");
  pdf.bl = d.bl;
  pdf.amended = d.amended;
  pdf.hasOrigin = d.hasOrigin;
  return pdf;
}

function blPdf(rows) {
  var own = prbStart("blPdf", "rows=" + (rows || []).length);
  try { return blPdfPrbBody(rows); }
  finally { prbEnd(own); }
}

/* ---------- editor ---------- */

/**
 * RUN FROM EDITOR ONCE: tsaraentry -> BonLivraisonServer.js -> blSetup
 * Creates the tab "BL origine" at the end of the Commandes file, exact
 * size, headers in row 1. Safe to re-run: an existing tab is left alone.
 * Also proves the Drive folder resolves.
 */
function blSetup() {
  const ss = SpreadsheetApp.openById(CMD_CFG.SS_ID);
  var sh = ss.getSheetByName(BL_CFG.SNAP_SHEET);
  const n = BL_CFG.SNAP_HEADERS.length;
  if (sh) {
    Logger.log("Onglet « " + BL_CFG.SNAP_SHEET + " » déjà présent : " +
               Math.max(0, sh.getLastRow() - 1) + " ligne(s).");
  } else {
    sh = ss.insertSheet(BL_CFG.SNAP_SHEET, ss.getSheets().length);
    if (sh.getMaxColumns() > n) sh.deleteColumns(n + 1, sh.getMaxColumns() - n);
    if (sh.getMaxRows() > 2) sh.deleteRows(3, sh.getMaxRows() - 2);
    sh.getRange(1, 1, 1, n).setValues([BL_CFG.SNAP_HEADERS])
      .setFontWeight("bold").setBackground(BL_CFG.HEAD_BG);
    sh.setFrozenRows(1);
    Logger.log("Onglet « " + BL_CFG.SNAP_SHEET + " » créé.");
  }
  const folder = DriveApp.getFolderById(BL_CFG.FOLDER_ID);
  Logger.log("Dossier : " + folder.getName() + " · " + folder.getUrl());
}

/**
 * RUN FROM EDITOR: tsaraentry -> BonLivraisonServer.js -> testBonLivraison
 * READ-ONLY. Takes the first delivered order with a N° BL in the
 * Paiement & livraison list and logs what its BL would print.
 */
function testBonLivraison() {
  const snap = SpreadsheetApp.openById(CMD_CFG.SS_ID).getSheetByName(BL_CFG.SNAP_SHEET);
  Logger.log(snap ? "Onglet « BL origine » : " + Math.max(0, snap.getLastRow() - 1) + " ligne(s)."
                  : "Onglet « BL origine » ABSENT : lancer blSetup.");
  const o = cmdFindOrders("", null, true, true, true).orders.filter(function (x) {
    return String(x.dateLivraison || "").trim() && String(x.bl || "").trim();
  })[0];
  if (!o) { Logger.log("Aucune commande livrée avec un N° BL dans la liste."); return; }
  const d = blData(o.rows);
  Logger.log(JSON.stringify({
    bl: d.bl, commande: d.orderNumber, client: d.client, livraison: d.dateLiv,
    lieu: d.lieu, mode: d.livTxt, hasOrigin: d.hasOrigin, amended: d.amended,
    current: d.current, origin: d.origin
  }, null, 1));
}
