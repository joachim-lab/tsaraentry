/***************************************************************
 * CommandesServer.js — TSARA Entry web app
 * Screen 3: Commandes — server side.
 *
 * Two modes (Kim's decision, 2026-08-09):
 *   A) create a new order
 *   B) find an existing order and record fulfilment
 *      (U paiement, V date livraison, X moyen de paiement)
 *
 * ORDER NUMBERS: column B is a dropdown (AL / GR / commande groupée)
 * that automatismescommande converts to AL-YY-### on edit. Edit
 * triggers do NOT fire on programmatic writes, so after appending a
 * row this calls AutoCommandes.generateOrderNumbersForRows() — the
 * same code path manual edits use. Library bound at HEAD.
 *
 * NEVER WRITTEN (formulas — the sheet owns these):
 *   K = (F*I)+J          argent alevins
 *   N = (L*1000)/M       nombre poissons à livrer
 *   Q = (O*L)+P          argent poisson
 *   W = IF(V<>"";"x";"") livré
 *   Y, Z                 engine log / error
 *
 * NOTE on column H ("Alevins à livrer +5%"): normally the formula
 * =(F*0,05)+F, but staff DO override it with a typed value (row 180:
 * F=20439, H=21593, which is not F*1.05). H is what the engine
 * actually deducts from stock, so the app pre-computes +5% and lets
 * it be overridden — matching current practice.
 ***************************************************************/

const CMD_CFG = {
  SS_ID: "1QYxnnfMoYidqN8l0ZQZZe5MHER5EBeCEXXvTzyhKan8",
  SHEET: "2026",
  START_ROW: 2,

  COL: {
    LOT: 1, ORDER_NO: 2, FACTURE: 3, BL: 4, DATE_CMD: 5,
    ALEVINS_NB: 6, ALEVINS_PM: 7, ALEVINS_LIVRER: 8, ALEVINS_PRIX: 9, TRANSPORT: 10,
    ARGENT_ALEVINS: 11,
    POISSON_KG: 12, POISSON_PM: 13, POISSON_NB: 14, PRIX_KG: 15, FRAIS: 16,
    ARGENT_POISSON: 17,
    CLIENT: 18, REMARQUES: 19, CONTACT: 20,
    PAIEMENT: 21, DATE_LIVRAISON: 22, LIVRE: 23, MOYEN_PAIEMENT: 24,
    LOG: 25, ERROR: 26, ANNULE: 27
  }
};

function cmdSheet() {
  const ss = SpreadsheetApp.openById(CMD_CFG.SS_ID);
  const sh = ss.getSheetByName(CMD_CFG.SHEET);
  if (!sh) throw new Error('Onglet introuvable: "' + CMD_CFG.SHEET + '"');
  return sh;
}

/**
 * First row after the last real order. Scans column A — getLastRow()
 * overshoots badly here (reports 2051 when real data ends at 180).
 */
function findNextCommandeRow(sh) {
  const physical = sh.getLastRow();
  if (physical < CMD_CFG.START_ROW) return CMD_CFG.START_ROW;
  const vals = sh.getRange(CMD_CFG.START_ROW, CMD_CFG.COL.LOT,
    physical - CMD_CFG.START_ROW + 1, 1).getDisplayValues();
  let lastData = CMD_CFG.START_ROW - 1;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || "").trim() !== "") lastData = CMD_CFG.START_ROW + i;
  }
  return lastData + 1;
}

/** Dropdown options read live from the sheet's own validation rules. */
function cmdGetOptions() {
  const sh = cmdSheet();
  const probeRow = Math.max(CMD_CFG.START_ROW, findNextCommandeRow(sh) - 1);

  function listAt(col) {
    const rule = sh.getRange(probeRow, col).getDataValidation();
    if (!rule) return [];
    try {
      const cv = rule.getCriteriaValues();
      return Array.isArray(cv[0]) ? cv[0] : [];
    } catch (e) { return []; }
  }

  return {
    lots: listAt(CMD_CFG.COL.LOT),
    types: listAt(CMD_CFG.COL.ORDER_NO)
  };
}

/**
 * MODE A — create an order, possibly spanning SEVERAL lots.
 *
 * A commercial order can cover more than one lot: the rows sit
 * consecutively, the FIRST carries the type (AL or GR) and the rest
 * carry "commande groupée", which automatismescommande resolves by
 * copying the order number from the row above. Live example: rows
 * 179-180 both show AL-26-163 for lots 24-4-C and 24-4-A.
 *
 * The type also decides the kind: AL = alevins, GR = poisson. Mixing
 * both in one order would contradict the number prefix, so every line
 * uses the block matching the type.
 *
 * payload = {
 *   type, dateCommande, facture, bl, client, contact, remarques,
 *   lines: [ { lot, ...quantities } ]
 * }
 * Returns { firstRow, lastRow, orderNumber, rowCount }.
 */
function cmdCreateOrder(payload) {
  const f = payload || {};
  const lines = f.lines || [];

  if (!f.type) throw new Error("Le type de commande est obligatoire.");
  if (!f.client) throw new Error("Le client est obligatoire.");
  if (!lines.length) throw new Error("Ajouter au moins un lot à la commande.");

  const isAlevins = String(f.type).toUpperCase().indexOf("AL") === 0;

  lines.forEach((ln, i) => {
    if (!ln.lot) throw new Error("Ligne " + (i + 1) + " : le lot est obligatoire.");
    if (isAlevins) {
      if (ln.alevinsNb === undefined || ln.alevinsNb === null || ln.alevinsNb === "") {
        throw new Error("Ligne " + (i + 1) + " : nombre d'alevins obligatoire.");
      }
    } else {
      if (ln.poissonKg === undefined || ln.poissonKg === null || ln.poissonKg === "") {
        throw new Error("Ligne " + (i + 1) + " : quantité de poisson (kg) obligatoire.");
      }
    }
  });

  const sh = cmdSheet();
  const firstRow = findNextCommandeRow(sh);
  const C = CMD_CFG.COL;

  lines.forEach((ln, i) => {
    const row = firstRow + i;

    function put(col, value) {
      if (value === undefined || value === null || value === "") return;
      sh.getRange(row, col).setValue(value);
    }

    put(C.LOT, ln.lot);
    // First row carries the real type; the rest group onto it.
    put(C.ORDER_NO, i === 0 ? f.type : "commande groupée");
    put(C.FACTURE, f.facture);
    put(C.BL, f.bl);
    put(C.DATE_CMD, f.dateCommande ? new Date(f.dateCommande) : undefined);

    if (isAlevins) {
      put(C.ALEVINS_NB, Number(ln.alevinsNb));
      put(C.ALEVINS_PM, ln.alevinsPm === "" ? undefined : Number(ln.alevinsPm));
      put(C.ALEVINS_LIVRER, ln.alevinsLivrer === "" ? undefined : Number(ln.alevinsLivrer));
      put(C.ALEVINS_PRIX, ln.alevinsPrix === "" ? undefined : Number(ln.alevinsPrix));
      put(C.TRANSPORT, ln.transport === "" ? undefined : Number(ln.transport));
    } else {
      put(C.POISSON_KG, Number(ln.poissonKg));
      put(C.POISSON_PM, ln.poissonPm === "" ? undefined : Number(ln.poissonPm));
      put(C.PRIX_KG, ln.prixKg === "" ? undefined : Number(ln.prixKg));
      put(C.FRAIS, ln.frais === "" ? undefined : Number(ln.frais));
    }

    // Client/contact/remarks repeat on every row, as in existing data.
    put(C.CLIENT, f.client);
    put(C.REMARQUES, f.remarques);
    put(C.CONTACT, f.contact);
  });

  const lastRow = firstRow + lines.length - 1;
  SpreadsheetApp.flush();

  // Order number: edit triggers never fire on programmatic writes.
  // Sweep the whole range so "commande groupée" rows resolve against
  // the first row's freshly generated number.
  let orderNumber = "";
  try {
    AutoCommandes.generateOrderNumbersForRows(firstRow, lastRow);
    SpreadsheetApp.flush();
    orderNumber = sh.getRange(firstRow, C.ORDER_NO).getDisplayValue();
  } catch (err) {
    throw new Error("La commande a été enregistrée (lignes " + firstRow + "-" + lastRow +
      ") mais le numéro de commande n'a pas pu être généré : " + err + ". Prévenir Kim.");
  }

  return { firstRow: firstRow, lastRow: lastRow, orderNumber: orderNumber, rowCount: lines.length };
}

/**
 * MODE B — find orders. Returns the most recent matching orders,
 * newest first. `query` matches order number, client or lot
 * (case-insensitive substring). Unfulfilled = no payment date yet.
 */
function cmdFindOrders(query, onlyUnfulfilled) {
  const sh = cmdSheet();
  const C = CMD_CFG.COL;
  const lastRow = findNextCommandeRow(sh) - 1;
  if (lastRow < CMD_CFG.START_ROW) return [];

  const n = lastRow - CMD_CFG.START_ROW + 1;
  const vals = sh.getRange(CMD_CFG.START_ROW, 1, n, C.ANNULE).getDisplayValues();
  const q = String(query || "").trim().toLowerCase();

  const out = [];
  for (let i = n - 1; i >= 0; i--) {         // newest first
    const r = vals[i];
    const rowNum = CMD_CFG.START_ROW + i;

    const orderNo = r[C.ORDER_NO - 1];
    const client  = r[C.CLIENT - 1];
    const lot     = r[C.LOT - 1];
    const annule  = String(r[C.ANNULE - 1] || "").trim();
    if (annule) continue;

    if (q) {
      const hay = (orderNo + " " + client + " " + lot).toLowerCase();
      if (hay.indexOf(q) === -1) continue;
    }

    const paiement = r[C.PAIEMENT - 1];
    if (onlyUnfulfilled && String(paiement || "").trim() !== "") continue;

    out.push({
      row: rowNum,
      orderNumber: orderNo,
      lot: lot,
      client: client,
      dateCommande: r[C.DATE_CMD - 1],
      alevinsNb: r[C.ALEVINS_NB - 1],
      poissonKg: r[C.POISSON_KG - 1],
      argentAlevins: r[C.ARGENT_ALEVINS - 1],
      argentPoisson: r[C.ARGENT_POISSON - 1],
      paiement: paiement,
      dateLivraison: r[C.DATE_LIVRAISON - 1],
      livre: r[C.LIVRE - 1],
      moyenPaiement: r[C.MOYEN_PAIEMENT - 1]
    });

    if (out.length >= 25) break;
  }
  return out;
}

/**
 * MODE B — record fulfilment on an existing order row.
 * Only U / V / X are written. W (livré) is a formula driven by V.
 * Returns { row, changed: [...] }.
 */
function cmdRecordFulfilment(row, payload) {
  const sh = cmdSheet();
  const C = CMD_CFG.COL;
  const r = Number(row);

  const lastRow = findNextCommandeRow(sh) - 1;
  if (!isFinite(r) || r < CMD_CFG.START_ROW || r > lastRow) {
    throw new Error("Ligne de commande invalide: " + row);
  }

  const f = payload || {};
  const changed = [];

  function put(col, value, label) {
    if (value === undefined || value === null || value === "") return;
    const cell = sh.getRange(r, col);
    const before = cell.getDisplayValue();
    cell.setValue(value);
    changed.push(label + ": " + (before || "(vide)") + " -> " + cell.getDisplayValue());
  }

  put(C.PAIEMENT, f.paiement ? new Date(f.paiement) : undefined, "Paiement reçu");
  put(C.DATE_LIVRAISON, f.dateLivraison ? new Date(f.dateLivraison) : undefined, "Date livraison");
  put(C.MOYEN_PAIEMENT, f.moyenPaiement, "Moyen paiement");

  SpreadsheetApp.flush();
  return { row: r, changed: changed };
}

/** RUN FROM EDITOR: read-only checks of the screen-3 server side. */
function testCommandesServer() {
  const sh = cmdSheet();
  Logger.log("Prochaine ligne libre: " + findNextCommandeRow(sh));

  const opts = cmdGetOptions();
  Logger.log("Lots (" + opts.lots.length + "): " + opts.lots.slice(0, 8).join(" | "));
  Logger.log("Types: " + opts.types.join(" | "));

  Logger.log("--- 3 dernières commandes ---");
  cmdFindOrders("", false).slice(0, 3).forEach(o => Logger.log(JSON.stringify(o)));

  Logger.log("--- commandes non soldées (max 3) ---");
  cmdFindOrders("", true).slice(0, 3).forEach(o =>
    Logger.log(o.orderNumber + " / " + o.client + " / " + o.lot));

  Logger.log("--- test binding AutoCommandes (plage vide, n'écrit rien) ---");
  Logger.log(AutoCommandes.generateOrderNumbersForRows(3, 2));
}
