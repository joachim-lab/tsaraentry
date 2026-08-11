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
 * Parse a "yyyy-MM-dd" string from a browser date input into a
 * midnight LOCAL date (script timezone), never new Date(str) directly.
 * new Date("2026-08-11") parses as UTC midnight, which becomes
 * 03:00:00 once Sheets displays it in Africa/Nairobi (UTC+3) — a
 * timestamp on a cell every other writer in this system leaves as a
 * plain date. If anything ever compares payment/delivery dates by
 * equality, a value carrying a time component could silently fail
 * to match a date-only value written elsewhere.
 */
function cmdParseDate(isoStr) {
  if (!isoStr) return undefined;
  const m = String(isoStr).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return undefined;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
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
    // Column B (order number) is written later, row by row — see the
    // sequential generation block below.
    put(C.FACTURE, f.facture);
    put(C.BL, f.bl);
    put(C.DATE_CMD, cmdParseDate(f.dateCommande));

    // ln.* values arrive already parsed to a number-or-null by the
    // browser's num() (see CommandesIndex.html). Re-wrapping them in
    // Number(...) here was the bug: Number(null) is 0, not "no value",
    // and 0 is not undefined/null/"" so put() wrote it as real data
    // (PM alevins / prix showed 0 when the fields were left blank).
    // put() already skips undefined/null/"" — pass values through as-is.
    if (isAlevins) {
      put(C.ALEVINS_NB, ln.alevinsNb);
      put(C.ALEVINS_PM, ln.alevinsPm);
      put(C.ALEVINS_LIVRER, ln.alevinsLivrer);
      put(C.ALEVINS_PRIX, ln.alevinsPrix);
      put(C.TRANSPORT, ln.transport);
    } else {
      put(C.POISSON_KG, ln.poissonKg);
      put(C.POISSON_PM, ln.poissonPm);
      put(C.PRIX_KG, ln.prixKg);
      put(C.FRAIS, ln.frais);
    }

    // Client/contact/remarks repeat on every row, as in existing data.
    put(C.CLIENT, f.client);
    put(C.REMARQUES, f.remarques);
    put(C.CONTACT, f.contact);
  });

  const lastRow = firstRow + lines.length - 1;

  // NOTE: the app does NOT copy formulas into new rows. Columns K, N,
  // Q and W are already pre-filled with formulas far below the last
  // data row (this is why getLastRow reports ~2051 while real orders
  // end around row 180), so new rows inherit them automatically.
  // An earlier version copied them from the row above; that was
  // removed because it would propagate a gap if the preceding row ever
  // had its formulas cleared by hand.

  SpreadsheetApp.flush();

  // Order numbers: edit triggers never fire on programmatic writes.
  //
  // Generate ROW BY ROW, not as one range. ac_enforceOrderRange_
  // resolves "commande groupée" by reading the cell ABOVE as it goes,
  // before generation happens later in the same pass — so sweeping the
  // whole block at once makes row 2 see a literal "AL" above it, treat
  // it as a prefix, and mint its own number (AL-26-164 / AL-26-165
  // instead of both being 164). Going one row at a time reproduces
  // exactly what sequential manual typing does: row 1 is generated
  // first, then each grouped row copies an already-generated number.
  let orderNumber = "";
  try {
    for (let i = 0; i < lines.length; i++) {
      const row = firstRow + i;
      sh.getRange(row, C.ORDER_NO).setValue(i === 0 ? f.type : "commande groupée");
      SpreadsheetApp.flush();
      AutoCommandes.generateOrderNumbersForRows(row, row);
      SpreadsheetApp.flush();
    }
    orderNumber = sh.getRange(firstRow, C.ORDER_NO).getDisplayValue();
  } catch (err) {
    throw new Error("La commande a été enregistrée (lignes " + firstRow + "-" + lastRow +
      ") mais le numéro de commande n'a pas pu être généré : " + err + ". Prévenir Kim.");
  }

  return { firstRow: firstRow, lastRow: lastRow, orderNumber: orderNumber, rowCount: lines.length };
}

/**
 * MODE B — find orders, GROUPED by order number.
 *
 * A commercial order can span several consecutive rows (one per lot,
 * sharing one number via "commande groupée"), and fulfilment applies
 * to the whole order — rows 179/180 carry the same payment date. So
 * this returns one entry per order, listing its rows.
 *
 * Rows with a blank order number can't be grouped safely, so each is
 * returned on its own, keyed by row.
 *
 * `query` matches order number, client or lot (case-insensitive).
 * onlyUnfulfilled = no payment date recorded on any of its rows.
 */
function cmdFindOrders(query, onlyUnfulfilled) {
  const sh = cmdSheet();
  const C = CMD_CFG.COL;
  const lastRow = findNextCommandeRow(sh) - 1;
  if (lastRow < CMD_CFG.START_ROW) return [];

  const n = lastRow - CMD_CFG.START_ROW + 1;
  const vals = sh.getRange(CMD_CFG.START_ROW, 1, n, C.ANNULE).getDisplayValues();
  const q = String(query || "").trim().toLowerCase();

  const groups = {};
  const order = [];

  for (let i = 0; i < n; i++) {
    const r = vals[i];
    const rowNum = CMD_CFG.START_ROW + i;
    if (String(r[C.ANNULE - 1] || "").trim()) continue;      // annulé

    const orderNo = String(r[C.ORDER_NO - 1] || "").trim();
    const key = orderNo || ("__row" + rowNum);

    if (!groups[key]) {
      groups[key] = {
        orderNumber: orderNo,
        key: key,
        rows: [],
        lots: [],
        client: r[C.CLIENT - 1],
        contact: r[C.CONTACT - 1],
        dateCommande: r[C.DATE_CMD - 1],
        paiement: r[C.PAIEMENT - 1],
        dateLivraison: r[C.DATE_LIVRAISON - 1],
        moyenPaiement: r[C.MOYEN_PAIEMENT - 1],
        livre: r[C.LIVRE - 1],
        facture: r[C.FACTURE - 1],
        bl: r[C.BL - 1],
        alevinsNb: [],
        poissonKg: []
      };
      order.push(key);
    }

    const g = groups[key];
    g.rows.push(rowNum);
    if (r[C.LOT - 1]) g.lots.push(r[C.LOT - 1]);
    if (r[C.ALEVINS_NB - 1]) g.alevinsNb.push(r[C.ALEVINS_NB - 1]);
    if (r[C.POISSON_KG - 1]) g.poissonKg.push(r[C.POISSON_KG - 1]);
    // Any row carrying fulfilment data represents the order's state.
    if (!g.paiement && r[C.PAIEMENT - 1]) g.paiement = r[C.PAIEMENT - 1];
    if (!g.dateLivraison && r[C.DATE_LIVRAISON - 1]) g.dateLivraison = r[C.DATE_LIVRAISON - 1];
    if (!g.moyenPaiement && r[C.MOYEN_PAIEMENT - 1]) g.moyenPaiement = r[C.MOYEN_PAIEMENT - 1];
    if (!g.facture && r[C.FACTURE - 1]) g.facture = r[C.FACTURE - 1];
    if (!g.bl && r[C.BL - 1]) g.bl = r[C.BL - 1];
  }

  const out = [];
  for (let i = order.length - 1; i >= 0; i--) {              // newest first
    const g = groups[order[i]];

    if (q) {
      const hay = (g.orderNumber + " " + g.client + " " + g.lots.join(" ")).toLowerCase();
      if (hay.indexOf(q) === -1) continue;
    }
    if (onlyUnfulfilled && String(g.paiement || "").trim() !== "") continue;

    out.push(g);
    if (out.length >= 25) break;
  }
  return out;
}

/**
 * MODE B — record fulfilment across EVERY row of one order.
 * Only U / V / X are written; W (livré) is a formula driven by V.
 * `rows` comes from cmdFindOrders, so the caller never guesses.
 * Returns { rows, changed: [...] }.
 */
function cmdRecordFulfilment(rows, payload) {
  const sh = cmdSheet();
  const C = CMD_CFG.COL;
  const lastRow = findNextCommandeRow(sh) - 1;

  const targets = (rows || []).map(Number).filter(r =>
    isFinite(r) && r >= CMD_CFG.START_ROW && r <= lastRow);
  if (!targets.length) throw new Error("Aucune ligne de commande valide à mettre à jour.");

  const f = payload || {};
  const changed = [];

  targets.forEach(r => {
    function put(col, value, label) {
      if (value === undefined || value === null || value === "") return;
      const cell = sh.getRange(r, col);
      const before = cell.getDisplayValue();
      cell.setValue(value);
      const after = cell.getDisplayValue();
      if (before !== after) changed.push("L" + r + " " + label + ": " + (before || "(vide)") + " -> " + after);
    }

    put(C.PAIEMENT, cmdParseDate(f.paiement), "Paiement reçu");
    put(C.DATE_LIVRAISON, cmdParseDate(f.dateLivraison), "Date livraison");
    put(C.MOYEN_PAIEMENT, f.moyenPaiement, "Moyen paiement");
    // Invoice / delivery-note numbers usually arrive after the order is
    // placed, and apply to the whole order (Kim, 2026-08-11).
    put(C.FACTURE, f.facture, "N° facture");
    put(C.BL, f.bl, "Bon de livraison");
  });

  SpreadsheetApp.flush();
  return { rows: targets, changed: changed };
}

/** RUN FROM EDITOR: read-only checks of the screen-3 server side. */
function testCommandesServer() {
  const sh = cmdSheet();
  Logger.log("Prochaine ligne libre: " + findNextCommandeRow(sh));

  const opts = cmdGetOptions();
  Logger.log("Lots (" + opts.lots.length + "): " + opts.lots.slice(0, 8).join(" | "));
  Logger.log("Types: " + opts.types.join(" | "));

  Logger.log("--- 3 dernières commandes (groupées) ---");
  cmdFindOrders("", false).slice(0, 3).forEach(o =>
    Logger.log(o.orderNumber + " | " + o.client + " | lots: " + o.lots.join(", ") +
      " | lignes: " + o.rows.join(",") + " | payé: " + (o.paiement || "non")));

  Logger.log("--- commandes non soldées (max 3) ---");
  cmdFindOrders("", true).slice(0, 3).forEach(o =>
    Logger.log(o.orderNumber + " / " + o.client + " / lots: " + o.lots.join(", ") +
      " / lignes: " + o.rows.join(",")));

  Logger.log("--- test binding AutoCommandes (plage vide, n'écrit rien) ---");
  Logger.log(AutoCommandes.generateOrderNumbersForRows(3, 2));
}

/* =============================================================
 * ITEM 3 — entry-time stock validation (2026-08-11)
 *
 * Mirrors the nightly engine's deduction logic so an order that would
 * fail overnight is caught at the counter instead.
 *
 * Rule set (agreed with Kim 2026-08-11, each grounded in live code):
 *   TOUT reservation   -> BLOCK   (engine: reserved = Infinity)
 *   qty > available    -> BLOCK   (engine: next < 0, or next < reserved)
 *   PM mismatch        -> WARN    (advisory only, engine ignores PM)
 *   key not found      -> WARN    (engine writes "NOT FOUND ON ")
 *
 * Why "not found" only warns: the Commandes lot dropdown is built by
 * tt_applyCommandesDropdown_ from Stock Poisson "lot"!N3:N50, while this
 * function reads the LOT FILE. Two different sources, so they can drift
 * apart legitimately. Blocking on a disagreement we cannot adjudicate
 * would stop a real sale being recorded. Quantity is different: that
 * number comes from the very cell the engine deducts, so it is certain
 * and safe to block on.
 * ============================================================= */

/**
 * Stock picture for one order key, as the engine would see it.
 *
 * @param {string} orderKey  raw lot key, e.g. "24-4-B"
 * @return {Object} {
 *   key, found, source, col|row, count, pm,
 *   reserved,            // number, or the string "TOUT"
 *   pending,             // already-entered rows not yet deducted
 *   available,           // count - reserved - pending (null if TOUT)
 *   reservedAll          // true => block outright
 * }
 */
function cmdGetLotAvailability(orderKey) {
  const key = cmdCanonKey(orderKey);
  const out = {
    key: key, found: false, source: null, count: null, pm: null,
    reserved: 0, pending: 0, available: null, reservedAll: false
  };
  if (!key) return out;

  // ---- reservation first: TOUT short-circuits everything ----
  // Must come before the arithmetic, otherwise available goes -Infinity
  // and that value could reach the UI.
  const res = cmdGetReservation(key);
  if (res === "TOUT") {
    out.reserved = "TOUT";
    out.reservedAll = true;
    return out;
  }
  out.reserved = res;

  // ---- locate the lot file ----
  const lotNum = key.split("-")[0];
  const list = getLotFileList();
  var fileId = null;
  for (var i = 0; i < list.length; i++) {
    if (cmdCanonKey(list[i].lotNumber) === lotNum) { fileId = list[i].fileId; break; }
  }
  if (!fileId) return out;   // found stays false -> caller warns

  // ---- the cell the engine would deduct ----
  const m = findSubLotColumnByOrderKey(SpreadsheetApp.openById(fileId), key);
  if (!m.found) return out;

  out.found = true;
  out.source = m.source;
  if (m.col) out.col = m.col;
  if (m.row) out.row = m.row;
  out.count = m.count;
  out.pm = m.pm;

  out.pending = cmdGetPendingQty(key);
  out.available = out.count - out.reserved - out.pending;
  return out;
}

/**
 * Reservation for one canon key. Mirrors tt_loadReservations.
 * Returns the string "TOUT", or a positive number, or 0.
 */
function cmdGetReservation(canonKey) {
  const ss = SpreadsheetApp.openById(CMD_CFG.SS_ID);
  const sh = ss.getSheetByName("Réservations");
  if (!sh) return 0;                    // missing tab = no reservations
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const vals = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  var qty = 0;
  for (var i = 0; i < vals.length; i++) {
    if (cmdCanonKey(vals[i][0]) !== canonKey) continue;
    const type = String(vals[i][1] || "").trim().toUpperCase();
    if (type === "TOUT") return "TOUT";
    if (type === "NOMBRE") {
      const n = cmdToNum(vals[i][2]);
      if (n != null && n > 0) qty = n;
    }
  }
  return qty;
}

/**
 * Sum of orders already in the sheet for this key that have NOT yet been
 * deducted. Mirrors tt_commandeIsEligible_: Y, Z and AA all empty, and
 * H-or-N resolves to a positive number.
 *
 * Z MATTERS. A row blocked on a previous night has Z filled and is never
 * retried by the engine, so it will never deduct. Counting it as pending
 * would understate availability and block good orders. Confirmed live
 * 2026-08-11 on 24-4-B: two LOT RÉSERVÉ rows, pending correctly 0.
 *
 * NOTE: this term has NO counterpart in engine_core.js. The engine
 * deducts every eligible row in one pass against a shrinking in-memory
 * row10, so it never needs it. This function validates BEFORE any
 * deduction has happened, so it does. Deliberate divergence — do not
 * "correct" it to match the engine.
 */
function cmdGetPendingQty(canonKey) {
  const sh = cmdSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < CMD_CFG.START_ROW) return 0;

  const data = sh.getRange(CMD_CFG.START_ROW, 1,
                           lastRow - CMD_CFG.START_ROW + 1, 27).getValues();
  var total = 0;
  for (var i = 0; i < data.length; i++) {
    const r = data[i];
    if (cmdCanonKey(r[CMD_CFG.COL.LOT - 1]) !== canonKey) continue;
    if (String(r[CMD_CFG.COL.LOG - 1]    || "").trim() !== "") continue;  // Y
    if (String(r[CMD_CFG.COL.ERROR - 1]  || "").trim() !== "") continue;  // Z
    if (String(r[CMD_CFG.COL.ANNULE - 1] || "").trim() !== "") continue;  // AA
    const ded = cmdDeduction(r[CMD_CFG.COL.ALEVINS_LIVRER - 1],
                             r[CMD_CFG.COL.POISSON_NB - 1]);
    if (ded == null) continue;
    total += ded;
  }
  return total;
}

/** H wins over N. Mirrors tt_commandeDeduction_. */
function cmdDeduction(valH, valN) {
  const h = cmdToNum(valH);
  const n = cmdToNum(valN);
  if (h != null && h > 0) return h;
  if (n != null && n > 0) return n;
  return null;
}

/** Mirrors tt_toNum_ : NBSP-tolerant, comma decimal. */
function cmdToNum(v) {
  if (v === "" || v == null) return null;
  const s = String(v).replace(/\u00A0/g, " ").replace(/\s+/g, "").replace(",", ".");
  const n = Number(s);
  return isFinite(n) ? n : null;
}

/**
 * Validate a whole submission before saving.
 *
 * Lines are summed BY CANON KEY, not checked one at a time: a single
 * order can list the same lot on two rows, and those rows are not yet in
 * the sheet so cmdGetPendingQty cannot see them. Checking row by row
 * would let 3000 + 3000 through against a stock of 5000.
 *
 * @param {Array} lines  [{ lot, qty, pm }]  qty already H-or-N resolved
 * @return {Object} { ok, blocks: [msg], warnings: [msg], detail: {key: avail} }
 */
function cmdValidateOrderLines(lines) {
  const blocks = [];
  const warnings = [];
  const detail = {};
  if (!lines || !lines.length) return { ok: true, blocks: blocks, warnings: warnings, detail: detail };

  // sum this submission by canon key
  const wanted = {};
  const pmSeen = {};
  lines.forEach(function (ln) {
    const k = cmdCanonKey(ln.lot);
    if (!k) return;
    const q = cmdToNum(ln.qty);
    wanted[k] = (wanted[k] || 0) + (q != null && q > 0 ? q : 0);
    if (ln.pm != null && ln.pm !== "") pmSeen[k] = cmdToNum(ln.pm);
  });

  Object.keys(wanted).forEach(function (k) {
    const a = cmdGetLotAvailability(k);
    detail[k] = a;

    if (a.reservedAll) {
      blocks.push(k + " : ce lot est réservé, choisissez un autre lot");
      return;
    }
    if (!a.found) {
      warnings.push(k + " : ce lot n'est pas trouvé dans le fichier lot ; " +
                        "la déduction échouera cette nuit");
      return;
    }
    // strict > : ordering exactly down to the reservation floor is legal,
    // because the engine blocks on next < reserved, not next <= reserved.
    if (wanted[k] > a.available) {
      blocks.push(k + " : stock insuffisant — demandé " + Math.round(wanted[k]) +
                      ", disponible " + a.available +
                      " (stock " + a.count +
                      (a.reserved ? ", réservé " + a.reserved : "") +
                      (a.pending ? ", en attente " + a.pending : "") + ")");
    }
    if (pmSeen[k] != null && a.pm != null && pmSeen[k] !== a.pm) {
      warnings.push(k + " : PM saisi " + pmSeen[k] + " ≠ PM du lot " + a.pm);
    }
  });

  return { ok: blocks.length === 0, blocks: blocks, warnings: warnings, detail: detail };
}
