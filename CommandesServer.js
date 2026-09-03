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

  // Stock guard on the server. The browser calls cmdValidateOrderLines
  // before saving, but that is the browser's choice; any other caller
  // would write an over-sold order unchecked. Same validator, same
  // verdict - the quantities are reshaped to the shape it expects.
  const vLines = lines.map(function (ln) {
    if (isAlevins) {
      return { lot: ln.lot, qty: ln.alevinsLivrer, pm: ln.alevinsPm };
    }
    const kg = cmdToNum(ln.poissonKg);
    const pm = cmdToNum(ln.poissonPm);
    return { lot: ln.lot, pm: pm,
             qty: (kg != null && pm != null && pm > 0) ? kg * 1000 / pm : null };
  });
  const verdict = cmdValidateOrderLines(vLines, f.type);
  if (!verdict.ok) {
    throw new Error("Commande refusée : " + verdict.blocks.join(" ; "));
  }

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
    //
    // Column C (N° facture) is NOT written here. AutoCommandes mints it
    // when a Date livraison is entered, and it skips any order that
    // already holds a value in C — so writing anything at order time
    // would permanently disable the generator for that order. The
    // fulfilment path (cmdRecordFulfilment) still writes C, which is
    // where finance overrides the generated number.
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

  // warnings: from the validation this function ran above. The browser
  // no longer makes its own cmdValidateOrderLines round trip (2026-08-30,
  // it doubled the save time), so this is how PM-mismatch and lot/type
  // warnings reach the screen.
  return { firstRow: firstRow, lastRow: lastRow, orderNumber: orderNumber,
           rowCount: lines.length, warnings: verdict.warnings || [] };
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
 */
/** Parse a fr-formatted display value ("1 500 000", "12,5") into a number; 0 if blank. */
function cmdNumFromDisplay_(s) {
  const t = String(s == null ? "" : s)
    .replace(/[  \s]/g, "")
    .replace(",", ".")
    .replace(/[^\d.-]/g, "");
  const n = Number(t);
  return isFinite(n) ? n : 0;
}

/**
 * Search open orders. An order is DELIVERED when it has a delivery date
 * (col V) and PAID when it has a payment date (col U); "livré" (col W) is
 * a formula from V, so V is the truth. Orders both delivered and paid are
 * closed and never listed. The two flags narrow the rest:
 *   wantDeliveredUnpaid — delivered, not paid
 *   wantUndelivered     — not delivered
 * Neither flag (or both) lists both groups.
 */
function cmdFindOrders(query, wantDeliveredUnpaid, wantUndelivered) {
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
        remarques: r[C.REMARQUES - 1],
        alevinsNb: [],
        poissonKg: [],
        alevinsTotal: 0,
        poissonKgTotal: 0,
        montantAr: 0
      };
      order.push(key);
    }

    const g = groups[key];
    g.rows.push(rowNum);
    if (r[C.LOT - 1]) g.lots.push(r[C.LOT - 1]);
    if (r[C.ALEVINS_NB - 1]) g.alevinsNb.push(r[C.ALEVINS_NB - 1]);
    if (r[C.POISSON_KG - 1]) g.poissonKg.push(r[C.POISSON_KG - 1]);
    g.alevinsTotal += cmdNumFromDisplay_(r[C.ALEVINS_NB - 1]);
    g.poissonKgTotal += cmdNumFromDisplay_(r[C.POISSON_KG - 1]);
    g.montantAr += cmdNumFromDisplay_(r[C.ARGENT_ALEVINS - 1]) +
                   cmdNumFromDisplay_(r[C.ARGENT_POISSON - 1]);
    // Any row carrying fulfilment data represents the order's state.
    if (!g.paiement && r[C.PAIEMENT - 1]) g.paiement = r[C.PAIEMENT - 1];
    if (!g.dateLivraison && r[C.DATE_LIVRAISON - 1]) g.dateLivraison = r[C.DATE_LIVRAISON - 1];
    if (!g.moyenPaiement && r[C.MOYEN_PAIEMENT - 1]) g.moyenPaiement = r[C.MOYEN_PAIEMENT - 1];
    if (!g.facture && r[C.FACTURE - 1]) g.facture = r[C.FACTURE - 1];
    if (!g.bl && r[C.BL - 1]) g.bl = r[C.BL - 1];
    if (!g.remarques && r[C.REMARQUES - 1]) g.remarques = r[C.REMARQUES - 1];
  }

  const out = [];
  let closedMatches = 0;
  for (let i = order.length - 1; i >= 0; i--) {              // newest first
    const g = groups[order[i]];

    if (q) {
      const hay = (g.orderNumber + " " + g.client + " " + g.lots.join(" ")).toLowerCase();
      if (hay.indexOf(q) === -1) continue;
    }
    const delivered = String(g.dateLivraison || "").trim() !== "";
    const paid      = String(g.paiement || "").trim() !== "";

    // Closed. Counted so the screen can say WHY nothing is listed rather
    // than reporting a search that found nothing.
    if (delivered && paid) { closedMatches++; continue; }

    // Each box INCLUDES a category. Fully closed orders (delivered and
    // paid) were already dropped above, so "delivered" here always means
    // delivered-and-unpaid. The two categories are therefore disjoint
    // and together cover every open order: both ticked = the whole list.
    // Nothing ticked selects nothing, and the client says so - it must
    // not silently fall back to showing everything.
    const keep = (wantDeliveredUnpaid && delivered) ||
                 (wantUndelivered && !delivered);
    if (!keep) continue;

    out.push(g);
    if (out.length >= 25) break;
  }
  return { orders: out, closedMatches: closedMatches };
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
    put(C.REMARQUES, f.remarques, "Remarques");
  });

  SpreadsheetApp.flush();

  // Mint the invoice number. THIS IS THE ONLY PLACE IT HAPPENS
  // (2026-08-30, Kim). AutoCommandes no longer mints from onEdit: that
  // made DELIVERY the trigger, so an order delivered and not paid
  // already carried a number, and opening it on this screen looked like
  // the click had invoiced it.
  //
  // THE RULE IS PAYMENT, NOT DELIVERY. A number is minted only when the
  // user records a payment date with the Enregistrer button. Opening an
  // order, reading it and leaving writes nothing, because nothing but
  // this button reaches this function.
  //
  // Both conditions are needed, and each for its own reason:
  //   paiement      — the rule itself. No payment, no invoice.
  //   dateLivraison — acFillFactureForRows skips any row with an empty
  //                   col V, so without it the call would do nothing
  //                   and the "généré" report would be a lie.
  // A prepayment (paid, not yet delivered) is therefore NOT invoiced
  // here. It stays listed as non-livrée and is invoiced on the save
  // that records its delivery.
  //
  // Calls the library's own function, so the numbering rule lives in
  // one place. It skips any order that already holds a value in column
  // C, which makes this safe to call every time, and keeps finance's
  // override (typed above) untouched.
  //
  // AFTER the flush: the function reads Date livraison back from the
  // sheet, so the write must have landed first.
  //
  // Never fatal. The fulfilment is already saved and correct; a missing
  // invoice number is recorded and reported, not raised as an error that
  // would suggest the save failed.
  //
  // THREE fields are returned, not one, so the screen can state the
  // invoice position after EVERY save instead of going quiet:
  //   facture     — the number minted BY THIS SAVE, or null.
  //   factureNow  — the number on the order after this save, minted now
  //                 or already there. null when the order has none.
  //   factureWhy  — why nothing was minted. null when one was.
  // Returned as their own fields, never dug out of `changed`: a message
  // this important must not depend on parsing a log line.
  const anyRow = Math.min.apply(null, targets);

  if (!cmdParseDate(f.paiement) || !cmdParseDate(f.dateLivraison)) {
    return {
      rows: targets, changed: changed, facture: null,
      factureNow: sh.getRange(anyRow, C.FACTURE).getDisplayValue() || null,
      factureWhy: !cmdParseDate(f.paiement)
        ? "la commande n'est pas encore payée"
        : "la commande n'est pas encore livrée"
    };
  }

  var facture = null;
  try {
    // acFillFactureForRows is given ONE row, not the span. The library
    // mints for every row sharing that row's order number, so a single
    // row is enough — while a span would sweep every delivered-but-
    // uninvoiced order that happens to sit between the first and last
    // row of this one, and invoice those too.
    //
    // Before/after, so a number finance typed above is not reported as
    // "généré". The library leaves such an order alone, and this says so.
    const before = targets.map(r => sh.getRange(r, C.FACTURE).getDisplayValue());
    AutoCommandes.acFillFactureForRows(sh, anyRow, anyRow);
    SpreadsheetApp.flush();
    targets.forEach((r, i) => {
      const after = sh.getRange(r, C.FACTURE).getDisplayValue();
      if (after && after !== before[i]) {
        changed.push("L" + r + " N° facture (généré): " + after);
        facture = after;
      }
    });
  } catch (err) {
    changed.push("N° facture non généré : " + err + ". Prévenir Kim.");
  }

  // Read back, so factureNow reports the sheet and not what this
  // function hoped it wrote.
  const factureNow = sh.getRange(anyRow, C.FACTURE).getDisplayValue() || null;

  return {
    rows: targets, changed: changed, facture: facture,
    factureNow: factureNow,
    factureWhy: facture ? null
      : (factureNow ? "cette commande a déjà un numéro"
                    : "numéro non généré — prévenir Kim")
  };
}

/**
 * Mark every row of an order cancelled (column AA = "x").
 *
 * One way only, by Kim's rule (2026-08-15): a cancelled order can never
 * be reactivated. cmdFindOrders skips cancelled rows, so the order
 * disappears from the search as soon as this returns.
 *
 * Stock is not touched here. The engine settles it on its next run: an
 * order it had already deducted is re-credited to the lot exactly once
 * (guarded against a second credit), and an order it never processed
 * simply becomes ineligible, so nothing moves.
 */
function cmdCancelOrder(rows) {
  const sh = cmdSheet();
  const C = CMD_CFG.COL;
  const lastRow = findNextCommandeRow(sh) - 1;

  const targets = (rows || []).map(Number).filter(function (r) {
    return isFinite(r) && r >= CMD_CFG.START_ROW && r <= lastRow;
  });
  if (!targets.length) throw new Error("Aucune ligne de commande valide \u00e0 annuler.");

  const already = [];
  targets.forEach(function (r) {
    if (String(sh.getRange(r, C.ANNULE).getDisplayValue() || "").trim() !== "") {
      already.push(r);
    }
  });
  if (already.length === targets.length) {
    throw new Error("Cette commande est d\u00e9j\u00e0 annul\u00e9e.");
  }

  targets.forEach(function (r) { sh.getRange(r, C.ANNULE).setValue("x"); });
  SpreadsheetApp.flush();
  return { rows: targets, alreadyCancelled: already };
}

/***************************************************************
 * MODIFIER LA COMMANDE (2026-08-31)
 *
 * Editable fields, per Kim: alevins -> nombre (F) + PM (G);
 * grossis -> kg (L) + PM (M). Nothing else. Money is NEVER written:
 * K, N, Q, W are sheet formulas (verified in the formula bar
 * 2026-08-31: K3 =(F3*I3)+J3, N3 =IFERROR((L3*1000)/M3;0),
 * Q3 =(O3*L3)+P3, W3 =IF(V3<>"";"x";"")), so the totals recompute
 * on their own. Writing K or Q would kill the formula for good.
 *
 * GATE: every row must have V (Date livraison) empty and AA empty.
 * Delivered or cancelled orders are refused.
 *
 * H IS REWRITTEN. Alevins "à livrer" (H) is a typed value, not a
 * formula; the browser seeds it as round(F*1.05) at order time. On
 * modify it is ALWAYS reset to round(newF*1.05) — a manual H
 * override typed at order time does not survive a modify. One rule.
 *
 * STOCK, at click time, not at night:
 *   Y empty        -> nothing to correct. The engine deducts the NEW
 *                     quantity tonight (H/N is the new value). If Z
 *                     holds a sticky error the row stays skipped until
 *                     Kim clears Z — reported, not touched.
 *   Y has [qty=X]  -> the engine deducted X. Apply (X - newDed) to the
 *                     exact cell the engine used (same resolver:
 *                     findSubLotColumnByOrderKey), then rewrite the
 *                     stamp's [qty=] IN PLACE to newDed and append an
 *                     "Ajusté" note.
 *
 * WHY THE STAMP IS REWRITTEN IN PLACE: tt_recreditQty in
 * engine_core.js reads the FIRST [qty=] with a non-global regex. A
 * second [qty=] token would make a later cancellation re-credit the
 * PRE-edit quantity. The stamp must always carry the amount currently
 * deducted. tsaracockpit matches "Processed on" with indexOf >= 0,
 * so the appended note is safe there.
 *
 * WRITE ORDER mirrors the engine: lot-file stock first, stamp after,
 * each flushed. Two spreadsheets, no transaction — a failure between
 * the two writes leaves an error THE SIZE OF THE EDIT, exactly the
 * exposure the nightly deduction already carries. There is no nightly
 * retry pass, deliberately: one path (Kim's rule). A failure is
 * reported on screen with "Prévenir Kim".
 *
 * STOCK GUARD on increases only: the DELTA is validated with
 * cmdValidateOrderLines, same verdict as order creation. Delta is
 * correct in both states: Y empty -> the old quantity already sits in
 * cmdGetPendingQty; Y stamped -> the lot count is already reduced by
 * the old quantity. Reductions are never blocked — lines with
 * delta <= 0 are not sent to the validator, because its found/reserved
 * checks fire before the quantity check and would block a legitimate
 * reduction on a lot that has since been reserved.
 ***************************************************************/

/** Per-row editable data for one order. rows comes from cmdFindOrders. */
function cmdGetOrderLines(rows) {
  const sh = cmdSheet();
  const C = CMD_CFG.COL;
  const lastRow = findNextCommandeRow(sh) - 1;
  const targets = (rows || []).map(Number).filter(function (r) {
    return isFinite(r) && r >= CMD_CFG.START_ROW && r <= lastRow;
  });
  if (!targets.length) throw new Error("Aucune ligne de commande valide.");

  return targets.map(function (r) {
    const v = sh.getRange(r, 1, 1, C.ANNULE).getValues()[0];
    const isAl = cmdToNum(v[C.ALEVINS_NB - 1]) != null;
    return {
      row: r,
      lot: String(v[C.LOT - 1] || ""),
      isAlevins: isAl,
      nombre: cmdToNum(v[C.ALEVINS_NB - 1]),
      pmAl: cmdToNum(v[C.ALEVINS_PM - 1]),
      kg: cmdToNum(v[C.POISSON_KG - 1]),
      pmGr: cmdToNum(v[C.POISSON_PM - 1])
    };
  });
}

/** payload.lines = [{row, nombre, pm}] (alevins) or [{row, kg, pm}]. */
function cmdModifyOrder(payload) {
  const f = payload || {};
  const lines = f.lines || [];
  if (!lines.length) throw new Error("Aucune modification re\u00e7ue.");

  const sh = cmdSheet();
  const C = CMD_CFG.COL;
  const lastRow = findNextCommandeRow(sh) - 1;

  // ---- read current state, gate, plan the jobs ----
  const jobs = [];
  var isAlevinsOrder = null;
  lines.forEach(function (ln) {
    const r = Number(ln.row);
    if (!isFinite(r) || r < CMD_CFG.START_ROW || r > lastRow) {
      throw new Error("Ligne " + ln.row + " invalide.");
    }
    const v = sh.getRange(r, 1, 1, C.ANNULE).getValues()[0];
    if (String(v[C.ANNULE - 1] || "").trim() !== "") {
      throw new Error("Commande annul\u00e9e \u2014 modification impossible.");
    }
    if (String(v[C.DATE_LIVRAISON - 1] || "").trim() !== "") {
      throw new Error("Commande d\u00e9j\u00e0 livr\u00e9e \u2014 modification impossible.");
    }

    const key = cmdCanonKey(v[C.LOT - 1]);
    if (!key) throw new Error("Ligne " + r + " : cl\u00e9 de lot vide.");
    const y = String(v[C.LOG - 1] || "");
    const m = /\[qty=([0-9][0-9.,]*)\]/.exec(y);
    const stampQty = m ? cmdToNum(m[1]) : null;
    const isAl = cmdToNum(v[C.ALEVINS_NB - 1]) != null;
    if (isAlevinsOrder === null) isAlevinsOrder = isAl;

    const pm = cmdToNum(ln.pm);
    if (pm == null || pm <= 0) throw new Error("Ligne " + r + " : PM invalide.");

    var job = {
      row: r, key: key, isAl: isAl, pm: pm, y: y, stampQty: stampQty,
      zFilled: String(v[C.ERROR - 1] || "").trim() !== ""
    };
    if (isAl) {
      const nombre = cmdToNum(ln.nombre);
      if (nombre == null || nombre <= 0) throw new Error("Ligne " + r + " : nombre invalide.");
      job.nombre = nombre;
      job.newDed = Math.round(nombre * 1.05);
      job.oldDed = cmdDeduction(v[C.ALEVINS_LIVRER - 1], null);
    } else {
      const kg = cmdToNum(ln.kg);
      if (kg == null || kg <= 0) throw new Error("Ligne " + r + " : kg invalide.");
      job.kg = kg;
      job.newDed = kg * 1000 / pm;
      job.oldDed = cmdDeduction(null, v[C.POISSON_NB - 1]);
    }
    jobs.push(job);
  });

  // ---- stock guard: validate the INCREASE only ----
  const incLines = jobs
    .map(function (j) {
      const base = (j.stampQty != null) ? j.stampQty : (j.oldDed || 0);
      return { lot: j.key, qty: j.newDed - base, pm: j.pm };
    })
    .filter(function (l) { return l.qty > 0; });
  var warnings = [];
  if (incLines.length) {
    const verdict = cmdValidateOrderLines(incLines, isAlevinsOrder ? "AL" : "GR");
    if (!verdict.ok) {
      throw new Error("Modification refus\u00e9e : " + verdict.blocks.join(" ; "));
    }
    warnings = verdict.warnings || [];
  }

  // ---- write the new order values ----
  const changed = [];
  jobs.forEach(function (j) {
    function put(col, value, label) {
      const cell = sh.getRange(j.row, col);
      const before = cell.getDisplayValue();
      cell.setValue(value);
      const after = cell.getDisplayValue();
      if (before !== after) {
        changed.push("L" + j.row + " " + label + ": " + (before || "(vide)") + " -> " + after);
      }
    }
    if (j.isAl) {
      put(C.ALEVINS_NB, j.nombre, "Nombre alevins");
      put(C.ALEVINS_PM, j.pm, "PM");
      put(C.ALEVINS_LIVRER, j.newDed, "\u00c0 livrer (+5%)");
    } else {
      put(C.POISSON_KG, j.kg, "Kg poisson");
      put(C.POISSON_PM, j.pm, "PM");
    }
  });
  SpreadsheetApp.flush();

  // ---- stock reconcile, per row, stamped rows only ----
  const stock = [];
  jobs.forEach(function (j) {
    if (j.stampQty == null) {
      stock.push(j.key + " : pas encore d\u00e9duit \u2014 le moteur d\u00e9duira la nouvelle quantit\u00e9 cette nuit.");
      if (j.zFilled) {
        stock.push("\u26a0 " + j.key + " : la colonne Z contient une erreur \u2014 le moteur ignorera cette ligne tant que Kim ne l'a pas effac\u00e9e.");
      }
      return;
    }

    // Grossis: N is a formula — read it back so the stamp and the
    // stock move carry EXACTLY what the sheet now says, not what this
    // function computed. Alevins: H was just written, newDed is exact.
    var newDed = j.newDed;
    if (!j.isAl) {
      const nBack = cmdToNum(sh.getRange(j.row, C.POISSON_NB).getValue());
      if (nBack != null && nBack > 0) newDed = nBack;
    }
    const adjust = j.stampQty - newDed;
    if (Math.abs(adjust) < 0.001) {
      stock.push(j.key + " : quantit\u00e9 d\u00e9duite inchang\u00e9e.");
      return;
    }

    // Locate the exact cell the engine deducted from.
    const lotNum = j.key.split("-")[0];
    const list = getLotFileList();
    var fileId = null;
    for (var i = 0; i < list.length; i++) {
      if (cmdCanonKey(list[i].lotNumber) === lotNum) { fileId = list[i].fileId; break; }
    }
    if (!fileId) {
      stock.push("\u26a0 " + j.key + " : fichier lot introuvable \u2014 stock NON ajust\u00e9. Pr\u00e9venir Kim.");
      return;
    }
    const lotSS = SpreadsheetApp.openById(fileId);
    const hit = findSubLotColumnByOrderKey(lotSS, j.key);
    if (!hit.found) {
      stock.push("\u26a0 " + j.key + " : cl\u00e9 introuvable dans le fichier lot \u2014 stock NON ajust\u00e9. Pr\u00e9venir Kim.");
      return;
    }

    // Stock first, stamp after — same order as the engine.
    const newCount = (Number(hit.count) || 0) + adjust;
    if (hit.source === "GROSS") {
      lotSS.getSheetByName(LOT_CFG.GROSS_SHEET)
           .getRange(LOT_CFG.GROSS_ROW_NOMBRE, hit.col).setValue(newCount);
    } else {
      lotSS.getSheetByName(hit.source).getRange(hit.row, 2).setValue(newCount);
    }
    SpreadsheetApp.flush();

    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(),
                                       "yyyy-MM-dd HH:mm");
    const newY = j.y.replace(/\[qty=[0-9][0-9.,]*\]/, "[qty=" + newDed + "]") +
                 " | Ajust\u00e9 le " + stamp +
                 " (" + j.stampQty + " -> " + newDed + ")";
    sh.getRange(j.row, C.LOG).setValue(newY);
    SpreadsheetApp.flush();

    stock.push(j.key + " : stock " +
      (adjust > 0 ? "recr\u00e9dit\u00e9 de " + Math.round(adjust)
                  : "d\u00e9duit de " + Math.round(-adjust) + " en plus") +
      " poisson(s) dans le fichier lot.");
  });

  return { changed: changed, stock: stock, warnings: warnings };
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
  // NOT getLastRow(): it reports ~2041 here against ~180 real orders,
  // because K, N, Q and W carry formulas far below the data. Reading to
  // it costs ~55 000 mostly-empty cells to sum about a dozen values.
  // findNextCommandeRow scans column A for the real end - the same
  // workaround cmdCreateOrder already depends on.
  const lastRow = findNextCommandeRow(sh) - 1;
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

/* =============================================================
 * LOT / TYPE GATE (2026-08-12)
 *
 * A fry order against a grow-out lot, or a fish order against a fry lot,
 * is a category error the engine cannot catch: it deducts whatever the
 * row asks for. So it is caught here.
 *
 * BOUNDARIES — where each number comes from:
 *   AL max  = top of the fry grid, READ FROM the Tarifs sheet (20 g
 *             today). Not a constant: raise the grid and this follows.
 *   GR block = 149 g. Below this the farm does not sell as grossis.
 *   GR warn  = 200 g. Between 149 and 200 the sale is allowed but
 *             flagged — a commercial judgement, not an error, so it
 *             must not cost a sale.
 *
 * WHY NOT 350 g: 350 is the production model's target harvest weight,
 * not a commercial floor. On 2026-08-12 the heaviest live lot was
 * 246 g and the Tarifs fish grid prices a "< 300 g" band — a 350 g
 * floor would have blocked 100% of Grossis orders. Checked against
 * live Stock Poisson data before this was written.
 * ============================================================= */

var CMD_GR_BLOCK_PM = 149;   // below this: refuse a grossis order
var CMD_GR_WARN_PM  = 200;   // below this: allow, but flag it

/** Top of the fry grid, from Tarifs. Never hardcode this. */
function cmdFryMaxPm() {
  return cmdGetTarifs().bands.slice(-1)[0].max;
}

/**
 * Verdict for one lot PM against one order type.
 * @return {Object} { level: "ok"|"warn"|"block", msg: string }
 * Pure function of its inputs — unit-testable, no sheet access.
 */
function cmdLotTypeVerdict(isAlevins, pm, fryMax) {
  if (pm == null || !isFinite(pm)) return { level: "ok", msg: "" };
  if (isAlevins) {
    if (pm > fryMax) {
      return { level: "block", msg: "PM du lot " + pm + " g > " + fryMax +
        " g — ce n'est pas un lot alevins" };
    }
    return { level: "ok", msg: "" };
  }
  if (pm < CMD_GR_BLOCK_PM) {
    return { level: "block", msg: "PM du lot " + pm + " g < " + CMD_GR_BLOCK_PM +
      " g — trop petit pour être vendu en grossis" };
  }
  if (pm < CMD_GR_WARN_PM) {
    return { level: "warn", msg: "PM du lot " + pm + " g — en dessous de " +
      CMD_GR_WARN_PM + " g ; vente possible mais à confirmer" };
  }
  return { level: "ok", msg: "" };
}

var LOT_PM_CACHE_KEY = "cmd_lot_pm_v1";
var LOT_PM_CACHE_SECONDS = 300;

/**
 * { canonKey: PM } for every lot in Stock Poisson, cached 5 min.
 *
 * Source: Stock Poisson "lot" tab, N = lot id, P = PM, from row 3 —
 * the block updateStockPoisson writes each night (engine_core.js
 * writes columns 14..21 from startRow 3). One read of one sheet; no
 * lot files are opened, so this costs nothing next to the existing
 * per-selection availability call.
 *
 * This drives the DROPDOWN only. A lot missing here is left selectable
 * and falls through to the normal per-lot check — a gap in Stock
 * Poisson must never make a real lot silently disappear from the list.
 */
function cmdGetLotPmMap() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(LOT_PM_CACHE_KEY);
  if (hit) return JSON.parse(hit);
  const out = buildLotPmMap();
  cache.put(LOT_PM_CACHE_KEY, JSON.stringify(out), LOT_PM_CACHE_SECONDS);
  return out;
}

/** Drop the cached PM map. */
function clearLotPmMapCache() {
  CacheService.getScriptCache().remove(LOT_PM_CACHE_KEY);
  Logger.log("lot PM map cache cleared");
}

/** The real read. Call cmdGetLotPmMap instead. */
function buildLotPmMap() {
  const ss = SpreadsheetApp.openById(STOCK_PM_CFG.SS_ID);
  const sh = ss.getSheetByName(STOCK_PM_CFG.SHEET);
  if (!sh) throw new Error('Onglet introuvable: "' + STOCK_PM_CFG.SHEET + '"');
  const lastRow = sh.getLastRow();
  const out = {};
  if (lastRow < STOCK_PM_CFG.START_ROW) return out;

  // N..P = lot id, nombre, PM
  const vals = sh.getRange(STOCK_PM_CFG.START_ROW, 14,
                           lastRow - STOCK_PM_CFG.START_ROW + 1, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    const key = cmdCanonKey(vals[i][0]);
    if (!key) continue;
    const pm = cmdToNum(vals[i][2]);
    if (pm != null) out[key] = pm;
  }
  return out;
}

/**
 * Every TOUT reservation, in one read: { canonKey: true }.
 *
 * Only TOUT is returned. A NOMBRE reservation does not make a lot
 * unorderable - it lowers the available count, which the per-lot
 * readout already shows and cmdValidateOrderLines already enforces.
 * Greying those out would refuse sales the farm can legitimately make.
 *
 * NOT cached. The Réservations tab is a handful of rows, and a stale
 * reservation is the one kind of staleness that could let a blocked
 * order be typed. Read fresh, once per page load.
 */
function buildReservedAllMap() {
  const ss = SpreadsheetApp.openById(CMD_CFG.SS_ID);
  const sh = ss.getSheetByName("Réservations");
  const out = {};
  if (!sh) return out;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return out;

  const vals = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    const key = cmdCanonKey(vals[i][0]);
    if (!key) continue;
    if (String(vals[i][1] || "").trim().toUpperCase() === "TOUT") out[key] = true;
  }
  return out;
}

/* -------------------------------------------------------------
 * "Pas encore vendable" map
 *
 * A sub-lot whose stock cell in the LOT FILE is empty cannot be sold —
 * fry before the tri are the usual case. Stock Poisson cannot tell this
 * apart from a sellable lot (it still shows the fish), so the only
 * source is the lot files themselves, and the dropdown is built long
 * before any lot file is opened.
 *
 * Hence: build it ahead of time, store it, serve it instantly.
 * Rebuilt nightly by refreshNotSellableMap, after the engine has
 * finished writing the lot files. Kim runs the same function by hand
 * after a tri, so a lot that became sellable during the day does not
 * stay greyed until the next night.
 *
 * ScriptProperties, not CacheService: an expiry here would silently
 * empty the map and quietly un-grey every lot.
 *
 * A lot MISSING from this map is left selectable and falls through to
 * the per-selection check, which blocks it anyway — same rule as the PM
 * map. A gap must never hide a real lot.
 * ------------------------------------------------------------- */

const NOT_SELLABLE_PROP_KEY = "cmd_not_sellable_v1";

/** Nightly trigger hour, script timezone. The engine finishes ~04:53. */
const NOT_SELLABLE_TRIGGER_HOUR = 6;

/** Stored map. Empty until refreshNotSellableMap has run once. */
function getNotSellableMap() {
  const raw = PropertiesService.getScriptProperties().getProperty(NOT_SELLABLE_PROP_KEY);
  return raw ? JSON.parse(raw) : {};
}

/**
 * The real scan. Opens each LOT FILE ONCE and resolves every key that
 * belongs to it, instead of calling cmdGetLotAvailability per key —
 * that reopens the lot file AND the Commandes file (reservations,
 * pending) for all 30 keys, which is what made the per-key measurement
 * 1 min 46 s. Reservations and pending quantities are irrelevant here:
 * this map answers "is there any stock at all", not "how much is left".
 *
 * Resolution itself goes through findSubLotColumnByOrderKey, the same
 * function the per-lot check and the engine mirror use. Do not inline a
 * second copy of that scan — two copies of the rule would drift apart.
 */
function buildNotSellableMap() {
  const list = getLotFileList();
  const byLot = {};
  Object.keys(cmdGetLotPmMap()).forEach(function (k) {
    const lotNum = k.split("-")[0];
    if (!byLot[lotNum]) byLot[lotNum] = [];
    byLot[lotNum].push(k);
  });

  const out = {};
  Object.keys(byLot).forEach(function (lotNum) {
    var fileId = null;
    for (var i = 0; i < list.length; i++) {
      if (cmdCanonKey(list[i].lotNumber) === lotNum) { fileId = list[i].fileId; break; }
    }
    if (!fileId) return;                       // no lot file -> stays selectable
    const ss = SpreadsheetApp.openById(fileId);
    byLot[lotNum].forEach(function (k) {
      const m = findSubLotColumnByOrderKey(ss, k);
      if (m.found && !m.count) out[k] = true;  // found, but the cell is empty
    });
  });
  return out;
}

/**
 * RUN FROM EDITOR after a tri, and nightly by trigger. Rebuilds and
 * stores the map. Writes only to ScriptProperties — no spreadsheet is
 * touched.
 */
function refreshNotSellableMap() {
  const started = new Date();
  const map = buildNotSellableMap();
  PropertiesService.getScriptProperties()
    .setProperty(NOT_SELLABLE_PROP_KEY, JSON.stringify(map));
  const ks = Object.keys(map).sort();
  Logger.log("Carte reconstruite en " + Math.round((new Date() - started) / 1000) + " s");
  Logger.log("Pas encore vendable (" + ks.length + ") : " + (ks.join(", ") || "aucun"));
  return map;
}

/**
 * RUN FROM EDITOR ONCE. Installs the nightly rebuild. Deletes any
 * existing trigger for the same handler first, so running it twice
 * cannot leave two triggers rebuilding the same map.
 */
function installNotSellableTrigger() {
  const existing = ScriptApp.getProjectTriggers();
  var removed = 0;
  existing.forEach(function (t) {
    if (t.getHandlerFunction() === "refreshNotSellableMap") {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  ScriptApp.newTrigger("refreshNotSellableMap")
    .timeBased()
    .everyDays(1)
    .atHour(NOT_SELLABLE_TRIGGER_HOUR)
    .create();
  Logger.log("Anciens déclencheurs supprimés : " + removed);
  Logger.log("Déclencheur quotidien installé vers " + NOT_SELLABLE_TRIGGER_HOUR + " h.");
}

/**
 * {canonKey: available fish} for the lot dropdown.
 * available = Stock Poisson count - NOMBRE reservation - orders
 * entered but not yet deducted. The same arithmetic the save check
 * enforces, so the number in the label is the number that will be
 * accepted. TOUT lots are omitted - the dropdown already greys them
 * via reservedAll, and a count would contradict the "reserve" label.
 * Advisory: Stock Poisson can lag the lot files by up to 24 h; the
 * per-selection readout and the save check remain the authority.
 *
 * 2026-09-02: restored. Deleted by accident in 94a9b9f while its
 * caller cmdGetLotGateData stayed. demCheckStock holds a second copy
 * of this arithmetic inline, with different filters (it skips
 * notSellable and non-positive stock, the dropdown must not). Merge
 * the two only if the filters are made explicit parameters.
 */
function buildAvailMap() {
  const ss = SpreadsheetApp.openById(STOCK_PM_CFG.SS_ID);
  const sh = ss.getSheetByName(STOCK_PM_CFG.SHEET);
  const out = {};
  if (!sh) return out;
  const lastRow = sh.getLastRow();
  if (lastRow < STOCK_PM_CFG.START_ROW) return out;
  const vals = sh.getRange(STOCK_PM_CFG.START_ROW, 14,
                           lastRow - STOCK_PM_CFG.START_ROW + 1, 3).getValues();
  const resv = demResMap();
  const pend = demPendingMap();
  for (var i = 0; i < vals.length; i++) {
    const key = cmdCanonKey(vals[i][0]);
    if (!key) continue;
    if (resv[key] === "TOUT") continue;
    const count = cmdToNum(vals[i][1]);
    if (count == null) continue;
    out[key] = Math.round(count - (resv[key] || 0) - (pend[key] || 0));
  }
  return out;
}

/**
 * Everything the Commandes screen needs to gate its lot dropdown:
 * the PM map, the TOUT reservations, the sub-lots with no stock, and
 * the boundaries - so the client never hardcodes any of them.
 */
function cmdGetLotGateData() {
  return {
    pm: cmdGetLotPmMap(),
    avail: buildAvailMap(),
    reservedAll: buildReservedAllMap(),
    notSellable: getNotSellableMap(),
    fryMax: cmdFryMaxPm(),
    grBlock: CMD_GR_BLOCK_PM,
    grWarn: CMD_GR_WARN_PM
  };
}

/** RUN FROM EDITOR: which lots the !found block would refuse. Writes nothing. */
function testFoundBlock() {
  const keys = Object.keys(cmdGetLotGateData().pm).sort();
  const blocked = [];
  keys.forEach(function (k) {
    let a;
    try { a = cmdGetLotAvailability(k); }
    catch (e) { Logger.log(k + "  ERREUR : " + e.message); return; }
    if (a.reservedAll) { Logger.log(k + "  [RESERVE TOUT]"); return; }
    if (!a.found) { blocked.push(k); Logger.log(k + "  >>> BLOQUE (found=false)"); return; }
    Logger.log(k + "  ok  dispo=" + a.available);
  });
  Logger.log("total=" + keys.length + "  bloques=" + blocked.length +
             "  : " + (blocked.join(", ") || "aucun"));
}

/**
 * RUN FROM EDITOR: which lots the client's "pas encore vendable" block
 * will catch, and why. Writes nothing. The client blocks on count == 0
 * (the engine's own deduct cell), so this prints count next to reserved,
 * pending and available — if a lot is caught for the wrong reason, the
 * numbers say so.
 */
function testNotSellable() {
  const keys = Object.keys(cmdGetLotGateData().pm).sort();
  const blocked = [];
  keys.forEach(function (k) {
    let a;
    try { a = cmdGetLotAvailability(k); }
    catch (e) { Logger.log(k + "  ERREUR : " + e.message); return; }
    if (a.reservedAll) { Logger.log(k + "  [RESERVE TOUT]"); return; }
    if (!a.found) { Logger.log(k + "  [NON TROUVE]"); return; }
    const line = k + "  count=" + a.count + "  reserve=" + a.reserved +
                 "  pending=" + a.pending + "  dispo=" + a.available +
                 "  pm=" + a.pm + "  source=" + a.source;
    if (!a.count) { blocked.push(k); Logger.log(line + "   >>> PAS ENCORE VENDABLE"); }
    else Logger.log(line);
  });
  Logger.log("total=" + keys.length + "  bloques=" + blocked.length +
             "  : " + (blocked.join(", ") || "aucun"));
}

/** RUN FROM EDITOR: what the dropdown will allow, per lot. */
function testLotGate() {
  const d = cmdGetLotGateData();
  Logger.log("fryMax=" + d.fryMax + "  grBlock=" + d.grBlock + "  grWarn=" + d.grWarn);
  const res = Object.keys(d.reservedAll);
  Logger.log("réservés TOUT (" + res.length + ") : " + (res.join(", ") || "aucun"));
  Object.keys(d.pm).sort().forEach(function (k) {
    const pm = d.pm[k];
    const al = cmdLotTypeVerdict(true, pm, d.fryMax).level;
    const gr = cmdLotTypeVerdict(false, pm, d.fryMax).level;
    Logger.log(k + "  PM=" + pm + "   AL:" + al + "   GR:" + gr +
               (d.reservedAll[k] ? "   [RÉSERVÉ]" : ""));
  });
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
 * @param {string=} orderType  raw type ("AL"/"GR"). Enables the lot/type
 *        gate. The client shows the same verdict at selection time;
 *        this is the ENFORCING copy, because only the server runs at save.
 * @return {Object} { ok, blocks: [msg], warnings: [msg], detail: {key: avail} }
 */
/* =============================================================
 * TARIFS — price list lookup for Prix alevin / Prix / kg (2026-08-12)
 *
 * Reads Tarifs!B2:G37 once, cached 5 min (same pattern as
 * getLotFileList / buildLotFileList in Lot.js). The client fetches
 * this ONCE at boot and does every lookup in the browser — no
 * per-line server round trip.
 *
 * Fry grid: 5 weight bands (rows 3, 8, 12, 17, 22), each occupying
 * 4 rows:
 *   r+0  band label (B) + column headers
 *   r+1  "Avec provende" label (B) ; C = avec/enlevement price ;
 *        D..G are still sub-headers ("< 5000 pieces" etc.)
 *   r+2  avec provende prices, D..G
 *   r+3  "Sans provende" label (B) ; C..G = sans provende prices
 * Verified live 2026-08-12 against all 5 bands.
 *
 * Fish: the block is found by the LABEL "Prix poisson grossi" in
 * column B. Label row r, headers Detail/Gros at r+1, the two
 * MGA/kg prices at r+2. Never pin this to a literal row number.
 * Rows 27-32 (poisson frais sur glace) were deleted 2026-08-31;
 * the old hardcoded constant then read empty cells and returned 0,
 * which under-billed every GR order in silence.
 * ============================================================= */

const TARIFS_CFG = {
  SS_ID: CMD_CFG.SS_ID,
  SHEET: "Tarifs",
  BAND_START_ROWS: [3, 8, 12, 17, 22],
  FISH_LABEL: "Prix poisson grossi",
  FISH_SCAN_ROWS: 60
};

/* Stock Poisson, "lot" tab: the nightly output of updateStockPoisson
   (engine_core.js). N = lot id, O = nombre, P = PM, data from row 3. */
const STOCK_PM_CFG = {
  SS_ID: "1Kfs5beQorhdheqzEDibgnBd5wQjy79MecncNlgKjRIE",
  SHEET: "lot",
  START_ROW: 3
};

const TARIFS_CACHE_KEY = "cmd_tarifs_v1";
const TARIFS_CACHE_SECONDS = 300;

function cmdGetTarifs() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(TARIFS_CACHE_KEY);
  if (hit) return JSON.parse(hit);

  const out = buildTarifs();
  cache.put(TARIFS_CACHE_KEY, JSON.stringify(out), TARIFS_CACHE_SECONDS);
  return out;
}

/** Drop the cached price list. Run after editing Tarifs to see it now. */
function clearTarifsCache() {
  CacheService.getScriptCache().remove(TARIFS_CACHE_KEY);
  Logger.log("tarifs cache cleared");
}

/** The real read of the Tarifs tab. Always hits Sheets - call cmdGetTarifs instead. */
function buildTarifs() {
  const ss = SpreadsheetApp.openById(TARIFS_CFG.SS_ID);
  const sh = ss.getSheetByName(TARIFS_CFG.SHEET);
  if (!sh) throw new Error('Onglet introuvable: "' + TARIFS_CFG.SHEET + '"');

  const bands = TARIFS_CFG.BAND_START_ROWS.map(function (r0) {
    const label = String(sh.getRange(r0, 2).getValue() || "").trim();
    const m = label.match(/([\d.,]+)\s*-\s*([\d.,]+)/);
    if (!m) throw new Error('Bande Tarifs illisible ligne ' + r0 + ': "' + label + '"');

    const avecEnlevement = Number(sh.getRange(r0 + 1, 3).getValue()) || 0;   // C{r0+1}
    const avecRow = sh.getRange(r0 + 2, 4, 1, 4).getValues()[0];            // D..G {r0+2}
    const sansRow = sh.getRange(r0 + 3, 3, 1, 5).getValues()[0];            // C..G {r0+3}

    return {
      min: Number(m[1].replace(",", ".")),
      max: Number(m[2].replace(",", ".")),
      avec: {
        enlevement: avecEnlevement,
        environsLe: Number(avecRow[0]) || 0,
        environsGt: Number(avecRow[1]) || 0,
        ambohimLe:  Number(avecRow[2]) || 0,
        ambohimGt:  Number(avecRow[3]) || 0
      },
      sans: {
        enlevement: Number(sansRow[0]) || 0,
        environsLe: Number(sansRow[1]) || 0,
        environsGt: Number(sansRow[2]) || 0,
        ambohimLe:  Number(sansRow[3]) || 0,
        ambohimGt:  Number(sansRow[4]) || 0
      }
    };
  });

  return {
    bands: bands,
    fish: tarifsFishPrices(sh)
  };
}

/**
 * Grossis Detail/Gros prices (MGA/kg), found by LABEL in column B.
 *
 * Layout:  r    "Prix poisson grossi"
 *          r+1  "Detail" | "Gros"      (headers)
 *          r+2   <detail> | <gros>     (the prices)
 *
 * Throws rather than returning 0. A 0 here is not a missing price,
 * it is a wrong invoice: the client leaves Prix/kg blank, put()
 * skips the null, and Q = (O*L)+P books the order at the frais
 * only. Fail loudly instead.
 *
 * @param {Sheet} sh  the Tarifs sheet
 * @return {{detail: number, gros: number}}
 */
function tarifsFishPrices(sh) {
  const colB = sh.getRange(1, 2, TARIFS_CFG.FISH_SCAN_ROWS, 1).getValues();
  const want = TARIFS_CFG.FISH_LABEL.toLowerCase();
  let r = 0;
  for (let i = 0; i < colB.length; i++) {
    if (String(colB[i][0] || "").trim().toLowerCase() === want) {
      r = i + 1;   // getValues is 0-based, sheet rows are 1-based
      break;
    }
  }
  if (!r) {
    throw new Error('Tarifs : libelle "' + TARIFS_CFG.FISH_LABEL +
                    '" introuvable en colonne B (lignes 1-' +
                    TARIFS_CFG.FISH_SCAN_ROWS + ')');
  }

  const row = sh.getRange(r + 2, 2, 1, 2).getValues()[0];
  const detail = Number(row[0]) || 0;
  const gros = Number(row[1]) || 0;
  if (!detail || !gros) {
    throw new Error("Tarifs : prix poisson grossi vide ligne " + (r + 2) +
                    " (Detail=" + row[0] + ", Gros=" + row[1] + ")");
  }
  return { detail: detail, gros: gros };
}

/** RUN FROM EDITOR: sanity-check the parsed price list. */
function testTarifs() {
  const t = buildTarifs();
  Logger.log(JSON.stringify(t, null, 2));
}

function cmdValidateOrderLines(lines, orderType) {
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
      blocks.push(k + " : ce lot n'est pas encore vendable — le stock " +
                      "n'est pas trouvé dans le fichier lot");
      return;
    }
    // ---- lot/type gate ----
    // PM here is the LOT FILE's PM (same source the engine deducts from),
    // not Stock Poisson's. The dropdown uses Stock Poisson because it is
    // one cheap read for all lots; this check uses the authoritative one.
    if (orderType) {
      const isAl = String(orderType).toUpperCase().indexOf("AL") === 0;
      const v = cmdLotTypeVerdict(isAl, a.pm, cmdFryMaxPm());
      if (v.level === "block") { blocks.push(k + " : " + v.msg); return; }
      if (v.level === "warn") warnings.push(k + " : " + v.msg);
    }
    // strict > : ordering exactly down to the reservation floor is legal,
    // because the engine blocks on next < reserved, not next <= reserved.
    if (wanted[k] > a.available) {
      const isAlType = orderType &&
                       String(orderType).toUpperCase().indexOf("AL") === 0;
      // Grossis only: the worker typed kg, so quote kg back. Rounded down.
      const grKg = (!isAlType && orderType &&
                    a.pm != null && a.available != null)
        ? " (" + Math.floor(a.available * a.pm / 1000) + " kg)" : "";
      // Tell the worker exactly how much to remove. Fish rounded UP
      // (ceil), kg rounded UP: reducing by the stated amount must always
      // bring the order back under the floor — never understate it.
      const excess = Math.ceil(wanted[k] - a.available);
      const reduire = (!isAlType && orderType && a.pm != null)
        ? " → réduire de " + Math.ceil(excess * a.pm / 1000) + " kg"
        : " → réduire de " + excess + " alevins";
      blocks.push(k + " : stock insuffisant — demandé " + Math.round(wanted[k]) +
                      ", disponible " + a.available + grKg +
                      " (stock " + a.count +
                      (a.reserved ? ", réservé " + a.reserved : "") +
                      (a.pending ? ", en attente " + a.pending : "") + ")" +
                      reduire);
    }
    if (pmSeen[k] != null && a.pm != null && pmSeen[k] !== a.pm) {
      warnings.push(k + " : PM saisi " + pmSeen[k] + " ≠ PM du lot " + a.pm);
    }
  });

  return { ok: blocks.length === 0, blocks: blocks, warnings: warnings, detail: detail };
}

/** RUN FROM EDITOR: prove the !found block refuses an order. Writes nothing.
 *  Case 1 must be ok=false with a "pas encore vendable" block.
 *  Case 2 must be ok=true — it proves the test can pass at all. */
function testFoundBlockVerdict() {
  // Case 1 must be a key that exists in NO lot file, so it reaches the
  // !found block. Do not use a real lot: 15 ter-C5 was used until
  // 2026-08-17, and once Lot-15 ter resolved, that case started failing
  // on the 149 g calibre gate instead — a pass for the wrong reason.
  // Case 2 must PASS, or a test that always refuses looks like a success.
  var cases = [
    { lot: "99-1-A", qty: 10, expect: "REFUS attendu (lot inexistant)" },
    { lot: "19-8-L", qty: 10, expect: "OK attendu" }
  ];
  cases.forEach(function (c) {
    var v = cmdValidateOrderLines([{ lot: c.lot, qty: c.qty }], "GR");
    Logger.log(c.lot + "  [" + c.expect + "]  ok=" + v.ok +
               "  blocks=" + JSON.stringify(v.blocks) +
               "  warnings=" + JSON.stringify(v.warnings));
  });
}

/***************************************************************
 * RESERVATIONS - third tab on the Commandes screen.
 *
 * The Réservations tab lives in the SAME spreadsheet as the orders
 * (CMD_CFG.SS_ID). Columns: A=Lot, B=Type, C=Quantité, D=Note.
 * Headers in row 1, data from row 2.
 *
 * Two types only, matching tt_loadReservations in engine_core.js:
 *   TOUT    - the whole lot is held; no order may touch it.
 *   NOMBRE  - a count is held back; availability drops by that much.
 *
 * ONE HOLD PER LOT - enforced here, on add and on edit.
 * cmdGetReservation's NOMBRE branch OVERWRITES rather than sums, so a
 * second NOMBRE row for the same lot is silently ignored at order time.
 * The read path already behaves as one-hold-per-lot; this makes it true.
 *
 * NO Z-CLEARING. TSARAENGINE tt_clearReservedFlags wipes every column-Z
 * "LOT RÉSERVÉ" at the start of each Commandes pass and re-applies it
 * only if the hold still stands. Deleting a row here is the whole job.
 *
 * NO CACHE TO CLEAR. buildReservedAllMap is deliberately uncached and
 * read fresh on every page load, so a change here shows on the next load.
 *
 * DEAD HOLDS ARE HARMLESS. A lot that has been caged or moved leaves its
 * key behind. The dropdown is built from Stock Poisson, so a dead key
 * greys out nothing, and lot numbers never repeat, so it can never be
 * inherited by a future lot. Rows are listed as they are, not flagged.
 *
 * ROW NUMBERS ARE THE HANDLE, and they shift when a row is deleted.
 * Every write therefore re-reads the row and checks the lot still
 * matches what the browser last saw. Mismatch = stale screen, refuse.
 *
 * resSheet() THROWS on a missing tab, where cmdGetReservation returns 0.
 * Deliberate: order validation must fail safe, but a management screen
 * must say the tab is gone rather than show an empty list.
 ***************************************************************/

const RES_SHEET = "Réservations";
const RES_TYPES = ["TOUT", "NOMBRE"];

function resSheet() {
  const ss = SpreadsheetApp.openById(CMD_CFG.SS_ID);
  const sh = ss.getSheetByName(RES_SHEET);
  if (!sh) throw new Error('Onglet introuvable: "' + RES_SHEET + '"');
  return sh;
}

/** Every hold, in sheet order. row is the handle for edit and delete. */
function resList() {
  const sh = resSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const vals = sh.getRange(2, 1, lastRow - 1, 4).getValues();
  const out = [];
  for (var i = 0; i < vals.length; i++) {
    const lot = String(vals[i][0] == null ? "" : vals[i][0]).trim();
    if (!lot) continue;
    out.push({
      row: i + 2,
      lot: lot,
      key: cmdCanonKey(lot),
      type: String(vals[i][1] == null ? "" : vals[i][1]).trim().toUpperCase(),
      qty: cmdToNum(vals[i][2]),
      note: String(vals[i][3] == null ? "" : vals[i][3]).trim()
    });
  }
  return out;
}

/**
 * Validate one payload and return cleaned values, or throw.
 * excludeRow is the row being edited, so it does not clash with itself.
 * Pass 0 when adding.
 */
function resClean(p, excludeRow) {
  const lot = String(p && p.lot != null ? p.lot : "").trim();
  if (!lot) throw new Error("Choisir un lot.");
  const key = cmdCanonKey(lot);
  if (!key) throw new Error("Clé de lot invalide : " + lot);

  const type = String(p && p.type != null ? p.type : "").trim().toUpperCase();
  if (RES_TYPES.indexOf(type) < 0) throw new Error("Type invalide : " + type);

  var qty = null;
  if (type === "NOMBRE") {
    qty = cmdToNum(p.qty);
    if (qty == null || qty <= 0) {
      throw new Error("Quantité requise (nombre positif) pour une réservation NOMBRE.");
    }
    qty = Math.round(qty);
  }

  const existing = resList();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].row === excludeRow) continue;
    if (existing[i].key === key) {
      throw new Error("Le lot " + existing[i].lot + " a déjà une réservation (ligne " +
                      existing[i].row + "). Modifier celle-ci plutôt que d'en ajouter une seconde.");
    }
  }

  return {
    lot: lot,
    key: key,
    type: type,
    qty: qty,
    note: String(p && p.note != null ? p.note : "").trim()
  };
}

/**
 * The row must still hold the lot the browser last saw. Guards against a
 * stale screen writing to a row that shifted after someone else deleted.
 */
function resCheckRow(sh, row, seenLot) {
  if (!(row >= 2)) throw new Error("Ligne invalide.");
  if (row > sh.getLastRow()) throw new Error("Ligne introuvable — recharger l'écran.");
  const now = cmdCanonKey(sh.getRange(row, 1).getValue());
  if (now !== cmdCanonKey(seenLot)) {
    throw new Error("La liste a changé depuis l'affichage. Recharger l'écran.");
  }
}

function resAdd(p) {
  const c = resClean(p, 0);
  const sh = resSheet();
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 4).setValues([[c.lot, c.type, c.qty == null ? "" : c.qty, c.note]]);
  return { row: row };
}

function resUpdate(p) {
  const row = Number(p && p.row);
  const sh = resSheet();
  resCheckRow(sh, row, p && p.seenLot);
  const c = resClean(p, row);
  sh.getRange(row, 1, 1, 4).setValues([[c.lot, c.type, c.qty == null ? "" : c.qty, c.note]]);
  return { row: row };
}

function resDelete(p) {
  const row = Number(p && p.row);
  const sh = resSheet();
  resCheckRow(sh, row, p && p.seenLot);
  sh.deleteRow(row);
  return { row: row };
}

/** RUN FROM EDITOR: tsaraentry -> CommandesServer.js -> testReservations
 *  Read-only. Prints every hold with its canonical key. */
function testReservations() {
  const rows = resList();
  Logger.log("Réservations : " + rows.length);
  rows.forEach(function (r) {
    Logger.log("  ligne " + r.row + "   " + r.lot + " [" + r.key + "]   " + r.type +
               (r.qty == null ? "" : "  " + r.qty) +
               (r.note ? "   — " + r.note : ""));
  });
}


/***************************************************************
 * PRE-COMMANDES - fourth tab on the Commandes screen.
 *
 * The Demandes tab lives in the SAME spreadsheet as the orders
 * (CMD_CFG.SS_ID). Row 1 = title, row 2 = headers, data from row 3.
 * Columns: A=date, B=client, C=contact, D=type (Alevins|Poisson),
 * E=nombre, F=commentaires, G=poids (g).
 * Column G was appended 2026-08-30. Legacy rows have it empty; an
 * edit of such a row must fill it before it saves.
 *
 * These are requests the farm cannot serve yet. NOTHING else reads
 * this tab - no engine, no other project. The Commander button on the
 * client prefills the Nouvelle commande form; after that order saves,
 * the client calls demDelete for the source row.
 *
 * ALL FIELDS MANDATORY except commentaires. The client checks first
 * for an instant message; this server check is the authority.
 *
 * POIDS: Alevins must use one of DEM_POIDS_AL (the UI dropdown).
 * Poisson takes any positive number of grams (free field).
 *
 * ROW NUMBERS ARE THE HANDLE, as in Réservations: every write
 * re-reads the row and compares client+contact against what the
 * browser last saw. Mismatch = stale screen, refuse.
 *
 * FUTURE (not built): nightly check of poids x nombre against stock,
 * e-mail when a pre-commande becomes servable. The numeric poids and
 * nombre columns exist for that purpose - keep them numeric.
 ***************************************************************/

const DEM_SHEET = "Demandes";
const DEM_START = 3;                       // row 1 = title, row 2 = headers
const DEM_TYPES = ["Alevins", "Poisson"];
const DEM_POIDS_AL = [0.5, 1, 1.5, 2, 3, 4, 5, 10];

function demSheet() {
  const ss = SpreadsheetApp.openById(CMD_CFG.SS_ID);
  const sh = ss.getSheetByName(DEM_SHEET);
  if (!sh) throw new Error('Onglet introuvable: "' + DEM_SHEET + '"');
  return sh;
}

/** Every pre-commande, in sheet order. row is the handle for edit,
 *  delete and Commander. Fully empty rows are skipped, part-filled
 *  legacy rows are kept and shown as they are. */
function demList() {
  const sh = demSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < DEM_START) return [];
  const vals = sh.getRange(DEM_START, 1, lastRow - DEM_START + 1, 11).getValues();
  const tz = Session.getScriptTimeZone();
  const out = [];
  for (var i = 0; i < vals.length; i++) {
    var empty = true;
    for (var j = 0; j < 7; j++) {
      if (vals[i][j] !== "" && vals[i][j] != null) { empty = false; break; }
    }
    if (empty) continue;
    const d = vals[i][0];
    out.push({
      row: DEM_START + i,
      date: (d instanceof Date) ? Utilities.formatDate(d, tz, "yyyy-MM-dd")
                                : String(d == null ? "" : d).trim(),
      client: String(vals[i][1] == null ? "" : vals[i][1]).trim(),
      contact: String(vals[i][2] == null ? "" : vals[i][2]).trim(),
      type: String(vals[i][3] == null ? "" : vals[i][3]).trim(),
      nombre: cmdToNum(vals[i][4]),
      commentaires: String(vals[i][5] == null ? "" : vals[i][5]).trim(),
      poids: cmdToNum(vals[i][6]),
      rang: cmdToNum(vals[i][7]),
      categorie: String(vals[i][8] == null ? "" : vals[i][8]).trim(),
      livraison: String(vals[i][9] == null ? "" : vals[i][9]).trim(),
      prix: cmdToNum(vals[i][10])
    });
  }
  // Manual priority (column H) decides the queue, and the queue decides
  // who gets the fish. A blank rang keeps sheet order and sorts last,
  // so a row added outside the app never jumps the queue.
  out.sort(function (a, b) {
    const ra = (a.rang == null ? Infinity : a.rang);
    const rb = (b.rang == null ? Infinity : b.rang);
    if (ra !== rb) return ra - rb;
    return a.row - b.row;
  });
  return out;
}

/**
 * Write 1..N into column H for `list` (queue order), in ONE block
 * write over DEM_START..lastRow. Rows not in the list - empty rows -
 * get "" so the block cannot invent a rank for them.
 *
 * The ROW NUMBER IS UNCHANGED: only column H is written, so every
 * handle the browser holds stays valid and demCheckRow still guards
 * edits exactly as before.
 */
function demWriteRanks(list) {
  const sh = demSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < DEM_START) return;
  const rankOf = {};
  list.forEach(function (d, i) { rankOf[d.row] = i + 1; });
  const col = [];
  for (var r = DEM_START; r <= lastRow; r++) col.push([rankOf[r] == null ? "" : rankOf[r]]);
  sh.getRange(DEM_START, 8, col.length, 1).setValues(col);
}

/**
 * Store a complete new order. `rows` is every row number, in the
 * order wanted. It must be exactly the set of rows demList returns -
 * a missing or extra row means the list changed under the browser,
 * and the order is refused rather than applied to the wrong lines.
 */
function demSetOrder(rows) {
  const want = (rows || []).map(Number);
  const list = demList();
  if (want.length !== list.length) {
    throw new Error("La liste a changé depuis l'affichage. Recharger l'écran.");
  }
  const byRow = {};
  list.forEach(function (d) { byRow[d.row] = d; });
  const ordered = [];
  const seen = {};
  for (var i = 0; i < want.length; i++) {
    const d = byRow[want[i]];
    if (!d || seen[want[i]]) {
      throw new Error("La liste a changé depuis l'affichage. Recharger l'écran.");
    }
    seen[want[i]] = true;
    ordered.push(d);
  }
  demWriteRanks(ordered);
  return { n: ordered.length };
}

/** RUN FROM EDITOR: tsaraentry -> CommandesServer.js -> demRenumber
 *  Writes 1..N into column H in the current order. Use once after
 *  adding the column, or to repair a hand-edited rang. */
function demRenumber() {
  const list = demList();
  demWriteRanks(list);
  Logger.log("Rangs réécrits : " + list.length);
  list.forEach(function (d, i) { Logger.log("  " + (i + 1) + "  " + d.client); });
}

/** Validate one payload and return cleaned values, or throw. */
function demClean(p) {
  const dateStr = String(p && p.date != null ? p.date : "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) throw new Error("Date requise.");

  const client = String(p && p.client != null ? p.client : "").trim();
  if (!client) throw new Error("Client requis.");

  const contact = String(p && p.contact != null ? p.contact : "").trim();
  if (!contact) throw new Error("Contact requis.");

  const type = String(p && p.type != null ? p.type : "").trim();
  if (DEM_TYPES.indexOf(type) < 0) throw new Error("Type invalide : " + type);

  var nombre = cmdToNum(p.nombre);
  if (nombre == null || nombre <= 0) throw new Error("Nombre requis (entier positif).");
  nombre = Math.round(nombre);

  const poids = cmdToNum(p.poids);
  if (poids == null || poids <= 0) throw new Error("Poids requis (grammes).");
  if (type === "Alevins" && DEM_POIDS_AL.indexOf(poids) < 0) {
    throw new Error("Poids alevins invalide : " + poids + " g. Choisir dans la liste.");
  }

  // The commercial choices behind the price. Alevins price the grid by
  // provende and delivery; Poisson is flat Détail/Gros with no delivery
  // choice - livraison is forced blank so a stale value cannot survive
  // a type change.
  const categorie = String(p && p.categorie != null ? p.categorie : "").trim();
  const catOk = (type === "Alevins") ? ["avec", "sans"] : ["detail", "gros"];
  if (catOk.indexOf(categorie) < 0) {
    throw new Error((type === "Alevins" ? "Type de client" : "Qualité") +
                    " invalide : " + categorie);
  }
  var livraison = "";
  if (type === "Alevins") {
    livraison = String(p && p.livraison != null ? p.livraison : "").trim();
    if (["enlevement", "environs", "ambohim"].indexOf(livraison) < 0) {
      throw new Error("Livraison invalide : " + livraison);
    }
  }
  const prix = cmdToNum(p.prix);
  if (prix == null || prix <= 0) throw new Error("Prix requis (Ar).");

  const m = dateStr.split("-");
  return {
    date: new Date(Number(m[0]), Number(m[1]) - 1, Number(m[2])),
    client: client,
    contact: contact,
    type: type,
    nombre: nombre,
    commentaires: String(p && p.commentaires != null ? p.commentaires : "").trim(),
    poids: poids,
    categorie: categorie,
    livraison: livraison,
    prix: prix
  };
}

/** The row must still hold the client+contact the browser last saw. */
function demCheckRow(sh, row, seenClient, seenContact) {
  if (!(row >= DEM_START)) throw new Error("Ligne invalide.");
  if (row > sh.getLastRow()) throw new Error("Ligne introuvable — recharger l'écran.");
  const v = sh.getRange(row, 2, 1, 2).getValues()[0];
  const client = String(v[0] == null ? "" : v[0]).trim();
  const contact = String(v[1] == null ? "" : v[1]).trim();
  if (client !== String(seenClient == null ? "" : seenClient).trim() ||
      contact !== String(seenContact == null ? "" : seenContact).trim()) {
    throw new Error("La liste a changé depuis l'affichage. Recharger l'écran.");
  }
}

function demAdd(p) {
  const c = demClean(p);
  const sh = demSheet();
  const row = Math.max(sh.getLastRow() + 1, DEM_START);
  // Last in the queue: a new demande must not take fish from one that
  // has been waiting. Kim re-ranks with the arrows when it is urgent.
  var rang = 0;
  demList().forEach(function (d) { if (d.rang != null && d.rang > rang) rang = d.rang; });
  sh.getRange(row, 1, 1, 11).setValues(
    [[c.date, c.client, c.contact, c.type, c.nombre, c.commentaires, c.poids,
      rang + 1, c.categorie, c.livraison, c.prix]]);
  return { row: row };
}

function demUpdate(p) {
  const row = Number(p && p.row);
  const sh = demSheet();
  demCheckRow(sh, row, p && p.seenClient, p && p.seenContact);
  const c = demClean(p);
  // Two writes around column H: rang belongs to the queue, not to the
  // record, and an edit must not move the line in the queue.
  sh.getRange(row, 1, 1, 7).setValues(
    [[c.date, c.client, c.contact, c.type, c.nombre, c.commentaires, c.poids]]);
  sh.getRange(row, 9, 1, 3).setValues([[c.categorie, c.livraison, c.prix]]);
  return { row: row };
}

function demDelete(p) {
  const row = Number(p && p.row);
  const sh = demSheet();
  demCheckRow(sh, row, p && p.seenClient, p && p.seenContact);
  sh.deleteRow(row);
  return { row: row };
}

/** RUN FROM EDITOR: tsaraentry -> CommandesServer.js -> testDemandes
 *  Read-only. Prints every pre-commande with its row handle. */
function testDemandes() {
  const rows = demList();
  Logger.log("Pré-commandes : " + rows.length);
  rows.forEach(function (r) {
    Logger.log("  ligne " + r.row + "   " + r.date + "   " + r.client +
               "   " + r.contact + "   " + r.type +
               "   n=" + r.nombre + "   p=" + r.poids + " g" +
               (r.commentaires ? "   — " + r.commentaires : ""));
  });
}


/***************************************************************
 * PRE-COMMANDES STOCK CHECK - three statuses, on demand and nightly.
 *
 * Two facts per pré-commande, asked in this order:
 *   1. enough fish of its TYPE across all lots?      no -> INDISPONIBLE
 *   2. enough fish in lots whose PM lies within
 *      ±DEM_PM_TOL of the requested poids?           no -> PM
 *                                                     yes -> DISPONIBLE
 * (Kim, 2026-09-02, replacing the amount-only rule of 2026-08-30.)
 *
 * available per lot = Stock Poisson count
 *                     - réservation (TOUT excludes the lot)
 *                     - commandes entered but not yet deducted
 * Same arithmetic as cmdGetLotAvailability, from Stock Poisson in ONE
 * read instead of opening every lot file. ADVISORY: Stock Poisson can
 * lag the lot files by up to 24 h (morts and movements of the day; sales
 * are already covered by the pending term). Commander re-reads the lot
 * file for each suggested lot, and the save check is the authority.
 *
 * Type eligibility is the dropdown's own verdict (cmdLotTypeVerdict);
 * "warn" lots count - the sale is possible. A lot with no PM counts for
 * the type test, never for the band test: it cannot prove a match.
 *
 * ALLOCATION, in sheet order (oldest at the top): only a DISPONIBLE
 * pré-commande takes fish, from the band lots closest-PM-first, and it
 * records which lots and how many. PM and INDISPONIBLE take nothing -
 * they cannot be served, so holding fish of the wrong size for them
 * would only make the pré-commandes below look worse than they are.
 * Two DISPONIBLE verdicts are therefore never built on the same fish.
 *
 * The nightly mail (demNotifyServable) calls this same function.
 ***************************************************************/

const DEM_PM_TOL = 0.25;   // ±25 % band around the requested weight

/** {canonKey: "TOUT"|number} - every reservation in one read. */
function demResMap() {
  const ss = SpreadsheetApp.openById(CMD_CFG.SS_ID);
  const sh = ss.getSheetByName(RES_SHEET);
  const out = {};
  if (!sh) return out;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return out;
  const vals = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    const key = cmdCanonKey(vals[i][0]);
    if (!key) continue;
    const type = String(vals[i][1] || "").trim().toUpperCase();
    if (type === "TOUT") { out[key] = "TOUT"; continue; }
    if (type === "NOMBRE" && out[key] !== "TOUT") {
      const n = cmdToNum(vals[i][2]);
      if (n != null && n > 0) out[key] = n;
    }
  }
  return out;
}

/** {canonKey: pending fish} - ONE scan of the orders sheet.
 *  Same row eligibility as cmdGetPendingQty: Y, Z, AA all empty. */
function demPendingMap() {
  const sh = cmdSheet();
  const lastRow = findNextCommandeRow(sh) - 1;
  const out = {};
  if (lastRow < CMD_CFG.START_ROW) return out;
  const data = sh.getRange(CMD_CFG.START_ROW, 1,
                           lastRow - CMD_CFG.START_ROW + 1, 27).getValues();
  for (var i = 0; i < data.length; i++) {
    const r = data[i];
    const key = cmdCanonKey(r[CMD_CFG.COL.LOT - 1]);
    if (!key) continue;
    if (String(r[CMD_CFG.COL.LOG - 1]    || "").trim() !== "") continue;
    if (String(r[CMD_CFG.COL.ERROR - 1]  || "").trim() !== "") continue;
    if (String(r[CMD_CFG.COL.ANNULE - 1] || "").trim() !== "") continue;
    const ded = cmdDeduction(r[CMD_CFG.COL.ALEVINS_LIVRER - 1],
                             r[CMD_CFG.COL.POISSON_NB - 1]);
    if (ded == null) continue;
    out[key] = (out[key] || 0) + ded;
  }
  return out;
}

function demCheckStock() {
  // Stock Poisson lot block: N=id, O=nombre, P=PM - the block
  // updateStockPoisson rewrites each night (same read as buildLotPmMap).
  const ss = SpreadsheetApp.openById(STOCK_PM_CFG.SS_ID);
  const sh = ss.getSheetByName(STOCK_PM_CFG.SHEET);
  if (!sh) throw new Error('Onglet introuvable: "' + STOCK_PM_CFG.SHEET + '"');
  const lastRow = sh.getLastRow();
  const lots = [];                                   // {key, avail, pm, al, gr}
  const pool = { Alevins: 0, Poisson: 0 };
  if (lastRow >= STOCK_PM_CFG.START_ROW) {
    const vals = sh.getRange(STOCK_PM_CFG.START_ROW, 14,
                             lastRow - STOCK_PM_CFG.START_ROW + 1, 3).getValues();
    const resv = demResMap();
    const pend = demPendingMap();
    const notSellable = getNotSellableMap();
    const fryMax = cmdFryMaxPm();
    for (var i = 0; i < vals.length; i++) {
      const key = cmdCanonKey(vals[i][0]);
      if (!key || notSellable[key]) continue;
      const r = resv[key];
      if (r === "TOUT") continue;
      const count = cmdToNum(vals[i][1]);
      if (count == null || count <= 0) continue;
      const avail = count - (r || 0) - (pend[key] || 0);
      if (avail <= 0) continue;
      const pm = cmdToNum(vals[i][2]);
      const lot = {
        key: key, avail: avail, pm: pm,
        al: cmdLotTypeVerdict(true,  pm, fryMax).level !== "block",
        gr: cmdLotTypeVerdict(false, pm, fryMax).level !== "block"
      };
      lots.push(lot);
      if (lot.al) pool.Alevins += avail;
      if (lot.gr) pool.Poisson += avail;
    }
  }

  const out = { pool: pool, tol: DEM_PM_TOL, rows: [] };
  demList().forEach(function (d) {
    const v = { row: d.row, statut: null, manque: null, bande: null,
                lots: [], proche: [] };
    const isAL = (d.type === "Alevins"), isGR = (d.type === "Poisson");
    if (d.nombre == null || d.nombre <= 0 || d.poids == null || d.poids <= 0 ||
        (!isAL && !isGR)) {
      out.rows.push(v);                              // incomplete row: no verdict
      return;
    }

    // 1. the type test, all lots together
    const elig = lots.filter(function (l) { return l.avail > 0 && (isAL ? l.al : l.gr); });
    var total = 0;
    elig.forEach(function (l) { total += l.avail; });
    if (total < d.nombre) {
      v.statut = "INDISPONIBLE";
      v.manque = Math.ceil(d.nombre - total);
      out.rows.push(v);
      return;
    }

    // 2. the band test, closest PM first
    const lo = d.poids * (1 - DEM_PM_TOL), hi = d.poids * (1 + DEM_PM_TOL);
    const band = elig.filter(function (l) {
      return l.pm != null && l.pm >= lo && l.pm <= hi;
    }).sort(function (a, b) {
      return Math.abs(a.pm - d.poids) - Math.abs(b.pm - d.poids);
    });
    var bandTotal = 0;
    band.forEach(function (l) { bandTotal += l.avail; });
    v.bande = Math.round(bandTotal);
    if (bandTotal < d.nombre) {
      v.statut = "PM";
      // What to use instead: the lots nearest in weight that together
      // cover the order. Band lots come first (distance ~0), then the
      // closest ones outside it. A lot with no PM cannot be ranked by
      // distance and is left out - it could not prove a match either.
      // NOTHING IS SUBTRACTED: a PM line takes no fish, so these lots
      // stay available to the pré-commandes below it.
      const near = elig.filter(function (l) {
        return l.pm != null && l.avail > 0;
      }).sort(function (a, b) {
        return Math.abs(a.pm - d.poids) - Math.abs(b.pm - d.poids);
      });
      var want = d.nombre;
      for (var m = 0; m < near.length && want > 0; m++) {
        const part = Math.min(near[m].avail, want);
        want -= part;
        v.proche.push({ lot: near[m].key, nb: Math.round(part), pm: near[m].pm });
      }
      out.rows.push(v);
      return;
    }

    // 3. DISPONIBLE: take the fish, remember where from
    var need = d.nombre;
    for (var k = 0; k < band.length && need > 0; k++) {
      const take = Math.min(band[k].avail, need);
      band[k].avail -= take;
      need -= take;
      v.lots.push({ lot: band[k].key, nb: Math.round(take), pm: band[k].pm });
    }
    v.statut = "DISPONIBLE";
    out.rows.push(v);
  });
  return out;
}

/** RUN FROM EDITOR: tsaraentry -> CommandesServer.js -> testDemCheckStock
 *  Read-only. Prints the pools and the verdict per pré-commande. */
function testDemCheckStock() {
  const r = demCheckStock();
  const byRow = {};
  demList().forEach(function (d) { byRow[d.row] = d; });
  Logger.log("Pool Alevins=" + r.pool.Alevins + "  Poisson=" + r.pool.Poisson +
             "   bande ±" + Math.round(r.tol * 100) + " %");
  r.rows.forEach(function (v) {
    const d = byRow[v.row] || {};
    Logger.log("  ligne " + v.row + "  " + d.client + "  " + d.nombre + " " + d.type +
               " " + d.poids + " g   " + demVerdictText(d, v));
  });
}

/** One-line verdict, shared by the editor tests and the mail. */
function demVerdictText(d, v) {
  if (!v.statut) return "— (ligne incomplète)";
  if (v.statut === "DISPONIBLE") {
    return "DISPONIBLE  lots : " + v.lots.map(function (l) {
      return l.lot + " (" + l.nb + ")";
    }).join(", ");
  }
  if (v.statut === "PM") {
    if (v.proche && v.proche.length) {
      return "PM PAS DISPONIBLE  le plus proche : " + v.proche.map(function (l) {
        return l.lot + " (" + l.nb + " à " + l.pm + " g)";
      }).join(", ");
    }
    return "PM PAS DISPONIBLE  " + v.bande + " à " + d.poids + " g ±" +
           Math.round(DEM_PM_TOL * 100) + " %";
  }
  return "INDISPONIBLE  manque " + v.manque;
}

/***************************************************************
 * PRE-COMMANDE STATUS MAIL (2026-08-30, three statuses 2026-09-02)
 *
 * Once a night: run demCheckStock. If ANY pré-commande is new or has
 * a different status than at the last run, mail the whole list in
 * three sections - Disponible, PM pas disponible, Indisponible - with
 * the changed lines marked. No change overnight, no mail.
 *
 * HOUR 7, one hour after refreshNotSellableMap at 6. demCheckStock
 * reads that map, so running before it would judge against yesterday's
 * "pas encore vendable" list. Both run after the engine (~04:53).
 *
 * IDENTITY IS NOT THE ROW NUMBER. Rows shift the moment anyone deletes
 * a pré-commande, and a shifted row would look like a new one. The
 * signature is client|contact|type|nombre|poids, which is what the
 * farm actually means by "the same demande".
 *
 * THE STORED MAP IS REPLACED, not merged: {signature: statut} at the
 * end of the last run. (Before 2026-09-02 it held {signature: true};
 * such entries read as "new" once, which yields one full mail and then
 * the normal regime. No migration step needed.)
 *
 * ADVISORY, like the screen: Stock Poisson can lag the lot files by
 * up to 24 h, and the band is ±DEM_PM_TOL, a tolerance, not a promise.
 * The mail says "à vérifier", never "confirmé".
 *
 * FAILURE IS SILENT BY DESIGN in one direction only: if the mail
 * throws, the stored map is NOT updated, so the next run reports the
 * same changes instead of losing them.
 ***************************************************************/

const DEM_MAIL_TO = "joachim@tilapia4food.com";
const DEM_MAIL_PROP_KEY = "dem_servable_v1";
const DEM_MAIL_TRIGGER_HOUR = 7;   // after refreshNotSellableMap at 6

/** Stable identity for one pré-commande. Row numbers shift; this does not. */
function demSignature(d) {
  return [d.client, d.contact, d.type, d.nombre, d.poids].join("|");
}

/** The set of signatures servable at the end of the last run. */
function demGetNotified() {
  const raw = PropertiesService.getScriptProperties().getProperty(DEM_MAIL_PROP_KEY);
  return raw ? JSON.parse(raw) : {};
}

/**
 * RUN BY TRIGGER, and safe to run from the editor.
 * tsaraentry -> CommandesServer.js -> demNotifyServable
 * Mails the new flips, then stores the current servable set.
 */
function demStatutLabel(st) {
  return st === "DISPONIBLE" ? "Disponible"
       : st === "PM"         ? "PM pas disponible"
       : st === "INDISPONIBLE" ? "Indisponible" : "—";
}

/**
 * Everything one run needs, computed once and shared by the real send
 * and the editor dry run: the current {signature: statut} map, how
 * many lines changed, the subject and the body.
 */
function demBuildReport() {
  const check = demCheckStock();
  const byRow = {};
  demList().forEach(function (d) { byRow[d.row] = d; });
  const known = demGetNotified();

  const current = {};
  const sections = { DISPONIBLE: [], PM: [], INDISPONIBLE: [] };
  var changed = 0;

  check.rows.forEach(function (v) {
    const d = byRow[v.row];
    if (!d || !v.statut) return;
    const sig = demSignature(d);
    current[sig] = v.statut;

    var mark = "";
    if (typeof known[sig] !== "string") {            // absent, or pre-09-02 "true"
      mark = "NOUVEAU"; changed++;
    } else if (known[sig] !== v.statut) {
      mark = "CHANGÉ (était " + demStatutLabel(known[sig]) + ")"; changed++;
    }

    var line = "  • " + d.client + "  (" + d.contact + ")  —  " +
               d.nombre + " " + d.type + " de " + d.poids + " g" +
               (mark ? "   <-- " + mark : "") + "\n" +
               "      " + demVerdictText(d, v) + "\n" +
               (d.commentaires ? "      " + d.commentaires + "\n" : "") +
               "      demandé le " + d.date + "\n";
    sections[v.statut].push(line);
  });

  const n = { DISPONIBLE: sections.DISPONIBLE.length,
              PM: sections.PM.length,
              INDISPONIBLE: sections.INDISPONIBLE.length };

  var body = "Pré-commandes — état du jour   (" + changed + " changement(s))\n\n";
  [["DISPONIBLE", "DISPONIBLES"], ["PM", "PM PAS DISPONIBLE"],
   ["INDISPONIBLE", "INDISPONIBLES"]].forEach(function (p) {
    body += p[1] + " (" + n[p[0]] + ")\n" +
            (sections[p[0]].length ? sections[p[0]].join("\n") : "  (aucune)\n") + "\n";
  });
  body += "Stock disponible : Alevins " + check.pool.Alevins +
          "  —  Poisson " + check.pool.Poisson + "\n\n" +
          "À VÉRIFIER avant de confirmer au client :\n" +
          "  - le poids est jugé à ±" + Math.round(DEM_PM_TOL * 100) +
          " % du poids demandé, sur le PM du lot\n" +
          "  - Stock Poisson peut avoir jusqu'à 24 h de retard\n" +
          "  - le contrôle définitif se fait à l'enregistrement de la commande\n\n" +
          "Onglet Pré-commandes de l'écran Commandes : le bouton Commander " +
          "pré-remplit les lots indiqués.";

  return {
    current: current, changed: changed, counts: n, body: body,
    subject: "Tsara — pré-commandes : " + n.DISPONIBLE + " disponible(s), " +
             n.PM + " PM, " + n.INDISPONIBLE + " indisponible(s)" +
             (changed ? "  (" + changed + " changement(s))" : "")
  };
}

/**
 * RUN BY TRIGGER, and safe to run from the editor.
 * tsaraentry -> CommandesServer.js -> demNotifyServable
 * Mails when something changed, then stores the current map.
 */
function demNotifyServable() {
  const r = demBuildReport();
  Logger.log("Disponibles " + r.counts.DISPONIBLE + "   PM " + r.counts.PM +
             "   Indisponibles " + r.counts.INDISPONIBLE +
             "   changements " + r.changed);

  if (r.changed) {
    // Mail first, store second. A throw here leaves the stored map
    // untouched, so the next run reports the same changes.
    MailApp.sendEmail(DEM_MAIL_TO, r.subject, r.body);
    Logger.log("Mail envoyé à " + DEM_MAIL_TO);
  } else {
    Logger.log("Aucun changement — pas de mail.");
  }

  PropertiesService.getScriptProperties()
    .setProperty(DEM_MAIL_PROP_KEY, JSON.stringify(r.current));
  return r.counts;
}

/**
 * RUN FROM EDITOR ONCE. Installs the nightly mail trigger.
 * Deletes any existing trigger for the same handler first, so running
 * it twice cannot leave two triggers mailing the same flips.
 */
function installDemNotifyTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "demNotifyServable") {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  ScriptApp.newTrigger("demNotifyServable")
    .timeBased()
    .everyDays(1)
    .atHour(DEM_MAIL_TRIGGER_HOUR)
    .create();
  Logger.log("Anciens déclencheurs supprimés : " + removed);
  Logger.log("Déclencheur quotidien installé vers " + DEM_MAIL_TRIGGER_HOUR + " h.");
}

/** RUN FROM EDITOR: what the next mail would contain. Sends NOTHING,
 *  stores NOTHING. Use this before installing the trigger. */
function testDemNotify() {
  const r = demBuildReport();
  Logger.log(r.changed ? "UN MAIL PARTIRAIT :" : "PAS DE MAIL (aucun changement). Contenu si envoyé :");
  Logger.log("Sujet : " + r.subject);
  Logger.log(r.body);
}

/** RUN FROM EDITOR: forget every notification, so the next run mails
 *  every servable pré-commande again. Use after testing. */
function resetDemNotified() {
  PropertiesService.getScriptProperties().deleteProperty(DEM_MAIL_PROP_KEY);
  Logger.log("Historique des notifications effacé.");
}


/***************************************************************
 * HISTORIQUE - fifth tab on the Commandes screen. READ ONLY.
 *
 * Every order of the CMD_CFG.SHEET year, newest first. Rows are
 * grouped by order number exactly as cmdFindOrders groups them: a
 * commercial order spans one row per lot, and the FIRST row carries
 * the number while the others say "commande groupée". Rows with no
 * number at all stay separate, keyed by row.
 *
 * NOT cmdFindOrders. That one is a work list: it drops orders that
 * are delivered AND paid, drops cancelled ones, and stops at 25 hits.
 * A history that hides finished business is not a history.
 *
 * STATUT - "ouvert / livré / payé" is three words for two independent
 * facts, delivery (col V) and payment (col U), which give four states.
 * All four are reported rather than collapsed, plus Annulé:
 *   Annulé          col AA filled
 *   Livré et payé   both dates present
 *   Livré           delivered, not yet paid
 *   Payé            paid, not yet delivered (prepayment)
 *   Ouvert          neither
 *
 * QUANTITY is the number ORDERED: col F for alevins, col N for
 * grossis (N = (L*1000)/M, the fish count behind the kg). Col H, the
 * +5% the engine actually deducts, is a delivery figure and is not
 * what this screen was asked for.
 *
 * DISPLAY VALUES, not raw values: dates and money arrive already
 * formatted by the sheet, so the table shows what the sheet shows.
 ***************************************************************/

function histList() {
  const sh = cmdSheet();
  const C = CMD_CFG.COL;
  const lastRow = findNextCommandeRow(sh) - 1;
  if (lastRow < CMD_CFG.START_ROW) return { orders: [], annee: CMD_CFG.SHEET };

  const n = lastRow - CMD_CFG.START_ROW + 1;
  const vals = sh.getRange(CMD_CFG.START_ROW, 1, n, C.ANNULE).getDisplayValues();

  const groups = {};
  const order = [];

  for (var i = 0; i < n; i++) {
    const r = vals[i];
    const rowNum = CMD_CFG.START_ROW + i;
    const orderNo = String(r[C.ORDER_NO - 1] || "").trim();
    const key = orderNo || ("__row" + rowNum);

    if (!groups[key]) {
      groups[key] = {
        orderNumber: orderNo,
        key: key,
        client: r[C.CLIENT - 1],
        dateCommande: r[C.DATE_CMD - 1],
        paiement: r[C.PAIEMENT - 1],
        dateLivraison: r[C.DATE_LIVRAISON - 1],
        facture: r[C.FACTURE - 1],
        annule: String(r[C.ANNULE - 1] || "").trim(),
        alevinsTotal: 0,
        poissonNbTotal: 0,
        montantAr: 0
      };
      order.push(key);
    }

    const g = groups[key];
    g.alevinsTotal   += cmdNumFromDisplay_(r[C.ALEVINS_NB - 1]);
    g.poissonNbTotal += cmdNumFromDisplay_(r[C.POISSON_NB - 1]);
    g.montantAr      += cmdNumFromDisplay_(r[C.ARGENT_ALEVINS - 1]) +
                        cmdNumFromDisplay_(r[C.ARGENT_POISSON - 1]);
    // Any row of the order can carry the state or the identity.
    if (!g.facture && r[C.FACTURE - 1]) g.facture = r[C.FACTURE - 1];
    if (!g.client && r[C.CLIENT - 1]) g.client = r[C.CLIENT - 1];
    if (!g.dateCommande && r[C.DATE_CMD - 1]) g.dateCommande = r[C.DATE_CMD - 1];
    if (!g.paiement && r[C.PAIEMENT - 1]) g.paiement = r[C.PAIEMENT - 1];
    if (!g.dateLivraison && r[C.DATE_LIVRAISON - 1]) g.dateLivraison = r[C.DATE_LIVRAISON - 1];
    if (!g.annule && String(r[C.ANNULE - 1] || "").trim()) {
      g.annule = String(r[C.ANNULE - 1]).trim();
    }
  }

  const out = [];
  for (var j = order.length - 1; j >= 0; j--) {     // newest first
    const g = groups[order[j]];
    const delivered = String(g.dateLivraison || "").trim() !== "";
    const paid      = String(g.paiement || "").trim() !== "";
    g.statut = g.annule ? "Annulé"
             : (delivered && paid) ? "Livré et payé"
             : delivered ? "Livré"
             : paid ? "Payé"
             : "Ouvert";
    out.push(g);
  }
  return { orders: out, annee: CMD_CFG.SHEET };
}

/** RUN FROM EDITOR: tsaraentry -> CommandesServer.js -> testHistorique
 *  Read-only. Prints the count per status and the money total. */
function testHistorique() {
  const h = histList();
  const byStat = {};
  var total = 0;
  h.orders.forEach(function (o) {
    byStat[o.statut] = (byStat[o.statut] || 0) + 1;
    total += o.montantAr;
  });
  Logger.log("Année " + h.annee + " : " + h.orders.length + " commandes");
  Object.keys(byStat).sort().forEach(function (k) {
    Logger.log("  " + k + " : " + byStat[k]);
  });
  Logger.log("Montant total : " + Math.round(total) + " Ar");
}


/***************************************************************
 * LEGACY FACTURES — READ ONLY, WRITES NOTHING. (2026-08-30)
 *
 * Until 2026-08-30 the invoice number was minted when a Date livraison
 * was entered, so every order delivered and NOT paid was given a
 * number before any payment existed. This lists those orders, so Kim
 * can decide whether to clear them by hand.
 *
 * Nothing in the code reads this list. Run it, read the log, delete
 * this function when the clean-up is done.
 *
 * RUN FROM EDITOR: tsaraentry -> CommandesServer.js -> testLegacyFactures
 ***************************************************************/
function testLegacyFactures() {
  const sh = cmdSheet();
  const C = CMD_CFG.COL;
  const lastRow = findNextCommandeRow(sh) - 1;
  if (lastRow < CMD_CFG.START_ROW) { Logger.log("Aucune commande."); return; }

  const n = lastRow - CMD_CFG.START_ROW + 1;
  const vals = sh.getRange(CMD_CFG.START_ROW, 1, n, C.ANNULE).getDisplayValues();

  const seen = {};
  var count = 0;
  Logger.log("Commandes LIVREES, NON PAYEES, avec un n° facture :");
  for (var i = 0; i < n; i++) {
    const r = vals[i];
    if (String(r[C.ANNULE - 1] || "").trim()) continue;
    if (!String(r[C.DATE_LIVRAISON - 1] || "").trim()) continue;
    if (String(r[C.PAIEMENT - 1] || "").trim()) continue;
    const fact = String(r[C.FACTURE - 1] || "").trim();
    if (!fact) continue;

    const key = String(r[C.ORDER_NO - 1] || "").trim() || ("__row" + (CMD_CFG.START_ROW + i));
    if (seen[key]) continue;
    seen[key] = true;
    count++;
    Logger.log("  ligne " + (CMD_CFG.START_ROW + i) + "   " + key +
               "   " + r[C.CLIENT - 1] + "   facture " + fact +
               "   livrée le " + r[C.DATE_LIVRAISON - 1]);
  }
  Logger.log("Total : " + count + " commande(s).");
}


/***************************************************************
 * COMMANDES RÉCURRENTES - fifth tab on the Commandes screen.
 *
 * A standing weekly order: "client X takes 200 kg of grossis every
 * Monday until the end date". Grossis only - alevins are never
 * recurrent (Kim, 2026-09-03).
 *
 * WHAT THIS DOES NOT DO: it never writes an order row. Every Monday
 * at 05h it writes ONE PRÉ-COMMANDE (Demandes tab) per due rule, and
 * stops there. The worker then uses the existing Commander button,
 * which prefills the order form from the lot allocation, and the
 * existing save path applies the full stock guard. The "2026" tab
 * keeps exactly two writers: cmdCreateOrder and cmdRecordFulfilment.
 *
 * WHY NO STORED WEIGHT PER FISH. A pré-commande stores a fish COUNT
 * and a weight in grams; the rule stores KG. The count is derived at
 * generation time from the live PM of the sellable grossi lots, so
 * nothing to maintain per client and nothing to go stale. That count
 * only feeds the availability verdict: the order form takes kg
 * (column L) and the sheet computes N = (L*1000)/M, and demCommander
 * rebuilds the kg line from the allocated lot's real PM. So a rough
 * PM here can never distort the order that is finally saved.
 *
 * NO PRICE IS WRITTEN. Column K of the Demandes row is left blank, the
 * legacy shape demCommander already handles: the order form's own
 * Tarifs lookup fills the price at order time, so it is never stale.
 * Consequence: a worker who EDITS a generated pré-commande in the UI
 * will be asked for a price, because demClean requires one. Editing
 * is not the normal path - Commander is.
 *
 * PRIORITY. A generated row is placed FIRST in the queue (rang 1),
 * ahead of one-off pré-commandes (Kim, 2026-09-03). demCheckStock
 * allocates in queue order, so a recurring client takes its fish
 * first, and a waiting one-off can flip to PM or INDISPONIBLE on a
 * Monday morning. That is the intended meaning of priority.
 *
 * IDEMPOTENT BY MARKER. Each generated row carries
 * "RÉCURRENTE <kg> kg — lundi <yyyy-MM-dd>" in commentaires. A rule
 * whose marker for this Monday is already present is skipped, so a
 * re-run of the trigger cannot double-order.
 *
 * ONE ACTIVE RULE PER CLIENT - enforced in recClean, as Réservations
 * enforces one hold per lot. It is what makes the marker unique.
 *
 * ROW NUMBERS ARE THE HANDLE, as in Réservations and Demandes: every
 * write re-reads the row and checks the client still matches what the
 * browser last saw. Mismatch = stale screen, refuse.
 ***************************************************************/

const REC_SHEET = "Récurrentes";
const REC_START = 2;                       // row 1 = headers
const REC_QUALITES = ["detail", "gros"];
const REC_TRIGGER_HOUR = 5;                // Monday 05h, before the farm morning

/* Overdue mail. CONFIRM THESE ADDRESSES BEFORE INSTALLING THE TRIGGER.
 * DEM_MAIL_TO above is joachim@tilapia4food.com, a different account. */
const REC_MAIL_TO = "joachim@tilapia4food.com,charles@jdsresearch.com," +
                    "audry@jdsresearch.com,hasina@jdsresearch.com";

const REC_HEADERS = ["Client", "Contact", "Kg", "Qualité", "Début", "Fin",
                     "Actif", "Remarques", "Dernier PM", "Dernière génération"];

function recSheet() {
  const ss = SpreadsheetApp.openById(CMD_CFG.SS_ID);
  const sh = ss.getSheetByName(REC_SHEET);
  if (!sh) throw new Error('Onglet introuvable: "' + REC_SHEET +
                           '". Lancer recEnsureSheet une fois.');
  return sh;
}

/**
 * RUN FROM EDITOR ONCE: tsaraentry -> CommandesServer.js -> recEnsureSheet
 *
 * Creates the tab AT THE END of the spreadsheet. Position matters:
 * CDR/poisson.js falls back to getSheets()[0] when "2026" is missing,
 * so no new tab may ever take first position.
 */
function recEnsureSheet() {
  const ss = SpreadsheetApp.openById(CMD_CFG.SS_ID);
  var sh = ss.getSheetByName(REC_SHEET);
  if (sh) {
    Logger.log('Onglet "' + REC_SHEET + '" déjà présent (position ' +
               sh.getIndex() + ").");
    return sh.getIndex();
  }
  sh = ss.insertSheet(REC_SHEET, ss.getNumSheets());
  sh.getRange(1, 1, 1, REC_HEADERS.length).setValues([REC_HEADERS])
    .setFontWeight("bold");
  sh.setFrozenRows(1);
  Logger.log('Onglet "' + REC_SHEET + '" créé en position ' + sh.getIndex() +
             " sur " + ss.getNumSheets() + ".");
  return sh.getIndex();
}

/** Midnight local date, or null. Mirrors cmdParseDate. */
function recParseDate(isoStr) {
  if (!isoStr) return null;
  const m = String(isoStr).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/** yyyy-MM-dd in script timezone, or "" for a non-date. */
function recIso(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/** Midnight of the Monday of the week CONTAINING d. */
function recMondayOf(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const shift = (x.getDay() + 6) % 7;        // Sunday=0 -> 6, Monday=1 -> 0
  x.setDate(x.getDate() - shift);
  return x;
}

/** The commentaires marker that makes a generated row recognisable. */
function recMarker(kg, mondayIso) {
  return "RÉCURRENTE " + kg + " kg — lundi " + mondayIso;
}

/** The yyyy-MM-dd inside a marker, or "" when the text is not one. */
function recMarkerDate(text) {
  const m = String(text == null ? "" : text)
    .match(/R[EÉ]CURRENTE[^\n]*?(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : "";
}

/** Every rule, in sheet order. row is the handle for edit and delete. */
function recList() {
  const sh = recSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < REC_START) return [];
  const vals = sh.getRange(REC_START, 1, lastRow - REC_START + 1, 10).getValues();
  const out = [];
  for (var i = 0; i < vals.length; i++) {
    const client = String(vals[i][0] == null ? "" : vals[i][0]).trim();
    if (!client) continue;
    out.push({
      row: REC_START + i,
      client: client,
      contact: String(vals[i][1] == null ? "" : vals[i][1]).trim(),
      kg: cmdToNum(vals[i][2]),
      qualite: String(vals[i][3] == null ? "" : vals[i][3]).trim().toLowerCase(),
      debut: recIso(vals[i][4]),
      fin: recIso(vals[i][5]),
      actif: String(vals[i][6] == null ? "" : vals[i][6]).trim() !== "",
      remarques: String(vals[i][7] == null ? "" : vals[i][7]).trim(),
      dernierPm: cmdToNum(vals[i][8]),
      derniereGen: recIso(vals[i][9])
    });
  }
  return out;
}

/**
 * Validate one payload and return cleaned values, or throw.
 * excludeRow is the row being edited, so it does not clash with itself.
 * Pass 0 when adding.
 */
function recClean(p, excludeRow) {
  const client = String(p && p.client != null ? p.client : "").trim();
  if (!client) throw new Error("Client requis.");

  const contact = String(p && p.contact != null ? p.contact : "").trim();
  if (!contact) throw new Error("Contact requis.");

  var kg = cmdToNum(p && p.kg);
  if (kg == null || kg <= 0) throw new Error("Quantité requise (kg, nombre positif).");
  kg = Math.round(kg * 10) / 10;

  const qualite = String(p && p.qualite != null ? p.qualite : "").trim().toLowerCase();
  if (REC_QUALITES.indexOf(qualite) < 0) throw new Error("Qualité invalide : " + qualite);

  const debut = recParseDate(p && p.debut);
  if (!debut) throw new Error("Date de début requise.");
  const fin = recParseDate(p && p.fin);
  if (!fin) throw new Error("Date de fin requise.");
  if (fin.getTime() < debut.getTime()) {
    throw new Error("La date de fin précède la date de début.");
  }

  // One ACTIVE rule per client: it is what makes the weekly marker
  // unique, and two standing orders for one client is an entry error.
  const actif = !!(p && p.actif);
  if (actif) {
    const existing = recList();
    for (var i = 0; i < existing.length; i++) {
      if (existing[i].row === excludeRow) continue;
      if (!existing[i].actif) continue;
      if (existing[i].client.toLowerCase() === client.toLowerCase()) {
        throw new Error("Le client " + existing[i].client +
                        " a déjà une commande récurrente active (ligne " +
                        existing[i].row + "). Modifier celle-ci.");
      }
    }
  }

  return {
    client: client, contact: contact, kg: kg, qualite: qualite,
    debut: debut, fin: fin, actif: actif,
    remarques: String(p && p.remarques != null ? p.remarques : "").trim()
  };
}

/** The row must still hold the client the browser last saw. */
function recCheckRow(sh, row, seenClient) {
  if (!(row >= REC_START)) throw new Error("Ligne invalide.");
  if (row > sh.getLastRow()) throw new Error("Ligne introuvable — recharger l'écran.");
  const now = String(sh.getRange(row, 1).getValue() || "").trim();
  if (now !== String(seenClient == null ? "" : seenClient).trim()) {
    throw new Error("La liste a changé depuis l'affichage. Recharger l'écran.");
  }
}

function recAdd(p) {
  const c = recClean(p, 0);
  const sh = recSheet();
  const row = Math.max(sh.getLastRow() + 1, REC_START);
  // Columns I and J (dernier PM, dernière génération) belong to the
  // generator alone and stay empty here.
  sh.getRange(row, 1, 1, 8).setValues(
    [[c.client, c.contact, c.kg, c.qualite, c.debut, c.fin,
      c.actif ? "x" : "", c.remarques]]);
  return { row: row };
}

function recUpdate(p) {
  const row = Number(p && p.row);
  const sh = recSheet();
  recCheckRow(sh, row, p && p.seenClient);
  const c = recClean(p, row);
  sh.getRange(row, 1, 1, 8).setValues(
    [[c.client, c.contact, c.kg, c.qualite, c.debut, c.fin,
      c.actif ? "x" : "", c.remarques]]);
  return { row: row };
}

function recDelete(p) {
  const row = Number(p && p.row);
  const sh = recSheet();
  recCheckRow(sh, row, p && p.seenClient);
  sh.deleteRow(row);
  return { row: row };
}

/**
 * Volume-weighted mean PM of the lots a grossi order could be served
 * from, in grams. Same pool arithmetic as demCheckStock - Stock
 * Poisson count minus réservations minus commandes non déduites, lots
 * excluded by getNotSellableMap or by the type gate - so the derived
 * fish count agrees with the verdict the screen will show.
 *
 * Returns null when no sellable grossi lot has a PM. The caller then
 * falls back to the rule's last used PM.
 */
function recGrossisPm() {
  const ss = SpreadsheetApp.openById(STOCK_PM_CFG.SS_ID);
  const sh = ss.getSheetByName(STOCK_PM_CFG.SHEET);
  if (!sh) throw new Error('Onglet introuvable: "' + STOCK_PM_CFG.SHEET + '"');
  const lastRow = sh.getLastRow();
  if (lastRow < STOCK_PM_CFG.START_ROW) return null;

  const vals = sh.getRange(STOCK_PM_CFG.START_ROW, 14,
                           lastRow - STOCK_PM_CFG.START_ROW + 1, 3).getValues();
  const resv = demResMap();
  const pend = demPendingMap();
  const notSellable = getNotSellableMap();
  const fryMax = cmdFryMaxPm();

  var fish = 0, mass = 0;
  for (var i = 0; i < vals.length; i++) {
    const key = cmdCanonKey(vals[i][0]);
    if (!key || notSellable[key]) continue;
    const r = resv[key];
    if (r === "TOUT") continue;
    const count = cmdToNum(vals[i][1]);
    if (count == null || count <= 0) continue;
    const avail = count - (r || 0) - (pend[key] || 0);
    if (avail <= 0) continue;
    const pm = cmdToNum(vals[i][2]);
    if (pm == null || pm <= 0) continue;
    if (cmdLotTypeVerdict(false, pm, fryMax).level === "block") continue;
    fish += avail;
    mass += avail * pm;
  }
  if (fish <= 0) return null;
  return Math.round((mass / fish) * 10) / 10;
}

/**
 * The rules due on `monday` that have not been generated for it yet.
 * Returns [{rule, marker}]. Pure read - used by the preview and by the
 * generator, so what the preview shows is what the generator writes.
 */
function recDueRules(monday) {
  const mondayIso = recIso(monday);
  const t = monday.getTime();
  const seen = {};
  demList().forEach(function (d) {
    const md = recMarkerDate(d.commentaires);
    if (md) seen[d.client.toLowerCase() + "|" + md] = true;
  });

  const out = [];
  recList().forEach(function (rule) {
    if (!rule.actif) return;
    if (rule.kg == null || rule.kg <= 0) return;
    const deb = recParseDate(rule.debut), fin = recParseDate(rule.fin);
    if (!deb || !fin) return;
    if (t < deb.getTime() || t > fin.getTime()) return;
    if (seen[rule.client.toLowerCase() + "|" + mondayIso]) return;
    out.push({ rule: rule, marker: recMarker(rule.kg, mondayIso) });
  });
  return out;
}

/**
 * RUN FROM EDITOR: tsaraentry -> CommandesServer.js -> recPreviewWeekly
 * Read-only. Shows exactly what the next Monday run would write.
 */
function recPreviewWeekly() {
  const monday = recMondayOf(new Date());
  const due = recDueRules(monday);
  const pm = recGrossisPm();
  Logger.log("Lundi de la semaine : " + recIso(monday));
  Logger.log("PM grossis pondéré  : " + (pm == null ? "AUCUN STOCK" : pm + " g"));
  Logger.log("Règles dues         : " + due.length);
  due.forEach(function (d) {
    const use = (pm != null) ? pm : d.rule.dernierPm;
    Logger.log("  " + d.rule.client + "   " + d.rule.kg + " kg   " +
               d.rule.qualite + "   PM " + (use == null ? "INCONNU" : use) +
               "   -> " + (use ? Math.round(d.rule.kg * 1000 / use) : "?") +
               " poissons   [" + d.marker + "]");
  });
  return due.length;
}

/**
 * THE MONDAY JOB. Writes one pré-commande per due rule, places them at
 * the head of the queue, stamps the rules, then mails whatever is
 * overdue. Installed by installRecTrigger.
 */
function recGenerateWeekly() {
  const monday = recMondayOf(new Date());
  const mondayIso = recIso(monday);
  const due = recDueRules(monday);
  Logger.log("Génération récurrentes — lundi " + mondayIso +
             " — " + due.length + " règle(s) due(s).");

  if (due.length) {
    const pm = recGrossisPm();
    const sh = demSheet();
    const newRows = [];

    due.forEach(function (d) {
      const use = (pm != null) ? pm : d.rule.dernierPm;
      if (use == null || use <= 0) {
        Logger.log("  " + d.rule.client + " : AUCUN PM disponible et aucun PM " +
                   "mémorisé — ligne non écrite.");
        return;
      }
      const nombre = Math.round(d.rule.kg * 1000 / use);
      const row = Math.max(sh.getLastRow() + 1, DEM_START);
      // Column H (rang) is left blank so demList sorts this row last;
      // the reorder below puts it first. Column K (prix) is left blank
      // on purpose - the order form prices from Tarifs at order time.
      sh.getRange(row, 1, 1, 11).setValues(
        [[monday, d.rule.client, d.rule.contact, "Poisson", nombre,
          d.marker + (d.rule.remarques ? " — " + d.rule.remarques : ""),
          use, "", d.rule.qualite, "", ""]]);
      newRows.push(row);
      // Stamp the rule: I = PM used, J = date generated.
      recSheet().getRange(d.rule.row, 9, 1, 2).setValues([[use, monday]]);
      Logger.log("  " + d.rule.client + " : " + d.rule.kg + " kg, PM " + use +
                 " g -> " + nombre + " poissons, ligne " + row + ".");
    });

    // PRIORITY: generated rows first, in rule order, then everything
    // else in its current queue order. demWriteRanks writes 1..N over
    // column H only, so every row handle the UI holds stays valid.
    if (newRows.length) {
      SpreadsheetApp.flush();
      const isNew = {};
      newRows.forEach(function (r) { isNew[r] = true; });
      const all = demList();
      const head = [], tail = [];
      all.forEach(function (d) { (isNew[d.row] ? head : tail).push(d); });
      head.sort(function (a, b) { return newRows.indexOf(a.row) - newRows.indexOf(b.row); });
      demWriteRanks(head.concat(tail));
    }
  }

  recMailOverdue(monday);
}

/**
 * Mail the team about generated pré-commandes from an EARLIER week that
 * are still sitting in the list - the order was never placed. Nothing
 * is sent when nothing is overdue (Kim, 2026-09-03).
 *
 * BLIND SPOT, ACCEPTED: this runs inside the Monday job, so a trigger
 * that never fires sends no mail and writes no row. Silence then looks
 * like "nothing overdue". Apps Script failure notifications are the
 * only alarm for a dead trigger.
 */
function recMailOverdue(monday) {
  const mondayIso = recIso(monday);
  const late = [];
  demList().forEach(function (d) {
    const md = recMarkerDate(d.commentaires);
    if (md && md < mondayIso) late.push({ d: d, when: md });
  });
  if (!late.length) {
    Logger.log("Aucune commande récurrente en retard.");
    return 0;
  }
  const lines = late.map(function (x) {
    return "- " + x.d.client + "   (" + x.d.commentaires + ")   rang " +
           (x.d.rang == null ? "?" : x.d.rang);
  });
  const body =
    "Ces commandes récurrentes ont été préparées mais la commande n'a " +
    "jamais été passée :\n\n" + lines.join("\n") +
    "\n\nÉcran Commandes -> onglet Pré-commandes -> bouton Commander.\n" +
    "Si le client a annulé, supprimer la ligne.\n";
  MailApp.sendEmail(REC_MAIL_TO,
                    "Tsara — " + late.length + " commande(s) récurrente(s) en retard",
                    body);
  Logger.log("Mail retard envoyé à " + REC_MAIL_TO + " (" + late.length + ").");
  return late.length;
}

/**
 * RUN FROM EDITOR ONCE: tsaraentry -> CommandesServer.js -> installRecTrigger
 * Monday 05h. Deletes any existing trigger for the same handler first,
 * so running it twice cannot leave two generators writing the same week.
 */
function installRecTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "recGenerateWeekly") {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  ScriptApp.newTrigger("recGenerateWeekly")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(REC_TRIGGER_HOUR)
    .create();
  Logger.log("Trigger recGenerateWeekly installé (lundi " + REC_TRIGGER_HOUR +
             "h). Anciens supprimés : " + removed);
}

/** RUN FROM EDITOR: tsaraentry -> CommandesServer.js -> testRecurrences
 *  Read-only. Prints every rule and its state. */
function testRecurrences() {
  const rows = recList();
  Logger.log("Commandes récurrentes : " + rows.length);
  rows.forEach(function (r) {
    Logger.log("  ligne " + r.row + "   " + r.client + "   " + r.kg + " kg   " +
               r.qualite + "   " + r.debut + " -> " + r.fin +
               (r.actif ? "   ACTIF" : "   inactif") +
               (r.dernierPm ? "   dernier PM " + r.dernierPm : "") +
               (r.derniereGen ? "   dernière génération " + r.derniereGen : ""));
  });
}
