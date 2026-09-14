/***************************************************************
 * ReductionServer.js - feed reduction box on the Nourrissage screen.
 *
 * Stock poisson, tab "lot":
 *   AF (32) = per-lot tick "Reduire"
 *   AG (33) = per-lot reduction rate in percent (30 = 30 %)
 * Column S consumes both:
 *   S = (O*P)*R*IF(AND(AF=TRUE;AG<>"");1-AG/100;1)
 * AH1 (the old single global rate) is no longer read by S.
 * This file only writes AF and AG. No feeding maths lives here.
 ***************************************************************/

const RED_CFG = {
  SHEET: "lot",
  START_ROW: 3,
  END_ROW: 50,
  LOT_COL: 14,    // N
  PM_COL: 16,     // P  = poids moyen (g)
  TICK_COL: 32,   // AF = Reduire
  RATE_COL: 33,   // AG = reduction %
  MIN_PM: 150     // strictly above
};

function redOpenLotSheet() {
  const ss = SpreadsheetApp.openById(CFG.STOCK_POISSON_SS_ID);
  const sh = ss.getSheetByName(RED_CFG.SHEET);
  if (!sh) throw new Error('Stock Poisson sheet not found: "' + RED_CFG.SHEET + '"');
  return sh;
}

/**
 * Rows for the box: every lot with PM > 150 g, PLUS every lot already
 * ticked - AF/AG apply to it whether or not it is heavy enough to show.
 * Returns { rows: [{ lot, pm, reduce, rate }] }, rate in percent.
 */
function getReductionLots() {
  const sh = redOpenLotSheet();
  const n = RED_CFG.END_ROW - RED_CFG.START_ROW + 1;
  const block = sh.getRange(RED_CFG.START_ROW, RED_CFG.LOT_COL, n,
                            RED_CFG.RATE_COL - RED_CFG.LOT_COL + 1).getValues();

  const rows = [];
  for (let i = 0; i < n; i++) {
    const lot = String(block[i][0] || "").trim();
    if (!lot) continue;
    const pm = Number(block[i][RED_CFG.PM_COL - RED_CFG.LOT_COL]);
    const tick = block[i][RED_CFG.TICK_COL - RED_CFG.LOT_COL] === true;
    const raw = block[i][RED_CFG.RATE_COL - RED_CFG.LOT_COL];
    if (!(isFinite(pm) && pm > RED_CFG.MIN_PM) && !tick) continue;
    rows.push({
      lot: lot,
      pm: isFinite(pm) ? Math.round(pm * 10) / 10 : "",
      reduce: tick,
      rate: (raw === "" || raw === null || !isFinite(Number(raw))) ? "" : Number(raw)
    });
  }
  return { rows: rows };
}

/**
 * req = { lots: [{ lot, reduce, rate }] }, rate in percent, "" allowed on an
 * unticked row (the rate is kept for later re-ticking if given).
 * Writes AF and AG for the listed lots only, then reads both columns back -
 * the confirmation describes Stock poisson, not the request. Rows are
 * re-read here: the weekly engine re-sorts them.
 */
function saveReduction(req) {
  if (!req || !req.lots || !req.lots.length) throw new Error("Aucun lot a enregistrer.");

  req.lots.forEach(function (item) {
    if (item.reduce !== true) return;
    const rate = Number(item.rate);
    if (item.rate === "" || item.rate === null || !isFinite(rate) || rate <= 0 || rate > 100) {
      throw new Error("Pourcentage invalide pour " + item.lot +
                      " (attendu 1 a 100, recu: " + item.rate + "). Rien n'a ete enregistre.");
    }
  });

  const sh = redOpenLotSheet();
  const n = RED_CFG.END_ROW - RED_CFG.START_ROW + 1;
  const lotVals = sh.getRange(RED_CFG.START_ROW, RED_CFG.LOT_COL, n, 1).getValues();
  const tickRange = sh.getRange(RED_CFG.START_ROW, RED_CFG.TICK_COL, n, 1);
  const rateRange = sh.getRange(RED_CFG.START_ROW, RED_CFG.RATE_COL, n, 1);
  const tickVals = tickRange.getValues();
  const rateVals = rateRange.getValues();

  const rowByLot = {};
  for (let i = 0; i < n; i++) {
    const lot = String(lotVals[i][0] || "").trim();
    if (lot) rowByLot[lot] = i;
  }

  const missing = [];
  let changed = 0;
  req.lots.forEach(function (item) {
    const lot = String(item.lot || "").trim();
    if (!(lot in rowByLot)) { missing.push(lot); return; }
    const i = rowByLot[lot];
    const wantTick = item.reduce === true;
    const wantRate = (item.rate === "" || item.rate === null) ? "" : Number(item.rate);
    if ((tickVals[i][0] === true) !== wantTick) changed++;
    if (String(rateVals[i][0]) !== String(wantRate)) changed++;
    tickVals[i][0] = wantTick;
    rateVals[i][0] = wantRate;
  });

  if (missing.length) {
    throw new Error("Lot(s) introuvable(s) dans Stock poisson: " + missing.join(", ") +
                    ". Rien n'a ete enregistre - rechargez l'ecran.");
  }

  tickRange.setValues(tickVals);
  rateRange.setValues(rateVals);
  SpreadsheetApp.flush();

  const backTicks = sh.getRange(RED_CFG.START_ROW, RED_CFG.TICK_COL, n, 1).getValues();
  const backRates = sh.getRange(RED_CFG.START_ROW, RED_CFG.RATE_COL, n, 1).getValues();
  const ticked = [];
  for (let i = 0; i < n; i++) {
    if (backTicks[i][0] !== true) continue;
    const lot = String(lotVals[i][0] || "").trim();
    if (lot) ticked.push({ lot: lot, rate: backRates[i][0] });
  }
  return { changed: changed, ticked: ticked };
}
