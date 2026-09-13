/***************************************************************
 * ReductionServer.js - feed reduction box on the Nourrissage screen.
 *
 * Stock poisson, tab "lot":
 *   AF (32) = per-lot tick "Reduire"
 *   AH1     = ONE reduction fraction for the whole selection (0.5 = 50 %)
 * Column S already consumes both:
 *   S = (O*P)*R*IF(AND($AH$1<>"",AF=TRUE),1-$AH$1,1)
 * This file only writes those two cells. No feeding maths lives here.
 ***************************************************************/

const RED_CFG = {
  SHEET: "lot",
  START_ROW: 3,
  END_ROW: 50,
  LOT_COL: 14,    // N
  PM_COL: 16,     // P  = poids moyen (g)
  TICK_COL: 32,   // AF = Reduire
  PCT_ROW: 1,     // AH1
  PCT_COL: 34,
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
 * ticked. A lot ticked directly in the sheet must never stay invisible
 * here - AH1 applies to it too.
 * Returns { pct, rows: [{ lot, pm, reduce }] }. pct is a percentage (50), not
 * the stored fraction (0.5).
 */
function getReductionLots() {
  const sh = redOpenLotSheet();
  const n = RED_CFG.END_ROW - RED_CFG.START_ROW + 1;
  const block = sh.getRange(RED_CFG.START_ROW, RED_CFG.LOT_COL, n,
                            RED_CFG.TICK_COL - RED_CFG.LOT_COL + 1).getValues();

  const rows = [];
  for (let i = 0; i < n; i++) {
    const lot = String(block[i][0] || "").trim();
    if (!lot) continue;
    const pm = Number(block[i][RED_CFG.PM_COL - RED_CFG.LOT_COL]);
    const tick = block[i][RED_CFG.TICK_COL - RED_CFG.LOT_COL] === true;
    if (!(isFinite(pm) && pm > RED_CFG.MIN_PM) && !tick) continue;
    rows.push({ lot: lot, pm: isFinite(pm) ? Math.round(pm * 10) / 10 : "", reduce: tick });
  }

  const raw = sh.getRange(RED_CFG.PCT_ROW, RED_CFG.PCT_COL).getValue();
  const pct = (raw === "" || raw === null || !isFinite(Number(raw)))
    ? "" : Math.round(Number(raw) * 1000) / 10;

  return { pct: pct, rows: rows };
}

/**
 * req = { pct: Number (0-100), lots: [{ lot, reduce }] }
 * Writes AF for the listed lots only - a tick on a row the screen did not
 * show is left alone - then writes AH1 = pct/100.
 * Rows are re-read here: the weekly engine re-sorts them, so a row index
 * captured when the screen loaded cannot be trusted.
 */
function saveReduction(req) {
  if (!req || !req.lots || !req.lots.length) throw new Error("Aucun lot a enregistrer.");
  const pct = Number(req.pct);
  if (!isFinite(pct) || pct < 0 || pct > 100) {
    throw new Error("Pourcentage invalide (attendu 0 a 100, recu: " + req.pct + ").");
  }

  const sh = redOpenLotSheet();
  const n = RED_CFG.END_ROW - RED_CFG.START_ROW + 1;
  const lotVals = sh.getRange(RED_CFG.START_ROW, RED_CFG.LOT_COL, n, 1).getValues();
  const tickRange = sh.getRange(RED_CFG.START_ROW, RED_CFG.TICK_COL, n, 1);
  const tickVals = tickRange.getValues();

  const rowByLot = {};
  for (let i = 0; i < n; i++) {
    const lot = String(lotVals[i][0] || "").trim();
    if (lot) rowByLot[lot] = i;
  }

  const missing = [];
  let changed = 0;
  req.lots.forEach(function (item) {
    const lot = String(item.lot || "").trim();
    const want = item.reduce === true;
    if (!(lot in rowByLot)) { missing.push(lot); return; }
    const i = rowByLot[lot];
    if ((tickVals[i][0] === true) !== want) changed++;
    tickVals[i][0] = want;
  });

  if (missing.length) {
    throw new Error("Lot(s) introuvable(s) dans Stock poisson: " + missing.join(", ") +
                    ". Rien n'a ete enregistre - rechargez l'ecran.");
  }

  tickRange.setValues(tickVals);
  sh.getRange(RED_CFG.PCT_ROW, RED_CFG.PCT_COL).setValue(pct / 100);
  SpreadsheetApp.flush();

  return { changed: changed, pct: pct };
}
