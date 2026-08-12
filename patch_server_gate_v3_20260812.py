"""
TSARA Entry — server: lot/type gate + Stock Poisson PM map.
PRECONDITION: CommandesServer.js is 26516 bytes (state after
patch_server_tarifs_20260812.py). Supersedes patch_server_gate_20260812.py,
which must NOT be run — its 350 g floor would block every Grossis order
(heaviest live lot is 246 g).
"""
path = "CommandesServer.js"
with open(path) as f:
    c = f.read()

# ---- 1. PM map from Stock Poisson (feeds the dropdown gate) ----
a = """/**
 * Validate a whole submission before saving."""
assert c.count(a) == 1, "validator doc anchor"
c = c.replace(a, """/* =============================================================
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
 * Everything the Commandes screen needs to gate its lot dropdown:
 * the PM map plus the boundaries, so the client never hardcodes them.
 */
function cmdGetLotGateData() {
  return {
    pm: cmdGetLotPmMap(),
    fryMax: cmdFryMaxPm(),
    grBlock: CMD_GR_BLOCK_PM,
    grWarn: CMD_GR_WARN_PM
  };
}

/** RUN FROM EDITOR: what the dropdown will allow, per lot. */
function testLotGate() {
  const d = cmdGetLotGateData();
  Logger.log("fryMax=" + d.fryMax + "  grBlock=" + d.grBlock + "  grWarn=" + d.grWarn);
  Object.keys(d.pm).sort().forEach(function (k) {
    const pm = d.pm[k];
    const al = cmdLotTypeVerdict(true, pm, d.fryMax).level;
    const gr = cmdLotTypeVerdict(false, pm, d.fryMax).level;
    Logger.log(k + "  PM=" + pm + "   AL:" + al + "   GR:" + gr);
  });
}

/**
 * Validate a whole submission before saving.""", 1)

# ---- 2. Stock Poisson config block, next to the Tarifs one ----
a = """const TARIFS_CACHE_KEY = "cmd_tarifs_v1";"""
assert c.count(a) == 1, "tarifs cache key anchor"
c = c.replace(a, """/* Stock Poisson, "lot" tab: the nightly output of updateStockPoisson
   (engine_core.js). N = lot id, O = nombre, P = PM, data from row 3. */
const STOCK_PM_CFG = {
  SS_ID: "1Kfs5beQorhdheqzEDibgnBd5wQjy79MecncNlgKjRIE",
  SHEET: "lot",
  START_ROW: 3
};

const TARIFS_CACHE_KEY = "cmd_tarifs_v1";""", 1)

# ---- 3. validator signature ----
a = """ * @param {Array} lines  [{ lot, qty, pm }]  qty already H-or-N resolved
 * @return {Object} { ok, blocks: [msg], warnings: [msg], detail: {key: avail} }
 */"""
assert c.count(a) == 1, "param doc anchor"
c = c.replace(a, """ * @param {Array} lines  [{ lot, qty, pm }]  qty already H-or-N resolved
 * @param {string=} orderType  raw type ("AL"/"GR"). Enables the lot/type
 *        gate. The client shows the same verdict at selection time;
 *        this is the ENFORCING copy, because only the server runs at save.
 * @return {Object} { ok, blocks: [msg], warnings: [msg], detail: {key: avail} }
 */""", 1)

a = """function cmdValidateOrderLines(lines) {"""
assert c.count(a) == 1, "validator signature"
c = c.replace(a, """function cmdValidateOrderLines(lines, orderType) {""", 1)

# ---- 4. the gate, after the found-check ----
a = """    if (!a.found) {
      warnings.push(k + " : ce lot n'est pas trouvé dans le fichier lot ; " +
                        "la déduction échouera cette nuit");
      return;
    }"""
assert c.count(a) == 1, "found-check anchor"
c = c.replace(a, """    if (!a.found) {
      warnings.push(k + " : ce lot n'est pas trouvé dans le fichier lot ; " +
                        "la déduction échouera cette nuit");
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
    }""", 1)

with open(path, "w") as f:
    f.write(c)
print("patched", path)
