import re

path = "CommandesServer.js"
with open(path) as f:
    c = f.read()

anchor = '''function cmdValidateOrderLines(lines) {'''
assert c.count(anchor) == 1, "anchor not unique or missing"

new_code = '''/* =============================================================
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
 * Fish: B35 "Prix poisson grossi" ; B37/C37 = Detail/Gros (MGA/kg).
 * Rows 27-32 (poisson frais sur glace) are NOT used (Kim, 2026-08-12).
 * ============================================================= */

const TARIFS_CFG = {
  SS_ID: CMD_CFG.SS_ID,
  SHEET: "Tarifs",
  BAND_START_ROWS: [3, 8, 12, 17, 22],
  FISH_ROW: 37
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
    const m = label.match(/([\\d.,]+)\\s*-\\s*([\\d.,]+)/);
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

  const fishRow = sh.getRange(TARIFS_CFG.FISH_ROW, 2, 1, 2).getValues()[0]; // B37/C37
  return {
    bands: bands,
    fish: { detail: Number(fishRow[0]) || 0, gros: Number(fishRow[1]) || 0 }
  };
}

/** RUN FROM EDITOR: sanity-check the parsed price list. */
function testTarifs() {
  const t = buildTarifs();
  Logger.log(JSON.stringify(t, null, 2));
}

''' + anchor

c2 = c.replace(anchor, new_code, 1)
assert c2 != c
with open(path, "w") as f:
    f.write(c2)
print("patched", path)
