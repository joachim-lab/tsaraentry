"""
TSARA Entry — grey out TOUT-reserved lots in the Commandes dropdown.
PRECONDITION: CommandesServer.js 32501 bytes, CommandesIndex.html 34755.
"""
import io

# ---------------- SERVER ----------------
p = "CommandesServer.js"
c = open(p).read()

a = """/**
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
}"""
assert c.count(a) == 1, "gate data anchor"
c = c.replace(a, """/**
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

/**
 * Everything the Commandes screen needs to gate its lot dropdown:
 * the PM map, the TOUT reservations, and the boundaries - so the
 * client never hardcodes any of them.
 */
function cmdGetLotGateData() {
  return {
    pm: cmdGetLotPmMap(),
    reservedAll: buildReservedAllMap(),
    fryMax: cmdFryMaxPm(),
    grBlock: CMD_GR_BLOCK_PM,
    grWarn: CMD_GR_WARN_PM
  };
}""", 1)

# testLotGate should show reservations too
a = """function testLotGate() {
  const d = cmdGetLotGateData();
  Logger.log("fryMax=" + d.fryMax + "  grBlock=" + d.grBlock + "  grWarn=" + d.grWarn);
  Object.keys(d.pm).sort().forEach(function (k) {
    const pm = d.pm[k];
    const al = cmdLotTypeVerdict(true, pm, d.fryMax).level;
    const gr = cmdLotTypeVerdict(false, pm, d.fryMax).level;
    Logger.log(k + "  PM=" + pm + "   AL:" + al + "   GR:" + gr);
  });
}"""
assert c.count(a) == 1, "testLotGate anchor"
c = c.replace(a, """function testLotGate() {
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
}""", 1)

open(p, "w").write(c)
print("patched", p)

# ---------------- CLIENT ----------------
p = "CommandesIndex.html"
c = open(p).read()

a = """    const lotOpts = OPTS.lots.map(l => {
      const pm = (GATE && GATE.pm) ? GATE.pm[canonKey(l)] : undefined;
      if (pm === undefined) return '<option value="'+esc(l)+'">'+esc(l)+'</option>';
      const v = lotTypeVerdict(alSel, pm);
      const tag = ' \\u00b7 ' + fmtNum(pm) + ' g' +
        (v.level === 'block' ? ' \\u2014 incompatible' : (v.level === 'warn' ? ' \\u2014 \\u00e0 confirmer' : ''));
      return '<option value="'+esc(l)+'"' + (v.level === 'block' ? ' disabled' : '') + '>' +
             esc(l) + tag + '</option>';
    }).join('');"""
assert c.count(a) == 1, "lotOpts anchor"
c = c.replace(a, """    const lotOpts = OPTS.lots.map(l => {
      const key = canonKey(l);
      // A TOUT reservation blocks regardless of PM or type, so it is
      // tested first and its label wins - "incompatible" would send
      // someone looking at the wrong problem.
      const isReserved = !!(GATE && GATE.reservedAll && GATE.reservedAll[key]);
      const pm = (GATE && GATE.pm) ? GATE.pm[key] : undefined;
      if (isReserved) {
        return '<option value="'+esc(l)+'" disabled>' + esc(l) +
               (pm === undefined ? '' : ' \\u00b7 ' + fmtNum(pm) + ' g') +
               ' \\u2014 r\\u00e9serv\\u00e9</option>';
      }
      if (pm === undefined) return '<option value="'+esc(l)+'">'+esc(l)+'</option>';
      const v = lotTypeVerdict(alSel, pm);
      const tag = ' \\u00b7 ' + fmtNum(pm) + ' g' +
        (v.level === 'block' ? ' \\u2014 incompatible' : (v.level === 'warn' ? ' \\u2014 \\u00e0 confirmer' : ''));
      return '<option value="'+esc(l)+'"' + (v.level === 'block' ? ' disabled' : '') + '>' +
             esc(l) + tag + '</option>';
    }).join('');""", 1)

open(p, "w").write(c)
print("patched", p)
