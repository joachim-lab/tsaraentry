"""
TSARA Entry — client: dropdown gate (grey out incompatible lots) +
three-level verdict at selection.
PRECONDITION: run AFTER patch_client_dynamic_20260812.py.
CommandesIndex.html must be 32827 bytes when this starts.
"""
path = "CommandesIndex.html"
with open(path) as f:
    c = f.read()

# ---- 1. state + boot fetch ----
a = """  let TARIFS = null;           // { bands:[...], fish:{detail,gros} } - loaded once at boot"""
assert c.count(a) == 1, "TARIFS state anchor"
c = c.replace(a, """  let TARIFS = null;           // { bands:[...], fish:{detail,gros} } - loaded once at boot
  let GATE = null;             // { pm:{key:PM}, fryMax, grBlock, grWarn } - lot/type gate""", 1)

a = """  google.script.run
    .withSuccessHandler(t => { TARIFS = t; recomputeAllPrices(); })
    .withFailureHandler(e => setStatus('warn', 'Tarifs indisponibles : ' + esc(e.message)))
    .cmdGetTarifs();"""
assert c.count(a) == 1, "tarifs boot anchor"
c = c.replace(a, """  google.script.run
    .withSuccessHandler(t => { TARIFS = t; recomputeAllPrices(); })
    .withFailureHandler(e => setStatus('warn', 'Tarifs indisponibles : ' + esc(e.message)))
    .cmdGetTarifs();

  // Lot/type gate data: one read of Stock Poisson, all lots at once.
  // If it fails the dropdown simply stays unfiltered - the per-lot check
  // at selection and the server check at save both still run.
  google.script.run
    .withSuccessHandler(g => { GATE = g; if (el('linesWrap')) renderLines(); })
    .withFailureHandler(() => { GATE = null; })
    .cmdGetLotGateData();""", 1)

# ---- 2. verdict helper (mirrors cmdLotTypeVerdict on the server) ----
a = """  /** First band whose max covers this PM. Null above the grid. */"""
assert c.count(a) == 1, "band fn anchor"
c = c.replace(a, """  /**
   * Client mirror of cmdLotTypeVerdict. Boundaries come from GATE, which
   * comes from the server - never hardcoded here, so the two copies
   * cannot drift apart. Returns 'ok' when GATE has not loaded.
   */
  function lotTypeVerdict(isAl, pm) {
    if (!GATE || pm === null || pm === undefined || !isFinite(pm)) return { level: 'ok', msg: '' };
    if (isAl) {
      if (pm > GATE.fryMax) return { level: 'block',
        msg: 'PM du lot ' + fmtNum(pm) + ' g > ' + fmtNum(GATE.fryMax) +
             ' g \\u2014 ce n\\u2019est pas un lot alevins.' };
      return { level: 'ok', msg: '' };
    }
    if (pm < GATE.grBlock) return { level: 'block',
      msg: 'PM du lot ' + fmtNum(pm) + ' g < ' + fmtNum(GATE.grBlock) +
           ' g \\u2014 trop petit pour \\u00eatre vendu en grossis.' };
    if (pm < GATE.grWarn) return { level: 'warn',
      msg: 'PM du lot ' + fmtNum(pm) + ' g \\u2014 en dessous de ' + fmtNum(GATE.grWarn) +
           ' g ; vente possible mais \\u00e0 confirmer.' };
    return { level: 'ok', msg: '' };
  }

  /** Canon key, mirroring cmdCanonKey: trim, upper, collapse spaces. */
  function canonKey(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/\\u00A0/g, ' ').trim().toUpperCase().replace(/\\s+/g, '');
  }

  /** First band whose max covers this PM. Null above the grid. */""", 1)

# ---- 3. dropdown: disable incompatible lots ----
a = """    const lotOpts = OPTS.lots.map(l => '<option value="'+esc(l)+'">'+esc(l)+'</option>').join('');"""
assert c.count(a) == 1, "lotOpts anchor"
c = c.replace(a, """    // Lots incompatible with the chosen type are shown but DISABLED -
    // visible and greyed, never hidden. A lot missing from Stock Poisson
    // has no PM, so it stays selectable and falls through to the normal
    // per-lot check: a gap in one sheet must not make a real lot vanish.
    const alSel = isAlevins();
    const lotOpts = OPTS.lots.map(l => {
      const pm = (GATE && GATE.pm) ? GATE.pm[canonKey(l)] : undefined;
      if (pm === undefined) return '<option value="'+esc(l)+'">'+esc(l)+'</option>';
      const v = lotTypeVerdict(alSel, pm);
      const tag = ' \\u00b7 ' + fmtNum(pm) + ' g' +
        (v.level === 'block' ? ' \\u2014 incompatible' : (v.level === 'warn' ? ' \\u2014 \\u00e0 confirmer' : ''));
      return '<option value="'+esc(l)+'"' + (v.level === 'block' ? ' disabled' : '') + '>' +
             esc(l) + tag + '</option>';
    }).join('');""", 1)

# ---- 4. selection-time verdict replaces the two hardcoded checks ----
a = """        const alSel = isAlevins();
        if (alSel && a.pm !== null && a.pm > alMaxPm()) {
          box.className = 'availBlock';
          box.innerHTML = '\\u26d4 PM du lot ' + fmtNum(a.pm) + ' g \\u2014 au-dessus de ' +
                          fmtNum(alMaxPm()) + ' g, ce n\\u2019est pas un lot alevins. ' +
                          'La commande sera refus\\u00e9e.';
          setAuto(i, 'alevinsPm', null);
          setAuto(i, 'alevinsPrix', null);
          return;
        }
        if (!alSel && a.pm !== null && a.pm < GR_MIN_PM) {
          box.className = 'availBlock';
          box.innerHTML = '\\u26d4 PM du lot ' + fmtNum(a.pm) + ' g \\u2014 en dessous du poids ' +
                          'de r\\u00e9colte (' + fmtNum(GR_MIN_PM) + ' g). ' +
                          'La commande sera refus\\u00e9e.';
          setAuto(i, 'poissonPm', null);
          return;
        }"""
assert c.count(a) == 1, "selection gate anchor"
c = c.replace(a, """        const alSel2 = isAlevins();
        const verdict = lotTypeVerdict(alSel2, a.pm);
        if (verdict.level === 'block') {
          box.className = 'availBlock';
          box.innerHTML = '\\u26d4 ' + verdict.msg + ' La commande sera refus\\u00e9e.';
          setAuto(i, alSel2 ? 'alevinsPm' : 'poissonPm', null);
          if (alSel2) setAuto(i, 'alevinsPrix', null);
          return;
        }""", 1)

# ---- 5. warn level: keep the price, show the flag under the readout ----
a = """        box.className = 'muted';
        box.innerHTML = 'disponible : <b>' + fmtNum(a.available) +
                        '</b> \\u00b7 PM : <b>' + fmtNum(a.pm) + ' g</b>';
        prefillPm(i, a.pm);
        recomputeLinePrice(i);"""
assert c.count(a) == 1, "readout anchor"
c = c.replace(a, """        box.className = (verdict.level === 'warn') ? 'availWarn' : 'muted';
        box.innerHTML = (verdict.level === 'warn' ? '\\u26a0 ' + verdict.msg + '<br>' : '') +
                        'disponible : <b>' + fmtNum(a.available) +
                        '</b> \\u00b7 PM : <b>' + fmtNum(a.pm) + ' g</b>';
        prefillPm(i, a.pm);
        recomputeLinePrice(i);""", 1)

# ---- 6. drop the now-unused local constants ----
a = """  /** Top of the fry grid (20 g today) - read from the sheet, not hardcoded. */
  function alMaxPm() {
    return TARIFS ? TARIFS.bands[TARIFS.bands.length - 1].max : 20;
  }
  /** Harvest weight floor for grossis orders. */
  const GR_MIN_PM = 350;

"""
assert c.count(a) == 1, "old constants anchor"
c = c.replace(a, "", 1)

with open(path, "w") as f:
    f.write(c)
print("patched", path)
