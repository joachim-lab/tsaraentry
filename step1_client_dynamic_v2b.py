path = "CommandesIndex.html"
with open(path) as f:
    c = f.read()

# ============ 1. livraison row gets an id, so GR can hide it ============
a = """        '<div class="row">' +
          '<div><label>Livraison</label>' +"""
assert c.count(a) == 1
c = c.replace(a, """        '<div class="row" id="livRow">' +
          '<div><label>Livraison</label>' +""", 1)

# ============ 2. renderLines: GR hides the whole livraison row ============
a = """    if (el('provendeWrap')) el('provendeWrap').style.display = al ? '' : 'none';
    if (el('qualiteWrap'))  el('qualiteWrap').style.display  = al ? 'none' : '';"""
assert c.count(a) == 1
c = c.replace(a, """    if (el('provendeWrap')) el('provendeWrap').style.display = al ? '' : 'none';
    if (el('qualiteWrap'))  el('qualiteWrap').style.display  = al ? 'none' : '';
    // GR has no livraison logic at all: flat Detail/Gros price, frais
    // additionnels typed by hand (Kim 2026-08-12). Hide the whole row.
    if (el('livRow')) el('livRow').style.display = al ? '' : 'none';""", 1)

# ============ 3. syncLivrer feeds the tier -> recompute the price ============
a = """    if (!lvInp.dataset.touched) lvInp.value = Math.round(n * 1.05);
  }"""
assert c.count(a) == 1
c = c.replace(a, """    if (!lvInp.dataset.touched) lvInp.value = Math.round(n * 1.05);
    // H decides the <=5000 / >5000 price tier, so a qty change can
    // change the price - and this programmatic write fires no input
    // event, so recompute explicitly.
    recomputeLinePrice(i);
  }""", 1)

# ============ 4. onLotChange: lot/type compatibility gate ============
a = """        box.className = 'muted';
        box.innerHTML = 'disponible : <b>' + fmtNum(a.available) +
                        '</b> \\u00b7 PM : <b>' + fmtNum(a.pm) + ' g</b>';
        prefillPm(i, a.pm);
        recomputeLinePrice(i);"""
assert c.count(a) == 1
c = c.replace(a, """        // ---- lot/type gate (Kim 2026-08-12) ----
        // A fry order against a lot of harvest-size fish (or a fish
        // order against a fry lot) is always a mistake. Red-block here,
        // clear any auto-filled PM/price from a previous selection, and
        // the server blocks it again at save (cmdValidateOrderLines).
        const alSel = isAlevins();
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
        }
        box.className = 'muted';
        box.innerHTML = 'disponible : <b>' + fmtNum(a.available) +
                        '</b> \\u00b7 PM : <b>' + fmtNum(a.pm) + ' g</b>';
        prefillPm(i, a.pm);
        recomputeLinePrice(i);""", 1)

# ============ 5. replace prefillPm + whole Tarifs block with the dynamic model ============
start = """  /**
   * Put the lot's PM into this line's PM field - but ONLY when the user
   * has left it empty. Never overwrite a typed value."""
end = """  function doCreate() {"""
i0 = c.find(start)
i1 = c.find(end)
assert i0 != -1 and i1 != -1 and i0 < i1

new_block = """  /**
   * OWNERSHIP MODEL for pre-filled fields (2026-08-12, replaces
   * "fill only when blank").
   *
   * "Fill only when blank" froze the form: the first keystroke of a km
   * value wrote 3km worth of transport, the field was no longer blank,
   * and every later change (provende, livraison, Gros, the rest of the
   * km) was refused. One cause, five observed bugs.
   *
   * Rule now: every managed field is either
   *   AUTO  - the app wrote it; the app may rewrite or clear it when
   *           inputs change,
   *   USER  - a human typed in it; the app never touches it again.
   * Typing in a field makes it USER. Clearing it hands it back to AUTO
   * on the next recompute. Flags live in lines[i].__auto so they follow
   * their line through add/remove splices.
   *
   * "What is saved is what is on screen" still holds: the app only ever
   * writes into the visible fields, never behind them.
   */
  const MANAGED_KEYS = { alevinsPm:1, alevinsPrix:1, transport:1, poissonPm:1, prixKg:1 };

  document.addEventListener('input', e => {
    const t = e.target;
    if (!t || !t.getAttribute) return;
    const k = t.getAttribute('data-k');
    const l = t.getAttribute('data-l');
    if (k && l !== null && MANAGED_KEYS[k]) {
      const i = Number(l);
      if (lines[i] && lines[i].__auto) delete lines[i].__auto[k];
    }
  });

  /**
   * Write `value` into line i's field IF the field is auto-owned or
   * blank. value === null means "no longer applicable": clear the field
   * only if the app owns it.
   */
  function setAuto(i, key, value) {
    const inp = document.querySelector('[data-l="'+i+'"][data-k="'+key+'"]');
    if (!inp) return;
    lines[i] = lines[i] || {};
    lines[i].__auto = lines[i].__auto || {};
    const isAuto = !!lines[i].__auto[key];
    const isBlank = String(inp.value).trim() === '';
    if (!isAuto && !isBlank) return;                 // USER-owned: never touch
    if (value === null || value === undefined || !isFinite(value) || value === 0) {
      if (isAuto) { inp.value = ''; delete lines[i].__auto[key]; lines[i][key] = ''; }
      return;
    }
    inp.value = String(Math.round(value * 100) / 100).replace('.', ',');
    lines[i].__auto[key] = true;
    lines[i][key] = inp.value;
  }

  /** Lot PM -> this line's PM field, through the same ownership rules. */
  function prefillPm(i, pm) {
    if (pm === null || pm === undefined || !isFinite(pm) || pm <= 0) return;
    setAuto(i, isAlevins() ? 'alevinsPm' : 'poissonPm', pm);
  }

  // ---------- Tarifs : dynamic price + transport (2026-08-12 v2) ----------

  /** First band whose max covers this PM. Null above the grid. */
  function tarifBandForPm(pm) {
    if (!TARIFS || pm === null || !isFinite(pm) || pm <= 0) return null;
    for (let i = 0; i < TARIFS.bands.length; i++) {
      if (pm <= TARIFS.bands[i].max) return TARIFS.bands[i];
    }
    return null;
  }

  /** Top of the fry grid (20 g today) - read from the sheet, not hardcoded. */
  function alMaxPm() {
    return TARIFS ? TARIFS.bands[TARIFS.bands.length - 1].max : 20;
  }
  /** Harvest weight floor for grossis orders. */
  const GR_MIN_PM = 350;

  /** Delivery + quantity tier -> which column of a band to use. */
  function tarifColumnKey(livraison, tier) {
    if (livraison === 'enlevement') return 'enlevement';
    if (livraison === 'environs') return tier === 'le' ? 'environsLe' : 'environsGt';
    if (livraison === 'ambohim')  return tier === 'le' ? 'ambohimLe'  : 'ambohimGt';
    return null;
  }

  /** km x rate/km x 2 (aller-retour). <=50 km : 860 ; >50 km : 950. */
  function transportAmount(km) {
    if (km === null || !isFinite(km) || km <= 0) return null;
    return km * (km <= 50 ? 860 : 950) * 2;
  }

  /**
   * One line's price, recomputed from CURRENT inputs, written through
   * setAuto - so it tracks every change and clears when off the grid.
   *
   * AL : band(PM) x provende x livraison x tier(H<=5000).
   *      H unknown -> 'gt' tier: the cheaper <=5000 price is only
   *      granted once the quantity is known to qualify.
   * GR : flat Detail/Gros. No livraison, no bands.
   */
  function recomputeLinePrice(i) {
    if (!TARIFS) return;
    if (isAlevins()) {
      const pmInp = document.querySelector('[data-l="'+i+'"][data-k="alevinsPm"]');
      const hInp = el('livrer'+i);
      if (!pmInp || !hInp) return;
      const band = tarifBandForPm(num(pmInp.value));
      let price = null;
      if (band) {
        const provende = (el('fProvende') || {}).value || 'avec';
        const livraison = (el('fLivraison') || {}).value || 'enlevement';
        const h = num(hInp.value);
        const tier = (h !== null && h <= 5000) ? 'le' : 'gt';
        price = band[provende][tarifColumnKey(livraison, tier)] || null;
      }
      setAuto(i, 'alevinsPrix', price);
    } else {
      const qualite = (el('fQualite') || {}).value || 'detail';
      setAuto(i, 'prixKg', TARIFS.fish[qualite] || null);
    }
  }

  /** All lines + transport. Called after render and on any order-level change. */
  function recomputeAllPrices() {
    if (!TARIFS) return;
    lines.forEach((ln, i) => recomputeLinePrice(i));
    recomputeTransport();
  }

  /**
   * Transport: AL orders only, Ambohimangakely only, line 1 only (one
   * delivery per order). Any other state CLEARS an auto-filled amount -
   * switching back to Enlevement must not leave a stale charge behind.
   * GR : frais additionnels is typed by hand, never auto.
   */
  function recomputeTransport() {
    if (!isAlevins()) return;
    const liv = (el('fLivraison') || {}).value;
    const amt = (liv === 'ambohim')
      ? transportAmount(num((el('fKm') || {}).value))
      : null;
    setAuto(0, 'transport', amt);
  }

  function onLivraisonChange() {
    const liv = (el('fLivraison') || {}).value;
    const kmWrap = el('kmWrap');
    if (kmWrap) kmWrap.style.display = (liv === 'ambohim') ? '' : 'none';
    recomputeAllPrices();
  }

  function doCreate() {"""

c = c[:i0] + new_block + c[i1 + len(end):]
# the replacement already ends with the doCreate opener, so re-add nothing

# ============ 6. pass the order type to the server validator ============
a = """      .cmdValidateOrderLines(vLines);"""
assert c.count(a) == 1
c = c.replace(a, """      .cmdValidateOrderLines(vLines, el('fType').value);""", 1)

with open(path, "w") as f:
    f.write(c)
print("patched", path)
