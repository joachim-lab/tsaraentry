path = "CommandesIndex.html"
with open(path) as f:
    c = f.read()

# ---- 1. boot: fetch TARIFS alongside OPTS ----
anchor1 = """  let OPTS = { lots: [], types: [] };
  let lines = [{}];            // create mode: one line per lot"""
assert c.count(anchor1) == 1
new1 = """  let OPTS = { lots: [], types: [] };
  let TARIFS = null;           // { bands:[...], fish:{detail,gros} } - loaded once at boot
  let lines = [{}];            // create mode: one line per lot"""
c = c.replace(anchor1, new1, 1)

anchor2 = """  google.script.run
    .withSuccessHandler(o => { OPTS = o; renderNew(); })
    .withFailureHandler(e => setStatus('err', 'Erreur chargement : ' + esc(e.message)))
    .cmdGetOptions();"""
assert c.count(anchor2) == 1
new2 = """  google.script.run
    .withSuccessHandler(o => { OPTS = o; renderNew(); })
    .withFailureHandler(e => setStatus('err', 'Erreur chargement : ' + esc(e.message)))
    .cmdGetOptions();

  // Tarifs load independently of OPTS - price pre-fill is a nice-to-have,
  // never blocks the form. recomputeAllPrices() no-ops while TARIFS is null.
  google.script.run
    .withSuccessHandler(t => { TARIFS = t; recomputeAllPrices(); })
    .withFailureHandler(e => setStatus('warn', 'Tarifs indisponibles : ' + esc(e.message)))
    .cmdGetTarifs();"""
c = c.replace(anchor2, new2, 1)

# ---- 2. Commande card: add order-level Tarification controls ----
anchor3 = """        '<label>Remarques / destination</label><input type="text" id="fRemarques">' +
      '</div>' +
      '<div id="linesWrap"></div>' +"""
assert c.count(anchor3) == 1
new3 = """        '<label>Remarques / destination</label><input type="text" id="fRemarques">' +
        '<div class="row">' +
          '<div id="provendeWrap"><label>Type de client</label>' +
            '<select id="fProvende" onchange="recomputeAllPrices()">' +
              '<option value="avec">Client provende</option>' +
              '<option value="sans">Pas client provende</option>' +
            '</select></div>' +
          '<div id="qualiteWrap" style="display:none"><label>Qualit\\u00e9</label>' +
            '<select id="fQualite" onchange="recomputeAllPrices()">' +
              '<option value="detail">D\\u00e9tail</option>' +
              '<option value="gros">Gros</option>' +
            '</select></div>' +
        '</div>' +
        '<div class="row">' +
          '<div><label>Livraison</label>' +
            '<select id="fLivraison" onchange="onLivraisonChange()">' +
              '<option value="enlevement">Enl\\u00e8vement \\u00e0 la ferme</option>' +
              '<option value="environs">Livraison environs de la ferme</option>' +
              '<option value="ambohim">Livraison Ambohimangakely</option>' +
            '</select></div>' +
          '<div id="kmWrap" style="display:none"><label>Km depuis Ambohimangakely</label>' +
            '<input type="text" inputmode="decimal" id="fKm" oninput="recomputeTransport()"></div>' +
        '</div>' +
      '</div>' +
      '<div id="linesWrap"></div>' +"""
c = c.replace(anchor3, new3, 1)

# ---- 3. renderLines: toggle provende/qualite visibility, recompute at the end ----
anchor4 = """    const lotOpts = OPTS.lots.map(l => '<option value="'+esc(l)+'">'+esc(l)+'</option>').join('');
    const al = isAlevins();
"""
assert c.count(anchor4) == 1
new4 = """    const lotOpts = OPTS.lots.map(l => '<option value="'+esc(l)+'">'+esc(l)+'</option>').join('');
    const al = isAlevins();
    if (el('provendeWrap')) el('provendeWrap').style.display = al ? '' : 'none';
    if (el('qualiteWrap'))  el('qualiteWrap').style.display  = al ? 'none' : '';
"""
c = c.replace(anchor4, new4, 1)

anchor5 = """      '</div>'
    ).join('');
    restoreLines();
  }"""
assert c.count(anchor5) == 1
new5 = """      '</div>'
    ).join('');
    restoreLines();
    recomputeAllPrices();
  }"""
c = c.replace(anchor5, new5, 1)

# ---- 4. line inputs: hook alevinsPm and alevinsLivrer to recompute ----
anchor6 = '''            '<div><label>PM alevins (g)</label>' +
              '<input type="text" inputmode="decimal" data-l="'+i+'" data-k="alevinsPm"></div>' +'''
assert c.count(anchor6) == 1
new6 = '''            '<div><label>PM alevins (g)</label>' +
              '<input type="text" inputmode="decimal" data-l="'+i+'" data-k="alevinsPm" oninput="recomputeLinePrice('+i+')"></div>' +'''
c = c.replace(anchor6, new6, 1)

anchor7 = '''            '<div><label>Alevins \\u00e0 livrer (+5%)</label>' +
              '<input type="text" inputmode="decimal" data-l="'+i+'" data-k="alevinsLivrer" id="livrer'+i+'">' +'''
# The source has a literal 'à' character, not an escape - match exactly as in file.
anchor7 = '''            '<div><label>Alevins à livrer (+5%)</label>' +
              '<input type="text" inputmode="decimal" data-l="'+i+'" data-k="alevinsLivrer" id="livrer'+i+'">' +'''
assert c.count(anchor7) == 1
new7 = '''            '<div><label>Alevins à livrer (+5%)</label>' +
              '<input type="text" inputmode="decimal" data-l="'+i+'" data-k="alevinsLivrer" id="livrer'+i+'" oninput="recomputeLinePrice('+i+')">' +'''
c = c.replace(anchor7, new7, 1)

# ---- 5. onLotChange: also recompute price after PM prefill ----
anchor8 = """        prefillPm(i, a.pm);
      })
      .withFailureHandler(e => { box.innerHTML = 'stock indisponible : ' + esc(e.message); })"""
assert c.count(anchor8) == 1
new8 = """        prefillPm(i, a.pm);
        recomputeLinePrice(i);
      })
      .withFailureHandler(e => { box.innerHTML = 'stock indisponible : ' + esc(e.message); })"""
c = c.replace(anchor8, new8, 1)

# ---- 6. new pricing functions, inserted right after prefillPm() ----
anchor9 = """    inp.value = String(Math.round(pm * 100) / 100).replace('.', ',');
    lines[i] = lines[i] || {};
    lines[i][k] = inp.value;
  }

  function doCreate() {"""
assert c.count(anchor9) == 1
new9 = """    inp.value = String(Math.round(pm * 100) / 100).replace('.', ',');
    lines[i] = lines[i] || {};
    lines[i][k] = inp.value;
  }

  // ---------- Tarifs : price + transport pre-fill (2026-08-12) ----------
  //
  // TARIFS is fetched once at boot (see boot section). Every lookup below
  // runs in the browser - no per-line server round trip. Same "fill only
  // when blank" rule as prefillPm: what is saved is what is on screen.

  /** First band whose max covers this PM. Null above 20g (off the grid). */
  function tarifBandForPm(pm) {
    if (!TARIFS || pm === null || !isFinite(pm) || pm <= 0) return null;
    for (let i = 0; i < TARIFS.bands.length; i++) {
      if (pm <= TARIFS.bands[i].max) return TARIFS.bands[i];
    }
    return null;
  }

  /** Delivery + quantity tier -> which column of a band to use. */
  function tarifColumnKey(livraison, tier) {
    if (livraison === 'enlevement') return 'enlevement';
    if (livraison === 'environs') return tier === 'le' ? 'environsLe' : 'environsGt';
    if (livraison === 'ambohim') return tier === 'le' ? 'ambohimLe' : 'ambohimGt';
    return null;
  }

  /** km x rate/km x 2 (aller-retour). <=50km : 860 ; >50km : 950. */
  function transportAmount(km) {
    if (km === null || !isFinite(km) || km <= 0) return null;
    const rate = km <= 50 ? 860 : 950;
    return km * rate * 2;
  }

  /** Fill a per-line field ONLY if it is currently blank. */
  function fillIfBlank(i, key, value) {
    if (value === null || value === undefined || !isFinite(value) || value === 0) return;
    const inp = document.querySelector('[data-l="'+i+'"][data-k="'+key+'"]');
    if (!inp || String(inp.value).trim() !== '') return;
    inp.value = String(value).replace('.', ',');
    lines[i] = lines[i] || {};
    lines[i][key] = inp.value;
  }

  /**
   * Prix alevin (AL) : band from this line's PM x client type x
   * livraison x this line's own H-quantity tier (<=5000 / >5000).
   * Prix / kg (GR) : flat Detail/Gros, no per-line dependency.
   */
  function recomputeLinePrice(i) {
    if (!TARIFS) return;
    if (isAlevins()) {
      const pmInp = document.querySelector('[data-l="'+i+'"][data-k="alevinsPm"]');
      const hInp = el('livrer'+i);
      if (!pmInp || !hInp) return;
      const band = tarifBandForPm(num(pmInp.value));
      if (!band) return;
      const provende = (el('fProvende') || {}).value || 'avec';
      const livraison = (el('fLivraison') || {}).value || 'enlevement';
      const h = num(hInp.value);
      const tier = (h !== null && h <= 5000) ? 'le' : 'gt';
      const colKey = tarifColumnKey(livraison, tier);
      if (!colKey) return;
      fillIfBlank(i, 'alevinsPrix', band[provende][colKey]);
    } else {
      const qualite = (el('fQualite') || {}).value || 'detail';
      fillIfBlank(i, 'prixKg', TARIFS.fish[qualite]);
    }
  }

  /** All visible lines. Called after render, and on any order-level change. */
  function recomputeAllPrices() {
    if (!TARIFS) return;
    lines.forEach((ln, i) => recomputeLinePrice(i));
    recomputeTransport();
  }

  /**
   * Transport applies to ONE delivery per order, so it is written on
   * line 1 only - matching how a single BL covers the whole order.
   * AL writes to 'transport' (J) ; GR writes to 'frais' (P), the only
   * field on that side the sheet's Q formula actually includes.
   */
  function recomputeTransport() {
    if (!TARIFS) return;
    const liv = (el('fLivraison') || {}).value;
    if (liv !== 'ambohim') return;
    const amt = transportAmount(num((el('fKm') || {}).value));
    if (amt === null) return;
    fillIfBlank(0, isAlevins() ? 'transport' : 'frais', amt);
  }

  function onLivraisonChange() {
    const liv = (el('fLivraison') || {}).value;
    const kmWrap = el('kmWrap');
    if (kmWrap) kmWrap.style.display = (liv === 'ambohim') ? '' : 'none';
    recomputeAllPrices();
  }

  function doCreate() {"""
c = c.replace(anchor9, new9, 1)

with open(path, "w") as f:
    f.write(c)
print("patched", path)
