path = "CommandesIndex.html"
with open(path) as f:
    c = f.read()

# ---- 1. CSS: two new readout states ----
a1 = """    .muted { color: #5f6368; font-size: 12px; font-weight: normal; }"""
assert c.count(a1) == 1
n1 = """    .muted { color: #5f6368; font-size: 12px; font-weight: normal; }
    /* Lot readout states. A blocking condition must not render in the same
       12px grey as an informational one - a reserved lot silently stops PM
       and price pre-fill, and that cause has to be visible (Kim 2026-08-12). */
    .availBlock { margin-top: 6px; padding: 8px 10px; border-radius: 4px;
                  background: #fce8e6; color: #c5221f; font-size: 15px;
                  font-weight: bold; border-left: 4px solid #c5221f; }
    .availWarn  { margin-top: 6px; padding: 8px 10px; border-radius: 4px;
                  background: #fef7e0; color: #b06000; font-size: 14px;
                  font-weight: bold; }"""
c = c.replace(a1, n1, 1)

# ---- 2. readout: set class per state, not a fixed 'muted' ----
a2 = """        if (a.reservedAll) { box.innerHTML = '<b>lot r\\u00e9serv\\u00e9 (TOUT)</b>'; return; }
        if (!a.found) { box.innerHTML = 'lot introuvable dans le fichier lot'; return; }
        box.innerHTML = 'disponible : <b>' + fmtNum(a.available) +
                        '</b> \\u00b7 PM : <b>' + fmtNum(a.pm) + ' g</b>';"""
assert c.count(a2) == 1
n2 = """        if (a.reservedAll) {
          box.className = 'availBlock';
          box.innerHTML = '\\u26d4 LOT R\\u00c9SERV\\u00c9 (TOUT) \\u2014 choisir un autre lot.' +
                          '<div style="font-weight:normal;font-size:13px;margin-top:4px">' +
                          'Ni le PM ni le prix ne seront pr\\u00e9-remplis, et la commande ' +
                          'sera refus\\u00e9e \\u00e0 l\\'enregistrement.</div>';
          return;
        }
        if (!a.found) {
          box.className = 'availWarn';
          box.innerHTML = '\\u26a0 Lot introuvable dans le fichier lot \\u2014 ' +
                          'la d\\u00e9duction \\u00e9chouera cette nuit.';
          return;
        }
        box.className = 'muted';
        box.innerHTML = 'disponible : <b>' + fmtNum(a.available) +
                        '</b> \\u00b7 PM : <b>' + fmtNum(a.pm) + ' g</b>';"""
c = c.replace(a2, n2, 1)

# ---- 3. reset class on the transient/empty/error states too ----
a3 = """    if (!key) { box.innerHTML = ''; return; }
    box.innerHTML = 'recherche du stock\\u2026';"""
assert c.count(a3) == 1
n3 = """    if (!key) { box.className = 'muted'; box.innerHTML = ''; return; }
    box.className = 'muted';
    box.innerHTML = 'recherche du stock\\u2026';"""
c = c.replace(a3, n3, 1)

a4 = """      .withFailureHandler(e => { box.innerHTML = 'stock indisponible : ' + esc(e.message); })"""
assert c.count(a4) == 1
n4 = """      .withFailureHandler(e => {
        box.className = 'availWarn';
        box.innerHTML = 'Stock indisponible : ' + esc(e.message);
      })"""
c = c.replace(a4, n4, 1)

with open(path, "w") as f:
    f.write(c)
print("patched", path)
