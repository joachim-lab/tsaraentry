#!/usr/bin/env python3
# tsaraentry - wire up screen 5 (Tracabilite): Router routes + Menu tile.
#
# TWO routes, not one:
#   ?screen=tracabilite -> TracabiliteIndex.html, the picker (a file here)
#   ?screen=tracform    -> LigneeUI's form, rendered live (NOT a file here)
#
# The form is deliberately not a file in this project. LigneeUI renders
# it, so the form and its validation rules exist in exactly one place.
# It is served as its own page rather than injected into the picker:
# HTML inserted into a live page does not execute its <script>, so an
# injected form would appear complete and be entirely dead - no
# validation, no save.

ROUTER = "Router.js"
MENU = "Menu.html"

PAIRS_ROUTER = [
    (' *   ?screen=morts        -> MortsIndex.html      (screen 4)\n *   no parameter         -> Menu.html            (chooser)', " *   ?screen=morts        -> MortsIndex.html      (screen 4)\n *   ?screen=tracabilite  -> TracabiliteIndex.html (screen 5 picker)\n *   ?screen=tracform     -> LigneeUI's form, rendered live (screen 5 form)\n *   no parameter         -> Menu.html            (chooser)"),
    ('  let file;\n  if (screen === "nourrissage") file = "Index";\n  else if (screen === "lot") file = "LotIndex";\n  else if (screen === "commandes") file = "CommandesIndex";\n  else if (screen === "morts") file = "MortsIndex";\n  else file = "Menu";\n\n  return HtmlService.createHtmlOutputFromFile(file)', '  // Screen 5\'s FORM is not a file in this project: LigneeUI renders it,\n  // so there is exactly one copy of the form and its rules. It arrives\n  // as a complete HTML document and is served as the whole page. It is\n  // NOT injected into a page — inserted HTML does not run its scripts,\n  // so an injected form would render and then be dead.\n  if (screen === "tracform") {\n    const p = (e && e.parameter) || {};\n    if (!p.lot || !p.op) throw new Error("Paramètres manquants (lot, op).");\n    const html = LigneeUI.lu_dialogHtmlForLot(p.lot, p.op, tracOperatorEmail(p.who));\n    return HtmlService.createHtmlOutput(html)\n      .setTitle("TSARA Entry")\n      .addMetaTag("viewport", "width=device-width, initial-scale=1")\n      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);\n  }\n\n  let file;\n  if (screen === "nourrissage") file = "Index";\n  else if (screen === "lot") file = "LotIndex";\n  else if (screen === "commandes") file = "CommandesIndex";\n  else if (screen === "morts") file = "MortsIndex";\n  else if (screen === "tracabilite") file = "TracabiliteIndex";\n  else file = "Menu";\n\n  return HtmlService.createHtmlOutputFromFile(file)'),
]

PAIRS_MENU = [
    ('        \'<a class="tile" href="\' + url + \'?screen=morts">\' +\n          \'<h2>Mortalité</h2><p>Déclaration des poissons morts</p></a>\';', '        \'<a class="tile" href="\' + url + \'?screen=morts">\' +\n          \'<h2>Mortalité</h2><p>Déclaration des poissons morts</p></a>\' +\n        \'<a class="tile" href="\' + url + \'?screen=tracabilite">\' +\n          \'<h2>Traçabilité</h2><p>Tri, déplacement, mélange, mise en cage</p></a>\';'),
]


def patch(path, pairs):
    with open(path, "r", encoding="utf-8") as f:
        c = f.read()
    for i, (old, new) in enumerate(pairs, 1):
        assert c.count(old) == 1, "%s anchor %d not found exactly once: %d" % (path, i, c.count(old))
        assert c.count(new) == 0, "%s anchor %d already patched" % (path, i)
    for old, new in pairs:
        c = c.replace(old, new)
    for i, (old, new) in enumerate(pairs, 1):
        assert c.count(new) == 1, "%s anchor %d replacement failed" % (path, i)
    with open(path, "w", encoding="utf-8") as f:
        f.write(c)


patch(ROUTER, PAIRS_ROUTER)
patch(MENU, PAIRS_MENU)
print("OK: Router routes and Menu tile added for screen 5.")
