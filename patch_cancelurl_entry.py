#!/usr/bin/env python3
# tsaraentry - give the form the picker's URL so Cancel can return.
# history.back() cannot work: the form lives in Apps Script's sandbox
# iframe and the picker navigated the TOP window to open it, leaving the
# iframe with no history. Verified dead live on 2026-08-14.

ROUTER = "Router.js"
PAIRS = [('const html = LigneeUI.lu_dialogHtmlForLot(p.lot, p.op, tracOperatorEmail(p.who));', 'const html = LigneeUI.lu_dialogHtmlForLot(\n      p.lot, p.op, tracOperatorEmail(p.who),\n      getWebAppUrl() + "?screen=tracabilite");')]


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


patch(ROUTER, PAIRS)
print("OK: Router passes the picker URL to the form.")
