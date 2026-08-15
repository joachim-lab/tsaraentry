#!/usr/bin/env python3
# patch_whoami_20260815f.py
# Step 1b - TEMPORARY identity diagnostic on the TSARA Entry menu page.
#   Router.js : adds whoAmITest()
#   Menu.html : adds a small grey line under the tiles showing the result
# Remove both once the answer is recorded.
#
# Run from:  ~/tsara/tsaraentry

import io
import os
import sys

FILES = ["Router.js", "Menu.html"]

for f in FILES:
    if not os.path.exists(f):
        sys.exit("ERREUR: %s introuvable. Lancez le script depuis ~/tsara/tsaraentry" % f)


def load(p):
    with io.open(p, "r", encoding="utf-8", newline="") as fh:
        return fh.read()


def save(p, c):
    with io.open(p, "w", encoding="utf-8", newline="") as fh:
        fh.write(c)


def swap(content, old, new, label):
    n = content.count(old)
    assert n == 1, "ANCRE NON UNIQUE (%s): %d occurrence(s), attendu 1" % (label, n)
    return content.replace(old, new)


# ---------------------------------------------------------------- Router.js

r = load("Router.js")

R_OLD = """function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}"""

R_NEW = R_OLD + """

/**
 * TEMPORAIRE - test d'identite Step 1b. A supprimer une fois la reponse notee.
 *
 * active    = l'utilisateur connecte, mais uniquement s'il appartient au meme
 *             domaine que le proprietaire du script. Chaine vide sinon.
 * effective = le compte sous lequel le script s'execute. Avec "Execute as: Me"
 *             c'est toujours le proprietaire. Sert de temoin: prouve que
 *             l'appel a bien atteint le serveur.
 */
function whoAmITest() {
  var active = '';
  try {
    active = Session.getActiveUser().getEmail() || '';
  } catch (e) {
    active = 'erreur: ' + e.message;
  }
  var effective = '';
  try {
    effective = Session.getEffectiveUser().getEmail() || '';
  } catch (e) {
    effective = 'erreur: ' + e.message;
  }
  return { active: active, effective: effective };
}"""

r = swap(r, R_OLD, R_NEW, "Router.js / getWebAppUrl")
save("Router.js", r)

# ---------------------------------------------------------------- Menu.html

m = load("Menu.html")

# 1. the display line, inserted just above the script block
M1_OLD = "  <script>\n"
M1_NEW = ('  <div id="whoami" '
          'style="margin-top:22px;font-size:12px;color:#5f6368;'
          'border-top:1px solid #dadce0;padding-top:10px;">'
          'identite : chargement</div>\n'
          '\n'
          '  <script>\n')
m = swap(m, M1_OLD, M1_NEW, "Menu.html / balise script")

# 2. the second server call, appended after the existing getWebAppUrl call
M2_OLD = """    }).getWebAppUrl();
  </script>"""

M2_NEW = """    }).getWebAppUrl();

    google.script.run
      .withSuccessHandler(function (r) {
        document.getElementById('whoami').textContent =
          'active: ' + (r.active || '(VIDE)') +
          '  |  effective: ' + (r.effective || '(VIDE)');
      })
      .withFailureHandler(function (err) {
        document.getElementById('whoami').textContent =
          'echec appel serveur: ' + err.message;
      })
      .whoAmITest();
  </script>"""

m = swap(m, M2_OLD, M2_NEW, "Menu.html / appel getWebAppUrl")
save("Menu.html", m)

# ---------------------------------------------------------------- report

print("OK - patch applique")
for f in FILES:
    print("  %-14s %8d octets" % (f, os.path.getsize(f)))
