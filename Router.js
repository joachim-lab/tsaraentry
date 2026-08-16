/***************************************************************
 * Router.js — TSARA Entry web app
 * Serves the right screen based on ?screen= in the URL.
 *
 *   ?screen=nourrissage  -> Index.html           (screen 1)
 *   ?screen=lot          -> LotIndex.html        (screen 2)
 *   ?screen=commandes    -> CommandesIndex.html  (screen 3)
 *   ?screen=morts        -> MortsIndex.html      (screen 4)
 *   ?screen=tracabilite  -> TracabiliteIndex.html (screen 5 picker)
 *   ?screen=tracform     -> LigneeUI's form, rendered live (screen 5 form)
 *   no parameter         -> Menu.html            (chooser)
 ***************************************************************/

function doGet(e) {
  const screen = (e && e.parameter && e.parameter.screen) || "";

  // Screen 5's FORM is not a file in this project: LigneeUI renders it,
  // so there is exactly one copy of the form and its rules. It arrives
  // as a complete HTML document and is served as the whole page. It is
  // NOT injected into a page — inserted HTML does not run its scripts,
  // so an injected form would render and then be dead.
  if (screen === "tracform") {
    const p = (e && e.parameter) || {};
    if (!p.lot || !p.op) throw new Error("Paramètres manquants (lot, op).");
    const html = LigneeUI.lu_dialogHtmlForLot(
      p.lot, p.op, tracCurrentOperatorEmail(),
      getWebAppUrl() + "?screen=tracabilite");
    return HtmlService.createHtmlOutput(html)
      .setTitle("TSARA Entry")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  let file;
  if (screen === "nourrissage") file = "Index";
  else if (screen === "lot") file = "LotIndex";
  else if (screen === "commandes") file = "CommandesIndex";
  else if (screen === "morts") file = "MortsIndex";
  else if (screen === "tracabilite") file = "TracabiliteIndex";
  else file = "Menu";

  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle("TSARA Entry")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** URL of this web app — used by Menu.html to build its links. */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}
