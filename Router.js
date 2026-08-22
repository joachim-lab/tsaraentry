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
 *   ?screen=creerlot     -> CreerLotIndex.html    (screen 6)
 *   ?screen=temperatures -> TemperaturesIndex.html (screen 7)
 *   ?screen=gestioninventaires -> GestionInventairesIndex.html (screen 8 picker)
 *   ?screen=inventaire   -> InventaireIndex.html  (screen 8a)
 *   ?screen=checkstock   -> CheckStockIndex.html  (screen 8b)
 *   ?screen=achats       -> AchatsIndex.html      (screen 8c)
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
      getWebAppUrl() + "?screen=tracabilite",
      String(p.src || ""));   // "26,30": lots absorbés (regroupement)
    return HtmlService.createHtmlOutput(withSpinner(html))
      .setTitle("Interface Tsara Tilapia")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  let file;
  if (screen === "nourrissage") file = "Index";
  else if (screen === "lot") file = "LotIndex";
  else if (screen === "commandes") file = "CommandesIndex";
  else if (screen === "morts") file = "MortsIndex";
  else if (screen === "tracabilite") file = "TracabiliteIndex";
  else if (screen === "creerlot") file = "CreerLotIndex";
  else if (screen === "temperatures") file = "TemperaturesIndex";
  else if (screen === "gestioninventaires") file = "GestionInventairesIndex";
  else if (screen === "inventaire") file = "InventaireIndex";
  else if (screen === "checkstock") file = "CheckStockIndex";
  else if (screen === "achats") file = "AchatsIndex";
  else file = "Menu";

  return HtmlService.createTemplateFromFile(file).evaluate()
    .setTitle("Interface Tsara Tilapia")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** URL of this web app — used by Menu.html to build its links. */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/** Inserts a shared HTML partial into a templated screen. */
function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

/**
 * Inserts the shared busy overlay into a page that was built outside
 * this project (the Tracabilite form, built by LigneeUI). The partial
 * must go before </head>, because the wrapper has to be installed
 * before any script of that page can call the server.
 */
function withSpinner(html) {
  const marker = "</head>";
  const i = html.indexOf(marker);
  if (i < 0) {
    throw new Error("Spinner : balise </head> absente dans la page recue.");
  }
  return html.slice(0, i) + include("Spinner") + html.slice(i);
}
