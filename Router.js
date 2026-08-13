/***************************************************************
 * Router.js — TSARA Entry web app
 * Serves the right screen based on ?screen= in the URL.
 *
 *   ?screen=nourrissage  -> Index.html           (screen 1)
 *   ?screen=lot          -> LotIndex.html        (screen 2)
 *   ?screen=commandes    -> CommandesIndex.html  (screen 3)
 *   ?screen=morts        -> MortsIndex.html      (screen 4)
 *   no parameter         -> Menu.html            (chooser)
 ***************************************************************/

function doGet(e) {
  const screen = (e && e.parameter && e.parameter.screen) || "";

  let file;
  if (screen === "nourrissage") file = "Index";
  else if (screen === "lot") file = "LotIndex";
  else if (screen === "commandes") file = "CommandesIndex";
  else if (screen === "morts") file = "MortsIndex";
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
