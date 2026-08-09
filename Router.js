/***************************************************************
 * Router.js — TSARA Entry web app
 * Serves the right screen based on ?screen= in the URL.
 *
 * Replaces the doGet that used to live in Code.js (which always
 * served the Nourrissage screen). Code.js's own doGet must be
 * removed when this file is added — two doGet functions in one
 * project is a conflict.
 *
 *   ?screen=nourrissage  -> Index.html    (screen 1)
 *   ?screen=lot          -> LotIndex.html (screen 2)
 *   no parameter         -> Menu.html     (chooser)
 ***************************************************************/

function doGet(e) {
  const screen = (e && e.parameter && e.parameter.screen) || "";

  let file;
  if (screen === "nourrissage") file = "Index";
  else if (screen === "lot") file = "LotIndex";
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
