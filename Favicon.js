/***************************************************************
 * Favicon.js — the Tsara Tilapia logo on the browser tab.
 *
 * HtmlService serves no static file, so the icon needs a public URL.
 * setFaviconUrl accepts a URL ONLY when the URL ends in an image
 * extension. Proved live on 2026-09-12: a data: URI, a Drive
 * uc?export=view URL, a Drive thumbnail URL and an lh3 URL are all
 * refused with "The favicon icon image type is not supported".
 * The exception throws inside doGet, so a wrong URL kills every screen.
 *
 * The icon is the 32x32 PNG committed at the root of the public
 * tsaraentry repo. Master copy: docs/assets/tsara-favicon-32.png in
 * the tsara-erp project folder. Change one, change all three.
 ***************************************************************/

const TSARA_FAVICON = "https://raw.githubusercontent.com/joachim-lab/tsaraentry/main/favicon.png";
