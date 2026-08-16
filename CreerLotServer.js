/***************************************************************
 * CreerLotServer.js — TSARA Entry web app
 * Screen 6: Créer un lot — server side.
 *
 * DELIBERATELY THIN. The numbering, the broodstock-name rules, the
 * duplicate check, the file copy and the root-cohort mint all live in
 * createLotFile inside the CreationLot library (Apps Script project
 * "creationnouveaufichierlot"), bound at HEAD. The Google Form trigger
 * calls the same function. Two doors, one implementation — a copy of
 * any rule here would be a second copy of it.
 *
 * NO EDITOR ADDRESS IS PASSED. createLotFile falls back to its own
 * DEFAULT_EDITOR_EMAIL whenever requestedEmail is absent or invalid, so
 * that address exists in exactly one place in the system. Passing the
 * signed-in worker instead would make that worker an Editor on the lot
 * file they create, which is the hole the menu decommission closed.
 * Do not add an e-mail field to this screen.
 *
 * NO MAIL IS SENT, AND NO oauthScopes BLOCK IS ADDED. A bound library
 * runs under the CALLER's authorisation, and Apps Script's automatic
 * scope detection reads the caller's own source only, never library
 * source. tsaraentry uses DriveApp and SpreadsheetApp but holds no
 * MailApp or GmailApp call anywhere, so a mail call inside the borrowed
 * function would fail this whole screen at run time. createLotFile
 * sends none. Declaring scopes by hand would switch auto-detection off
 * and require naming every scope the app already uses; omit one and a
 * working screen breaks.
 *
 * Accepted, recorded consequence: if the cohort mint fails on THIS
 * path, nobody is mailed. The screen shows mintWarning, the engine
 * auto-mints on the lot's first Stock Poisson appearance, and
 * RUN_findDuplicateRootCohorts surfaces it on demand.
 *
 * NO LOT-NUMBER PREVIEW. A preview costs a full folder scan plus a
 * Stock Poisson scan, and the number can still change before the script
 * lock takes effect, so a preview can only ever be a promise the system
 * may break. The result panel shows the real number.
 ***************************************************************/

/**
 * Create one lot file.
 *
 * Nothing is validated here. createLotFile refuses a bad broodstock
 * name and a duplicate name itself, and returns a French message for
 * each; re-checking here would be a second copy of the rules.
 *
 * @param {boolean} isBroodstock  true = Géniteur, false = Grossissement
 * @param {string}  broodName     family name — read only when Géniteur
 * @return {Object} refused -> { ok:false, reason, message }
 *                  created -> { ok:true, lotId, isBroodstock, newName,
 *                               fileUrl, editorEmail, cohortId, mintWarning }
 *                  System faults (lock timeout, numbering conflict,
 *                  missing template sheet) throw and reach the client's
 *                  failure handler.
 */
function creerCreateLot(isBroodstock, broodName) {
  return CreationLot.createLotFile({
    isBroodstock: !!isBroodstock,
    broodName: broodName
  });
}

/**
 * RUN FROM EDITOR: read-only check that the library binding resolves.
 *
 * Writes nothing and creates nothing. Run this after every clasp push
 * and BEFORE creating a new deployment. A HEAD binding that Apps Script
 * refuses fails here, with no deployment made and no worker exposed.
 */
function testCreerLotServer() {
  Logger.log("CreationLot.createLotFile reachable: " +
    (typeof CreationLot.createLotFile === "function"));
}
