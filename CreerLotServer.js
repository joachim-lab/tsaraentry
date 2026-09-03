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
 * Read-only. What the continuation box should show, or null when there
 * is no harvest waiting to be continued. Thin by design: the carry's
 * rules live in CarryServer.js.
 *
 * @return {Object|null}
 */
function creerCarryStatus() {
  return carryDescribe();
}

/**
 * Create a growout lot and fill it from the harvest carry.
 *
 * SEQUENCING ONLY. Every rule still belongs to createLotFile — the
 * numbering, the duplicate check, the cohort mint — and every cell
 * write still belongs to submitLotEntry, reached through
 * carryApplyToLot. This function does nothing but call the two in
 * order and report what happened.
 *
 * Always a GROWOUT lot. A géniteur family is never a harvest
 * continuation, and the screen hides the box for that type.
 *
 * If the fill fails the file still exists and is returned with
 * carry.ok false. The carry itself is left untouched by
 * carryApplyToLot in that case, so the remainder is not lost.
 *
 * @param {string} bassin  bassin of the NEW lot
 * @param {string} happa   happa of the NEW lot
 * @return {Object} createLotFile result, plus fileId and carry
 */
function creerCreateLotWithCarry(bassin, happa) {
  const res = CreationLot.createLotFile({ isBroodstock: false });
  if (!res.ok) return res;

  // createLotFile returns the URL, not the id. Reading the id out of
  // the URL keeps this change inside one project; changing the library
  // signature would touch the Google Form path too, for one field.
  const m = String(res.fileUrl || "").match(/\/d\/([a-zA-Z0-9_-]{15,})/);
  if (!m) {
    return Object.assign({}, res, {
      carry: {
        ok: false,
        message: "Le fichier a été créé mais son identifiant n'a pas pu être lu. " +
          "Saisissez ce lot à la main dans Échantillonnage."
      }
    });
  }

  let applied;
  try {
    applied = carryApplyToLot(m[1], bassin, happa);
  } catch (e) {
    applied = { ok: false, message: String(e && e.message ? e.message : e) };
  }

  return Object.assign({}, res, { fileId: m[1], carry: applied });
}

/**
 * Drive folder that receives archived lot files.
 *
 * Archiving MOVES the file: DriveApp.moveTo removes every existing
 * parent. Once moved, the lot is outside LOT_CFG.LOTS_FOLDER_ID, so
 * buildLotFileList no longer sees it and neither does any screen or
 * engine pass that enumerates that folder. That is the whole point of
 * the button. Nothing is deleted and nothing is renamed, so the move
 * is undone by dragging the file back in Drive.
 *
 * WHO RUNS THE MOVE: the deployment is executeAs USER_DEPLOYING
 * (appsscript.json), so every server call on this screen runs under
 * Kim's authorisation, not the clicking worker's. No worker needs any
 * right on the archive folder, and a worker who cannot open the lot
 * file at all can still archive it from here. The access check that
 * remains is the one this function performs itself.
 */
const ARCHIVE_FOLDER_ID = "1ei75vejg3_CUtY4QdpQG9EmHHi1ujGYw";

/**
 * Read-only. Every lot file currently in the lots folder, for the
 * archive dropdown.
 *
 * Reuses getLotFileList (Lot.js), the same list screens 1 and 2 read,
 * so the dropdown can never show a lot the rest of the app does not.
 * That list is cached for 5 minutes: a lot created moments ago may not
 * appear yet. Archiving one clears the cache.
 *
 * @return {Array<Object>} [{ lotNumber, fileId, fileName }]
 */
function creerListLots() {
  return getLotFileList();
}

/**
 * Move one lot file out of the lots folder and into the archive folder.
 *
 * The parent check is not decoration. fileId arrives from the client,
 * and moveTo would happily move ANY file this worker can edit. Reading
 * the real parents of the real file — not the cached list — is the one
 * thing that proves the target is a lot file that is still in the lots
 * folder.
 *
 * @param {string} fileId  Drive id of the lot file
 * @return {Object} { ok:false, message } or { ok:true, fileName, fileUrl }
 *                  A permission fault throws and reaches the client's
 *                  failure handler.
 */
function creerArchiveLot(fileId) {
  const id = String(fileId || "").trim();
  if (!id) return { ok: false, message: "Aucun lot sélectionné." };

  const file = DriveApp.getFileById(id);
  const name = file.getName();

  if (!/^Lot-/.test(name)) {
    return { ok: false, message: "Ce fichier n'est pas un fichier lot : " + name };
  }

  let inLots = false;
  const parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === LOT_CFG.LOTS_FOLDER_ID) inLots = true;
  }
  if (!inLots) {
    clearLotFileListCache();
    return {
      ok: false,
      message: "Ce lot n'est plus dans le dossier des lots. La liste a été rafraîchie."
    };
  }

  file.moveTo(DriveApp.getFolderById(ARCHIVE_FOLDER_ID));
  clearLotFileListCache();

  return { ok: true, fileName: name, fileUrl: file.getUrl() };
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
