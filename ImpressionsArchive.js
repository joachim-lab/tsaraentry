/***************************************************************
 * ImpressionsArchive.js — TSARA Entry web app, screen 7 (Impressions)
 *
 * WHAT THIS ADDS. Until now the three Impressions buttons built a PDF,
 * sent the bytes to the browser and kept nothing. This file saves a
 * copy of every printout in its own Drive folder, and keeps the 30 most
 * recent in each folder.
 *
 * WHERE THE WRAPPERS SIT, AND WHY. The builders stay untouched in
 * TemperaturesServer.js and RecetteMtServer.js. The three functions
 * here call a builder, then save. Two consequences, both wanted:
 *   - testTemperaturesServer and testRecetteMtServer keep calling the
 *     raw builders, so an editor test never lands in the archive;
 *   - a change to a PDF layout stays a one-file change.
 *
 * FOLDER IDS ARE HARD-CODED, NOT SCRIPT PROPERTIES. A missing property
 * makes DriveApp.createFolder write a new folder into My Drive root —
 * the FACT_FOLDER_ID trap. A hard-coded id cannot do that: a deleted
 * folder gives a clear refusal instead of a second folder.
 *
 * SAVING NEVER BREAKS PRINTING. The save runs after the bytes exist,
 * inside a try/catch. A Drive fault returns the PDF anyway and reports
 * itself in result.archiveError, which the screen shows under the
 * button. Nothing fails in silence.
 *
 * THE ARCHIVED NAME IS THE NAME THE BROWSER GETS. Two prints in the
 * same minute make two files with the same name. Both are real
 * printouts; pruning is by creation date, so duplicates are harmless.
 ***************************************************************/

const IMPR_CFG = {
  KEEP: 30,         // printouts kept per folder — older ones go to the bin
  LOCK_MS: 10000,   // save and prune are serialised
  FOLDERS: {
    STOCK:     { id: "1zQD-i_VOR-koU0QdWGU5uJ_gwjLvItp9", name: "Impressions Stock poisson" },
    PROGRAMME: { id: "10Gamf754_CYjrtl1UhqxA2rkRFkkOUI0", name: "Impressions Programme" },
    RECETTE:   { id: "1ESDrR05hOKxMJkizXf5L26HkiDlWB8gO", name: "Impressions Recette MT" }
  }
};

/** The archive folder of one kind. Throws a French message when it is gone. */
function imprFolder(kind) {
  const cfg = IMPR_CFG.FOLDERS[kind];
  if (!cfg) throw new Error("Type d'impression inconnu : " + kind);

  var folder;
  try {
    folder = DriveApp.getFolderById(cfg.id);
  } catch (e) {
    throw new Error("Dossier « " + cfg.name + " » introuvable (" + cfg.id +
      "). Kim : tsaraentry -> ImpressionsArchive.js -> IMPR_CFG.");
  }
  if (folder.isTrashed()) {
    throw new Error("Dossier « " + cfg.name + " » est à la corbeille (" + cfg.id + ").");
  }
  return folder;
}

/**
 * Keeps the IMPR_CFG.KEEP newest files of one folder, by creation date.
 * The rest go to the bin — DriveApp cannot delete for good.
 *
 * @return {number} how many went to the bin
 */
function imprPrune(folder) {
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    files.push({ file: f, when: f.getDateCreated().getTime() });
  }
  if (files.length <= IMPR_CFG.KEEP) return 0;

  files.sort(function (a, b) { return b.when - a.when; });

  var trashed = 0;
  for (var i = IMPR_CFG.KEEP; i < files.length; i++) {
    files[i].file.setTrashed(true);
    trashed++;
  }
  return trashed;
}

/**
 * Saves one built PDF in its folder and prunes that folder.
 *
 * @param {string} kind  a key of IMPR_CFG.FOLDERS
 * @param {Object} pdf   { base64, filename, size } from tempFetchPdf
 * @return {string} the url of the saved file
 */
function imprArchive(kind, pdf) {
  const blob = Utilities.newBlob(
    Utilities.base64Decode(pdf.base64), MimeType.PDF, pdf.filename);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(IMPR_CFG.LOCK_MS)) {
    throw new Error("Une autre impression est en cours : copie non enregistrée.");
  }
  try {
    const folder = imprFolder(kind);
    const file = folder.createFile(blob);
    imprPrune(folder);
    return file.getUrl();
  } finally {
    lock.releaseLock();
  }
}

/** Save, then hand the PDF back whatever happened to the save. */
function imprWithArchive(kind, pdf) {
  try {
    pdf.archiveUrl = imprArchive(kind, pdf);
  } catch (e) {
    pdf.archiveError = e.message;
    console.error("Archivage " + kind + " : " + e.message);
  }
  return pdf;
}

/* ---------- the three calls of the Impressions screen ---------- */

/** Stock poisson. The temperature gate of tempPrintStockPdf still applies. */
function imprPrintStock() {
  return imprWithArchive("STOCK", tempPrintStockPdf());
}

/** Programme, sheet "planning". */
function imprPrintProgramme() {
  return imprWithArchive("PROGRAMME", tempBuildProgrammePdf());
}

/** Recette provende MT of one lot. */
function imprPrintRecette(fileId) {
  return imprWithArchive("RECETTE", recetteBuildPdf(fileId));
}

/**
 * RUN FROM EDITOR: tsaraentry -> ImpressionsArchive.js -> testImpressionsArchive
 * Read-only. Proves the three folders resolve and counts what is in them.
 */
function testImpressionsArchive() {
  Object.keys(IMPR_CFG.FOLDERS).forEach(function (k) {
    const folder = imprFolder(k);
    var n = 0;
    const it = folder.getFiles();
    while (it.hasNext()) { it.next(); n++; }
    Logger.log(k + " · " + folder.getName() + " · " + n + " fichier(s) · " + folder.getUrl());
  });
}
