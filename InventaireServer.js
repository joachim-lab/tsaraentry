/***************************************************************
 * InventaireServer.js — TSARA Entry web app
 * Screen 8a: Inventaire (commandes et receptions)
 *
 * Appends rows to the "Nourrissage & inventaire" file, sheet
 * "Inventaire":
 *   A Type · B Date Commande · C Unite · D Nbre commande
 *   E Date Reception · F Nbre recu · H Commentaire
 *
 * Column G ("Difference entre commande et recu") is a sheet
 * formula (=D-F). It is NEVER written as a value. It is copied
 * from the row above with PASTE_FORMULA, so the sheet's own
 * formula stays the single source of truth — the same method
 * Code.js uses for column I of "Consommation provende".
 *
 * Both dropdown lists are read live from the sheet's own data
 * validation rules, never hardcoded, so they cannot drift from
 * the sheet.
 ***************************************************************/

const INV_CFG = {
  SS_ID: "1JBoH5c7BqZc2V5czcDAnEt-2hvkNKlJAxuDSupttTfs",
  SHEET: "Inventaire",
  START_ROW: 2,
  TYPE_COL: 1,       // A
  UNITE_COL: 3,      // C
  BLOCK_COLS: 6,     // A:F written as one block
  DATE_RECEPTION_COL: 5, // E
  RECU_COL: 6,       // F
  DIFF_COL: 7,       // G — formula, copied not written
  COMMENT_COL: 8     // H
};

/** Open the Inventaire sheet, or fail loudly. */
function openInventaireSheet() {
  const sh = SpreadsheetApp.openById(INV_CFG.SS_ID).getSheetByName(INV_CFG.SHEET);
  if (!sh) throw new Error('Onglet introuvable: "' + INV_CFG.SHEET + '"');
  return sh;
}

/** Read a dropdown list live from the data validation rule on one cell. */
function readInventaireListFromColumn(col) {
  const sh = openInventaireSheet();
  const rule = sh.getRange(INV_CFG.START_ROW, col).getDataValidation();
  if (!rule) return [];
  const criteria = rule.getCriteriaValues(); // [ [list values], ... ]
  return criteria[0] || [];
}

/** Type list — from the dropdown on Inventaire column A. */
function getInventaireTypes() {
  return readInventaireListFromColumn(INV_CFG.TYPE_COL);
}

/** Unite list — from the dropdown on Inventaire column C. */
function getInventaireUnites() {
  return readInventaireListFromColumn(INV_CFG.UNITE_COL);
}

/**
 * First row after the last real entry, scanning column A only.
 * getLastRow() is unusable here: column G holds pre-filled =D-F
 * formulas far below the data, so it reports row 996, not the
 * last entry.
 */
function findNextInventaireRow(sh) {
  const lastPhysical = sh.getLastRow();
  const n = lastPhysical - INV_CFG.START_ROW + 1;
  if (n < 1) return INV_CFG.START_ROW;

  const aVals = sh.getRange(INV_CFG.START_ROW, INV_CFG.TYPE_COL, n, 1).getValues();

  let lastData = INV_CFG.START_ROW - 1;
  for (let i = 0; i < n; i++) {
    const a = aVals[i][0];
    if (a !== "" && a !== null) lastData = INV_CFG.START_ROW + i;
  }
  return lastData + 1;
}

/**
 * Append one or more inventory entries.
 *
 * entries: [{ type, dateCommande, unite, nbreCommande,
 *             dateReception, nbreRecu, commentaire }]
 * Dates arrive as "yyyy-MM-dd" from the browser. Reception fields
 * may be empty — an order that is not yet delivered is a valid row.
 * A received quantity without a reception date is refused: the
 * ecarts check reads column E to decide the month, so such a row
 * would never be counted.
 *
 * Returns { written, startRow, endRow }.
 */
function submitInventaire(entries) {
  if (!entries || !entries.length) throw new Error("Aucune entrée à enregistrer.");

  const rows = entries.map(en => {
    const type = String(en.type || "").trim();
    const unite = String(en.unite || "").trim();
    const dateCommande = String(en.dateCommande || "").trim();
    const nbreCommande = Number(en.nbreCommande);

    if (!type || !dateCommande || !unite) {
      throw new Error("Entrée incomplète : type, date commande et unité sont obligatoires.");
    }
    if (!isFinite(nbreCommande) || nbreCommande <= 0) {
      throw new Error("Nbre commandé invalide (reçu: " + en.nbreCommande + ").");
    }

    const dateReception = String(en.dateReception || "").trim();
    const hasRecu = en.nbreRecu !== "" && en.nbreRecu !== null && en.nbreRecu !== undefined;
    let nbreRecu = "";
    if (hasRecu) {
      nbreRecu = Number(en.nbreRecu);
      if (!isFinite(nbreRecu) || nbreRecu < 0) {
        throw new Error("Nbre reçu invalide (reçu: " + en.nbreRecu + ").");
      }
      if (nbreRecu > 0 && !dateReception) {
        throw new Error("Date réception obligatoire quand le nbre reçu est supérieur à zéro.");
      }
    }

    return [
      type,
      new Date(dateCommande),
      unite,
      nbreCommande,
      dateReception ? new Date(dateReception) : "",
      nbreRecu
    ];
  });

  const comments = entries.map(en => [String(en.commentaire || "").trim()]);

  const sh = openInventaireSheet();
  const startRow = findNextInventaireRow(sh);
  const endRow = startRow + rows.length - 1;

  sh.getRange(startRow, INV_CFG.TYPE_COL, rows.length, INV_CFG.BLOCK_COLS).setValues(rows); // A:F
  sh.getRange(startRow, INV_CFG.COMMENT_COL, rows.length, 1).setValues(comments);           // H

  // Column G: copy the sheet's own =D-F formula down from the row above.
  if (startRow > INV_CFG.START_ROW) {
    const src = sh.getRange(startRow - 1, INV_CFG.DIFF_COL);
    const dst = sh.getRange(startRow, INV_CFG.DIFF_COL, rows.length, 1);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  }

  return { written: rows.length, startRow: startRow, endRow: endRow };
}

/** Display form of a cell that may hold a Date or plain text. */
function formatInventaireDate(value, tz) {
  if (value === "" || value === null || value === undefined) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, tz, "dd/MM/yyyy");
  }
  return String(value);
}

/**
 * Orders that are still waiting for their delivery: column A filled,
 * columns E (Date Reception) AND F (Nbre recu) both empty. A row with
 * "Nbre recu = 0" is settled, not pending — that is how a rupture de
 * stock is recorded.
 *
 * Returns [{ row, type, dateCommande, unite, nbreCommande, commentaire }].
 */
function getPendingInventaireRows() {
  const sh = openInventaireSheet();
  const tz = sh.getParent().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const lastDataRow = findNextInventaireRow(sh) - 1;
  if (lastDataRow < INV_CFG.START_ROW) return [];

  const n = lastDataRow - INV_CFG.START_ROW + 1;
  const vals = sh.getRange(INV_CFG.START_ROW, 1, n, INV_CFG.COMMENT_COL).getValues();

  const out = [];
  for (let i = 0; i < n; i++) {
    const v = vals[i];
    const type = v[0];
    const dateReception = v[4];
    const nbreRecu = v[5];
    if (type === "" || type === null) continue;
    const hasDate = dateReception !== "" && dateReception !== null;
    const hasRecu = nbreRecu !== "" && nbreRecu !== null;
    if (hasDate || hasRecu) continue;

    out.push({
      row: INV_CFG.START_ROW + i,
      type: String(type),
      dateCommande: formatInventaireDate(v[1], tz),
      unite: String(v[2] === null ? "" : v[2]),
      nbreCommande: v[3] === "" || v[3] === null ? "" : Number(v[3]),
      commentaire: String(v[7] === null ? "" : v[7])
    });
  }
  return out;
}

/**
 * Fill the delivery of one pending row: column E (Date Reception),
 * column F (Nbre recu) and, only when the user typed one, column H
 * (Commentaire). Columns A-D are never touched, and column G keeps
 * its own =D-F formula.
 *
 * The row is re-read and re-checked immediately before the write. A
 * list held open in a browser can be minutes old, and the row may
 * already be settled by then.
 *
 * item: { row, dateReception, nbreRecu, commentaire }
 * Returns { row, type }.
 */
function completeInventaireReception(item) {
  if (!item || !item.row) throw new Error("Ligne manquante.");
  const row = Number(item.row);
  if (!isFinite(row) || row < INV_CFG.START_ROW) {
    throw new Error("Numéro de ligne invalide (reçu: " + item.row + ").");
  }

  const dateReception = String(item.dateReception || "").trim();
  if (!dateReception) throw new Error("Date réception obligatoire.");

  const nbreRecu = Number(item.nbreRecu);
  if (!isFinite(nbreRecu) || nbreRecu < 0) {
    throw new Error("Nbre reçu invalide (reçu: " + item.nbreRecu + ").");
  }

  const sh = openInventaireSheet();
  const cur = sh.getRange(row, 1, 1, INV_CFG.COMMENT_COL).getValues()[0];
  const type = cur[0];
  const curDateReception = cur[4];
  const curRecu = cur[5];

  if (type === "" || type === null) {
    throw new Error("La ligne " + row + " est vide. La liste est périmée : rechargez l'écran.");
  }
  if ((curDateReception !== "" && curDateReception !== null) ||
      (curRecu !== "" && curRecu !== null)) {
    throw new Error("La ligne " + row + " a déjà une réception. Rechargez l'écran.");
  }

  sh.getRange(row, INV_CFG.DATE_RECEPTION_COL, 1, 2)
    .setValues([[new Date(dateReception), nbreRecu]]); // E:F

  const commentaire = String(item.commentaire || "").trim();
  if (commentaire) sh.getRange(row, INV_CFG.COMMENT_COL).setValue(commentaire);

  return {
    type: String(type),
    unite: String(cur[2] === null ? "" : cur[2]),
    nbreRecu: nbreRecu
  };
}
