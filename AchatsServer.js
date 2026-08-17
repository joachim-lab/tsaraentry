/***************************************************************
 * AchatsServer.js — TSARA Entry web app
 * Screen 8c: Achats provende (achats du mois sur facture)
 *
 * Writes into the "Nourrissage & inventaire" file, sheet
 * "achats provende". That sheet is a grid of month blocks, not a
 * log. Every block is exactly 8 columns:
 *
 *   slot 1..6 = one invoice date each (row 3 = the date,
 *               rows 4-17 = bags bought per feed type)
 *   slot 7    = "Total <mois>"   =sum(<slot1>4:<slot6>4)
 *   slot 8    = "Valeur <mois>"  =<total>4*$C4
 *
 *   row 1 of the first column of a block = the month, 1st day
 *   row 2 = the block label
 *   row 18 = the monthly totals
 *
 * The script writes ONLY the six invoice slots. The Total and
 * Valeur columns and row 18 keep the sheet's own formulas and are
 * never written as values — checkEcartsInventaireVsAchats() in the
 * controleconsoprovende project reads the Total column, so it must
 * stay a live formula.
 *
 * A missing month block is created by copying the last existing
 * block one block to the right. The copy carries the sheet's own
 * formulas and its green input formatting, so nothing is rebuilt
 * from a hardcoded template. Only the month date, the Total and
 * Valeur labels, and the cleared quantities are then set.
 ***************************************************************/

const AP_CFG = {
  SS_ID: "1JBoH5c7BqZc2V5czcDAnEt-2hvkNKlJAxuDSupttTfs",
  SHEET: "achats provende",
  MONTH_ROW: 1,          // month, 1st day, at the first column of a block
  HEADER_ROW: 3,         // invoice dates, then "Total <mois>" / "Valeur <mois>"
  FIRST_DATA_ROW: 4,     // feed types
  LAST_DATA_ROW: 17,
  TOTAL_ROW: 18,         // monthly totals — formulas, never written
  NAME_COL: 1,           // A = Article
  FIRST_BLOCK_COL: 4,    // D = first month block
  BLOCK_WIDTH: 8,
  SLOTS: 6               // invoice-date columns per month
};

const AP_MONTHS = [
  "janvier", "février", "mars", "avril", "mai", "juin",
  "juillet", "août", "septembre", "octobre", "novembre", "décembre"
];

/** Open the achats sheet, or fail loudly. */
function openAchatsSheet() {
  const sh = SpreadsheetApp.openById(AP_CFG.SS_ID).getSheetByName(AP_CFG.SHEET);
  if (!sh) throw new Error('Onglet introuvable: "' + AP_CFG.SHEET + '"');
  return sh;
}

/** "2026-08" from a Date. */
function achatsMonthKey(d) {
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
}

/** "août 2026" from a Date. */
function achatsMonthLabel(d) {
  return AP_MONTHS[d.getMonth()] + " " + d.getFullYear();
}

/**
 * Every month block present in the sheet, left to right.
 * Returns [{ col, date, key, label }].
 */
function findAchatsBlocks(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < AP_CFG.FIRST_BLOCK_COL) return [];

  const row1 = sh.getRange(AP_CFG.MONTH_ROW, 1, 1, lastCol).getValues()[0];
  const out = [];
  for (let c = AP_CFG.FIRST_BLOCK_COL; c <= lastCol; c++) {
    const v = row1[c - 1];
    if (Object.prototype.toString.call(v) !== "[object Date]") continue;
    out.push({ col: c, date: v, key: achatsMonthKey(v), label: achatsMonthLabel(v) });
  }
  return out;
}

/** The six invoice-date headers of one block. Returns [{ slot, date }]. */
function readAchatsSlots(sh, startCol, tz) {
  const headers = sh.getRange(AP_CFG.HEADER_ROW, startCol, 1, AP_CFG.SLOTS).getValues()[0];
  const out = [];
  for (let i = 0; i < AP_CFG.SLOTS; i++) {
    const v = headers[i];
    const isDate = Object.prototype.toString.call(v) === "[object Date]";
    out.push({
      slot: i,
      used: v !== "" && v !== null,
      date: isDate ? Utilities.formatDate(v, tz, "dd/MM/yyyy") : (v === "" || v === null ? "" : String(v)),
      time: isDate ? v.getTime() : null
    });
  }
  return out;
}

/**
 * Everything the screen needs: the feed-type list from column A, the
 * month blocks that exist with their free slots, and the one month
 * that may still be created.
 *
 * Only the month directly after the last block is offered. Blocks
 * must stay contiguous and in date order, so allowing a gap would
 * break the layout the ecarts check reads.
 */
function getAchatsContext() {
  const sh = openAchatsSheet();
  const tz = sh.getParent().getSpreadsheetTimeZone() || Session.getScriptTimeZone();

  const n = AP_CFG.LAST_DATA_ROW - AP_CFG.FIRST_DATA_ROW + 1;
  const names = sh.getRange(AP_CFG.FIRST_DATA_ROW, AP_CFG.NAME_COL, n, 1).getValues();
  const feeds = [];
  for (let i = 0; i < n; i++) {
    const nm = String(names[i][0] || "").trim();
    if (nm) feeds.push(nm);
  }

  const blocks = findAchatsBlocks(sh);
  const months = blocks.map(b => {
    const slots = readAchatsSlots(sh, b.col, tz);
    return {
      key: b.key,
      label: b.label,
      exists: true,
      freeSlots: slots.filter(s => !s.used).length,
      usedDates: slots.filter(s => s.used).map(s => s.date)
    };
  });

  if (blocks.length) {
    const last = blocks[blocks.length - 1].date;
    const next = new Date(last.getFullYear(), last.getMonth() + 1, 1);
    months.push({
      key: achatsMonthKey(next),
      label: achatsMonthLabel(next),
      exists: false,
      freeSlots: AP_CFG.SLOTS,
      usedDates: []
    });
  }

  const today = new Date();
  return {
    feeds: feeds,
    months: months,
    currentMonthKey: achatsMonthKey(today),
    slots: AP_CFG.SLOTS
  };
}

/**
 * Create the month block that follows the last existing one.
 * The whole 18-row block is copied from the last block, so the
 * Total and Valeur formulas, row 18, and the green input formatting
 * all come from the sheet itself. Returns the new start column.
 */
function createAchatsMonthBlock(sh, monthDate) {
  const blocks = findAchatsBlocks(sh);
  if (!blocks.length) throw new Error("Aucun bloc de mois existant : impossible d'en créer un nouveau.");

  const last = blocks[blocks.length - 1];
  const expected = new Date(last.date.getFullYear(), last.date.getMonth() + 1, 1);
  if (achatsMonthKey(monthDate) !== achatsMonthKey(expected)) {
    throw new Error(
      "Seul le mois qui suit " + achatsMonthLabel(last.date) + " peut être créé (" +
      achatsMonthLabel(expected) + "). Les blocs doivent rester contigus."
    );
  }

  const srcStart = last.col;
  const dstStart = srcStart + AP_CFG.BLOCK_WIDTH;
  const needed = dstStart + AP_CFG.BLOCK_WIDTH - 1;
  const maxCols = sh.getMaxColumns();
  if (needed > maxCols) sh.insertColumnsAfter(maxCols, needed - maxCols);

  sh.getRange(1, srcStart, AP_CFG.TOTAL_ROW, AP_CFG.BLOCK_WIDTH)
    .copyTo(sh.getRange(1, dstStart, AP_CFG.TOTAL_ROW, AP_CFG.BLOCK_WIDTH));

  sh.getRange(AP_CFG.MONTH_ROW, dstStart).setValue(monthDate);
  sh.getRange(AP_CFG.HEADER_ROW, dstStart, 1, AP_CFG.SLOTS).clearContent();

  const label = AP_MONTHS[monthDate.getMonth()];
  sh.getRange(AP_CFG.HEADER_ROW, dstStart + AP_CFG.SLOTS, 1, 2)
    .setValues([["Total " + label, "Valeur " + label]]);

  const rows = AP_CFG.LAST_DATA_ROW - AP_CFG.FIRST_DATA_ROW + 1;
  sh.getRange(AP_CFG.FIRST_DATA_ROW, dstStart, rows, AP_CFG.SLOTS).clearContent();

  return dstStart;
}

/**
 * Record one invoice: its date goes in the first free slot of the
 * month, the bags per feed type go in rows 4-17 of that slot.
 *
 * payload: { monthKey: "2026-08", date: "yyyy-MM-dd",
 *            counts: [{ name, value }] }
 *
 * Quantities are matched to rows BY NAME, never by position.
 * An empty value writes an empty cell; the Total formula reads it
 * as zero, so a blank and a zero mean the same thing here.
 *
 * Returns { month, date, filled, created }.
 */
function submitAchats(payload) {
  if (!payload) throw new Error("Aucune donnée reçue.");

  const monthKey = String(payload.monthKey || "").trim();
  if (!monthKey) throw new Error("Mois obligatoire.");

  const dateStr = String(payload.date || "").trim();
  if (!dateStr) throw new Error("Date de la facture obligatoire.");

  const counts = payload.counts || [];
  if (!counts.length) throw new Error("Aucune quantité saisie.");

  const sh = openAchatsSheet();
  const tz = sh.getParent().getSpreadsheetTimeZone() || Session.getScriptTimeZone();

  const invoiceDate = new Date(dateStr);
  if (isNaN(invoiceDate.getTime())) throw new Error("Date invalide (reçu: " + dateStr + ").");
  if (achatsMonthKey(invoiceDate) !== monthKey) {
    throw new Error("La date de la facture n'est pas dans le mois choisi.");
  }

  let created = false;
  let block = findAchatsBlocks(sh).filter(b => b.key === monthKey)[0];
  if (!block) {
    const parts = monthKey.split("-");
    const monthDate = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
    const col = createAchatsMonthBlock(sh, monthDate);
    created = true;
    block = { col: col, date: monthDate, key: monthKey, label: achatsMonthLabel(monthDate) };
  }

  const slots = readAchatsSlots(sh, block.col, tz);

  const sameDay = slots.filter(s => s.time !== null &&
    Utilities.formatDate(new Date(s.time), tz, "yyyy-MM-dd") === dateStr)[0];
  if (sameDay) {
    throw new Error(
      "Une facture du " + sameDay.date + " est déjà enregistrée pour " + block.label +
      ". Corrigez-la directement dans la feuille."
    );
  }

  const free = slots.filter(s => !s.used)[0];
  if (!free) {
    throw new Error(
      "Les " + AP_CFG.SLOTS + " colonnes de " + block.label + " sont utilisées. " +
      "Ajoutez la facture directement dans la feuille."
    );
  }

  const nRows = AP_CFG.LAST_DATA_ROW - AP_CFG.FIRST_DATA_ROW + 1;
  const names = sh.getRange(AP_CFG.FIRST_DATA_ROW, AP_CFG.NAME_COL, nRows, 1).getValues();
  const rowIndexByName = {};
  for (let i = 0; i < nRows; i++) {
    const nm = String(names[i][0] || "").trim();
    if (nm) rowIndexByName[nm] = i;
  }

  const values = [];
  for (let i = 0; i < nRows; i++) values.push([""]);

  let filled = 0;
  for (let k = 0; k < counts.length; k++) {
    const nm = String(counts[k].name || "").trim();
    if (!nm) continue;
    if (!(nm in rowIndexByName)) {
      throw new Error('Type de provende absent de l\'onglet: "' + nm + '". Rechargez l\'écran.');
    }
    const raw = counts[k].value;
    if (raw === "" || raw === null || raw === undefined) continue;
    const v = Number(raw);
    if (!isFinite(v) || v < 0) {
      throw new Error("Quantité invalide pour " + nm + " (reçu: " + raw + ").");
    }
    values[rowIndexByName[nm]] = [v];
    filled++;
  }

  if (!filled) throw new Error("Aucune quantité saisie.");

  const col = block.col + free.slot;
  sh.getRange(AP_CFG.HEADER_ROW, col).setValue(invoiceDate);
  sh.getRange(AP_CFG.FIRST_DATA_ROW, col, nRows, 1).setValues(values);

  return {
    month: block.label,
    date: Utilities.formatDate(invoiceDate, tz, "dd/MM/yyyy"),
    filled: filled,
    created: created
  };
}
