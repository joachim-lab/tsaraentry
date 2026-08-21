/***************************************************************
 * CarryServer.js — TSARA Entry web app
 *
 * HARVEST CARRY. At harvest the fry are weighed once, all together, then
 * split across several happas. Each happa becomes its own lot. Everything
 * except bassin and happa is identical for every lot of that harvest:
 * the same day, the same water temperature, the same fish, therefore the
 * same millimetrage. Before this file the worker retyped all of it.
 *
 * A carry holds what the next lot needs, plus how much biomasse and how
 * many lots are still to come.
 *
 * ONE CARRY, FARM-WIDE. The web app is deployed "Execute as: Me", so
 * getUserProperties() returns the owner's store no matter which worker is
 * signed in. Pretending the carry is per-worker would be a lie, so it is
 * script-scoped and honest about it. Two harvests split in the same hour
 * would collide — the tick box in Créer un lot always names the source
 * lot and the amount, so a stale carry is visible, never silent.
 *
 * EXPIRY. None by clock. A carry dies when it is consumed, or when the
 * next save with a remainder replaces it. Kim's rule, 2026-08-21.
 ***************************************************************/

const CARRY_KEY = "tt_harvest_carry_v1";

function carryStore_() {
  return PropertiesService.getScriptProperties();
}

/**
 * @return {Object|null} the stored carry, or null when there is none.
 * An unparseable value is deleted rather than thrown — a corrupt carry
 * must never block lot creation.
 */
function carryGet() {
  const raw = carryStore_().getProperty(CARRY_KEY);
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (e) {
    console.log("Carry unreadable, discarded: " + (e && e.message ? e.message : e));
    carryStore_().deleteProperty(CARRY_KEY);
    return null;
  }
}

/** Remove the carry. Always safe. */
function carryClear() {
  carryStore_().deleteProperty(CARRY_KEY);
  return true;
}

/**
 * Store a carry. Called by the Echantillonnage screen after a successful
 * save that left a remainder.
 *
 * @param {Object} c
 * @param {string} c.sourceLot          lot id the remainder came from
 * @param {*}      c.b4                 température, exactly as the screen sent it
 * @param {*}      c.b5                 date de départ, exactly as the screen sent it
 * @param {Array<number>} c.samples     millimetrage of the harvest
 * @param {number} c.remainingBiomasse  grams still to place
 * @param {number} c.lotsRemaining      lots still to create (>= 1)
 * @return {Object} { ok, carry } or { ok:false, message }
 */
function carrySet(c) {
  const p = c || {};
  const samples = Array.isArray(p.samples) ? p.samples.map(Number).filter(v => v > 0) : [];
  const remaining = Number(p.remainingBiomasse);
  const lots = Math.floor(Number(p.lotsRemaining));

  if (!(remaining > 0)) {
    return { ok: false, message: "Biomasse restante absente ou nulle — aucun report enregistré." };
  }
  if (!(lots >= 1)) {
    return { ok: false, message: "Nombre de lots restants invalide — aucun report enregistré." };
  }

  const carry = {
    sourceLot: String(p.sourceLot == null ? "" : p.sourceLot),
    b4: p.b4 === undefined ? null : p.b4,
    b5: p.b5 === undefined ? null : p.b5,
    samples: samples,
    remainingBiomasse: remaining,
    lotsRemaining: lots,
    createdAt: new Date().toISOString()
  };

  carryStore_().setProperty(CARRY_KEY, JSON.stringify(carry));
  console.log("Carry stored from lot " + carry.sourceLot + ": " +
    remaining + " g over " + lots + " lot(s), " + samples.length + " mesure(s).");

  return { ok: true, carry: carry };
}

/**
 * What Créer un lot shows in its continuation box.
 * @return {Object|null} null when there is nothing to continue.
 */
function carryDescribe() {
  const c = carryGet();
  if (!c) return null;

  const share = c.remainingBiomasse / c.lotsRemaining;

  return {
    sourceLot: c.sourceLot,
    remainingBiomasse: c.remainingBiomasse,
    lotsRemaining: c.lotsRemaining,
    nextShare: share,
    sampleCount: c.samples.length,
    createdAt: c.createdAt
  };
}

/**
 * Write the carried data into a freshly created lot file.
 *
 * Writes NOTHING itself. It builds a payload and calls submitLotEntry, so
 * every guard the Echantillonnage screen enforces applies here too. That
 * is why bassin and happa must be supplied: requireFields refuses a save
 * on tab 1-5 without them, and this path does not get an exemption.
 *
 * The carry is updated only AFTER the write succeeds. A failed write
 * leaves the carry untouched, so the worker can retry or enter the lot by
 * hand without losing the remainder.
 *
 * @param {string} fileId  the new lot file
 * @param {string} bassin
 * @param {string} happa
 * @return {Object} { ok, changeCount, share, lotsRemaining } or { ok:false, message }
 */
function carryApplyToLot(fileId, bassin, happa) {
  const c = carryGet();
  if (!c) return { ok: false, message: "Aucun report disponible." };

  if (!String(bassin || "").trim() || !String(happa || "").trim()) {
    return { ok: false, message: "Bassin et Happa sont obligatoires pour la continuation." };
  }

  const stage = getLotStage(fileId);
  if (stage.stage !== "s-tab" || stage.tabName !== "1-5") {
    return {
      ok: false,
      message: "Ce lot n'est pas un lot neuf sur l'onglet 1-5 (état : " +
        stage.stage + (stage.tabName ? " / " + stage.tabName : "") +
        "). Aucune donnée reportée."
    };
  }

  const share = c.remainingBiomasse / c.lotsRemaining;

  const payload = {
    fields: {
      B2: bassin,
      B3: happa,
      B4: c.b4,
      B5: c.b5,
      D11: share,
      samples: c.samples.slice()
    }
  };

  const res = submitLotEntry(fileId, stage, payload, { dryRun: false });

  const left = c.remainingBiomasse - share;
  const lotsLeft = c.lotsRemaining - 1;

  if (lotsLeft >= 1 && left > 0) {
    c.remainingBiomasse = left;
    c.lotsRemaining = lotsLeft;
    carryStore_().setProperty(CARRY_KEY, JSON.stringify(c));
  } else {
    carryClear();
  }

  console.log("Carry applied to " + fileId + ": " + res.changeCount +
    " cell(s), share " + share + " g, " + lotsLeft + " lot(s) left.");

  return {
    ok: true,
    changeCount: res.changeCount,
    share: share,
    lotsRemaining: lotsLeft
  };
}

/***************************************************************
 * EDITOR TESTS — read-only unless the name says LIVE.
 ***************************************************************/

/** RUN FROM EDITOR: show the current carry. Writes nothing. */
function testCarryShow() {
  const c = carryGet();
  Logger.log(c ? JSON.stringify(c, null, 2) : "No carry stored.");
  Logger.log("Describe: " + JSON.stringify(carryDescribe(), null, 2));
}

/** RUN FROM EDITOR: store a fake carry so the screens can be tested. */
function testCarrySetFake() {
  const res = carrySet({
    sourceLot: "32",
    b4: 26,
    b5: "2026-08-21",
    samples: [11, 12, 12, 13, 11],
    remainingBiomasse: 27000,
    lotsRemaining: 1
  });
  Logger.log(JSON.stringify(res, null, 2));
}

/** RUN FROM EDITOR: remove whatever carry is stored. */
function testCarryClear() {
  carryClear();
  Logger.log("Carry cleared.");
}
