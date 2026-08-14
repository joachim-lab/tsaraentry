/***************************************************************
 * TracabiliteServer.js — TSARA Entry web app
 * Screen 5: Traçabilité — server side.
 *
 * DELIBERATELY THIN. The form, every validation rule and the actual
 * Lignée write all live in LigneeUI, bound at HEAD. This file only
 * supplies the picker's options and proxies the save. It re-checks
 * NOTHING: LigneeUI already refuses more fish than Stock Poisson holds,
 * rejects a cage parent on a tri, caps children per surface, and builds
 * the column-I payload through CohortRegistry.lp_serialise, which
 * throws before any row is written. A second copy of those rules here
 * would be a third copy overall — the main pitfall named in the port
 * scope (§6.4).
 *
 * COLUMN OWNERSHIP is unchanged and must stay that way:
 *   A–G + I  LigneeUI      H  registry consumer      J  LigneeWriter
 * This screen writes nothing directly. It cannot touch H or J.
 *
 * WHY AN OPERATOR LIST EXISTS HERE. The lot-file menu records
 * Session.getActiveUser(), which is the real worker. tsaraentry runs as
 * USER_DEPLOYING, so that call returns the DEPLOYER for everyone. The
 * chosen operator's e-mail travels in payload.operator and lands in
 * column G, keeping G homogeneous with every row the menu ever wrote.
 *
 * Reuses, unchanged, from elsewhere in this project:
 *   getLotFileList()   (Lot.js) — {lotNumber, fileId, fileName}
 ***************************************************************/

const TRAC_CFG = {
  /* Farm workers who may log a traceability event. The dropdown shows
   * name; column G receives email, matching what the menu path writes.
   * Adding a worker is one line here — no sheet, no lookup, no cache. */
  OPERATORS: [
    { name: "Hasina",  email: "fenoharijaonatsinjohasina@gmail.com" },
    { name: "Audry",   email: "audryjoach21@gmail.com" },
    { name: "Charles", email: "charlesmascar48@gmail.com" }
  ]
};

/**
 * Everything the picker needs, in one round trip.
 *
 * Operations come from LigneeUI.LU_CFG.OPS rather than a list written
 * here. LEGACY_OPS is excluded on purpose: those labels are accepted on
 * input so old rows stay valid, but must never be produced again.
 *
 * @return {Object} { lots, operators, ops }
 */
function tracGetOptions() {
  const lots = getLotFileList().map(function (l) {
    return { lotNumber: l.lotNumber, fileId: l.fileId };
  });
  const ops = LigneeUI.LU_CFG.OPS.map(function (o) {
    return { key: o.key, label: o.label };
  });
  const operators = TRAC_CFG.OPERATORS.map(function (o, i) {
    return { idx: i, name: o.name };
  });
  return { lots: lots, operators: operators, ops: ops };
}

/**
 * E-mail for an operator index, as sent by the picker.
 * Throws rather than defaulting: an unattributable event is worse than
 * a refused one, and a silent fallback would write the deployer's
 * address while looking correct.
 *
 * @param {string|number} idx
 * @return {string}
 */
function tracOperatorEmail(idx) {
  const i = Number(idx);
  const o = TRAC_CFG.OPERATORS[i];
  if (!o) throw new Error("Opérateur inconnu (index: " + idx + ").");
  return o.email;
}

/**
 * Proxy for the form's save call.
 *
 * The form is LigneeUI's own HTML, served as the whole page. Its
 * google.script.run reaches THIS project, so a function of this exact
 * name must exist here — the same shim pattern each lot file uses in
 * its bound Lignee.gs. Nothing is added or checked on the way through.
 *
 * @param {Object} payload  built by the form; carries operator
 * @param {string} ssId     lot-file id baked into the form
 * @return {string} "OK"
 */
function lu_submitEvent(payload, ssId) {
  return LigneeUI.lu_submitEvent(payload, ssId);
}

/** RUN FROM EDITOR: read-only check that the bindings resolve. */
function testTracabiliteServer() {
  const o = tracGetOptions();
  Logger.log("Lots (" + o.lots.length + "): " +
    o.lots.slice(0, 5).map(function (l) { return l.lotNumber; }).join(" | "));
  Logger.log("Opérations: " +
    o.ops.map(function (x) { return x.key + "=" + x.label; }).join(" | "));
  Logger.log("Opérateurs: " +
    o.operators.map(function (x) { return x.idx + "=" + x.name; }).join(" | "));
  Logger.log("CohortRegistry reachable: " + (typeof CohortRegistry.lp_serialise === "function"));
}
