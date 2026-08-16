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
 * OPERATOR IDENTITY IS DETECTED, NOT CHOSEN. Session.getActiveUser()
 * returns the SIGNED-IN WORKER's address here, because the workers hold
 * accounts inside the jdsresearch.com domain and the app is deployed to
 * that domain. An earlier comment in this file claimed the call returns
 * the deployer for everyone; that was wrong, and the operator dropdown
 * it justified has been removed. The detected address travels in
 * payload.operator and lands in column G, keeping G homogeneous with
 * every row the lot-file menu ever wrote.
 *
 * Reuses, unchanged, from elsewhere in this project:
 *   getLotFileList()   (Lot.js) — {lotNumber, fileId, fileName}
 ***************************************************************/

/**
 * Everything the picker needs, in one round trip.
 *
 * Operations come from LigneeUI.LU_CFG.OPS rather than a list written
 * here. LEGACY_OPS is excluded on purpose: those labels are accepted on
 * input so old rows stay valid, but must never be produced again.
 *
 * @return {Object} { lots, operator, ops }
 */
function tracGetOptions() {
  const lots = getLotFileList().map(function (l) {
    return { lotNumber: l.lotNumber, fileId: l.fileId };
  });
  const ops = LigneeUI.LU_CFG.OPS.map(function (o) {
    return { key: o.key, label: o.label };
  });
  return { lots: lots, operator: tracCurrentOperatorEmail(), ops: ops };
}

/**
 * The signed-in worker's e-mail address.
 *
 * Throws rather than defaulting. An unattributable event is worse than
 * a refused one, and an empty string in column G would look like a
 * saved row while being untraceable. The throw surfaces on the picker's
 * first round trip, so a wrong or missing sign-in is seen before a lot
 * is chosen, not after the form is filled in.
 *
 * @return {string}
 */
function tracCurrentOperatorEmail() {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error(
      "Impossible d'identifier l'opérateur. Vérifiez que vous êtes connecté " +
      "avec votre compte @jdsresearch.com, puis rechargez la page.");
  }
  return email;
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
  Logger.log("Opérateur détecté: " + o.operator);
  Logger.log("CohortRegistry reachable: " + (typeof CohortRegistry.lp_serialise === "function"));
}
