/**
 * StatusEngine.js — status transitions and the registry that guards them.
 *
 * The graph itself is data: Status_Master holds every status of every module and
 * the transitions each one allows, so a new status becomes available by filling
 * in Allowed_Next rather than by changing code (SPEC §5.3, §10.3).
 *
 * A module attaches its own conditions without touching anyone else's:
 *   StatusEngine.registerRule('OPEN', 'CLOSED', Rules.check);
 *   StatusEngine.registerRecheck(Rules.recheck);
 */
var StatusEngine = (function () {

  var rules = [];                         // { from, to, fn }
  var rechecks = [];                      // fn(caseRecord) -> { revertTo, message } | null
  var installed = false;

  /**
   * Fills the registry the first time anything asks for it.
   *
   * Registration happens here, at run time, rather than while the module files
   * are being loaded, because Apps Script picks its own file evaluation order
   * (see Bootstrap.js). The flag is set before install() runs so a rule that
   * registers another rule cannot recurse.
   */
  function ensureInstalled() {
    if (installed) return;
    installed = true;
    if (typeof Bootstrap !== 'undefined') Bootstrap.install();
  }

  /**
   * fn receives { user, caseRecord, from, to, reason } and throws to block.
   * `from` or `to` may be '*' to match any status.
   */
  function registerRule(from, to, fn) {
    rules.push({ from: from, to: to, fn: fn });
  }

  /**
   * A recheck answers one question: given the Case as it is now, must it fall
   * back to an earlier status? Returning null means "nothing to do".
   */
  function registerRecheck(fn) {
    rechecks.push(fn);
  }

  /** Empties the registry; the defaults are reinstalled on the next use. */
  function __resetRegistry() {
    rules = [];
    rechecks = [];
    installed = false;
  }

  function rulesFor(from, to) {
    ensureInstalled();
    return rules.filter(function (r) {
      return (r.from === '*' || r.from === from) && (r.to === '*' || r.to === to);
    });
  }

  /* ------------------------------------------------------------- the graph */

  function sequenceOf(statusCode) {
    var status = Config.getStatus(statusCode);
    return status ? status.sequence : 0;
  }

  /** Going back down the sequence always needs a written reason (SPEC §6.3). */
  function isBackwards(from, to) {
    return sequenceOf(to) < sequenceOf(from);
  }

  /**
   * SPEC §7 — who may make a given move.
   *
   * The buyer who owns a Case drives it from end to end, closing it themselves
   * with the PR number EPICOR gave them; waiting for HEAD to type in a number
   * they never saw would only add a handoff. HEAD may move any Case, and nobody
   * moves a Case that is already closed — reopen is the only way out of that.
   */
  function canTransition(user, caseRecord, toStatus) {
    if (Auth.isTerminal(caseRecord.Status)) return false;          // reopen is the only way out
    if (Auth.isHead(user)) return true;
    return user.role === Auth.ROLES.BUYER && Auth.isOwner(user, caseRecord);
  }

  /** The buttons Case Detail should offer, already filtered by permission. */
  function allowedNextFor(user, caseRecord) {
    ensureInstalled();
    var status = Config.getStatus(caseRecord.Status);
    if (!status) return [];
    return status.allowedNext
      .filter(function (code) { return !!Config.getStatus(code); })
      .filter(function (code) { return canTransition(user, caseRecord, code); })
      .map(function (code) {
        var target = Config.getStatus(code);
        return {
          code: code,
          label: target.label,
          backwards: isBackwards(caseRecord.Status, code),
          requiresReason: isBackwards(caseRecord.Status, code) || code === 'CANCELLED'
        };
      });
  }

  /* --------------------------------------------------------- the transition */

  /**
   * `extraPatch` holds the fields that are part of the decision itself rather
   * than an edit that happens to come before it — the PR number and company a
   * Case is closed with. They travel with the transition so the rules see the
   * Case as it will be, and so one Repository.update() writes both: a Case can
   * never end up CLOSED without its PR number, or carrying a PR number it was
   * never closed under.
   */
  function transition(user, caseId, toStatus, version, reason, extraPatch) {
    var caseRecord = CaseService.requireCase(caseId);
    Auth.assertCanViewCase(user, caseRecord);

    var from = caseRecord.Status;
    if (from === toStatus) {
      throw Err.validation('งานนี้อยู่ในสถานะ ' + Auth.statusLabel(toStatus) + ' อยู่แล้ว');
    }
    var target = Config.getStatus(toStatus);
    if (!target) {
      throw Err.validation('ไม่รู้จักสถานะ ' + toStatus, { field: 'toStatus' });
    }
    if (Auth.isTerminal(from)) {
      throw Err.forbidden('งานนี้ปิดแล้ว ต้องให้หัวหน้าฝ่ายจัดซื้อกด Reopen ก่อนจึงจะเปลี่ยนสถานะได้');
    }
    var current = Config.getStatus(from);
    if (!current || current.allowedNext.indexOf(toStatus) === -1) {
      throw Err.ruleViolation('เปลี่ยนสถานะจาก ' + Auth.statusLabel(from) + ' เป็น ' +
        target.label + ' ไม่ได้', { from: from, to: toStatus, allowed: current ? current.allowedNext : [] });
    }
    if (!canTransition(user, caseRecord, toStatus)) {
      throw Err.forbidden('คุณไม่มีสิทธิ์เปลี่ยนสถานะงานนี้เป็น ' + target.label,
        { role: user.role, to: toStatus });
    }

    var explained = Utils.isBlank(reason) ? '' : String(reason).trim();
    if (isBackwards(from, toStatus) || toStatus === 'CANCELLED') {
      explained = Validation.requireReason(reason,
        toStatus === 'CANCELLED' ? 'การยกเลิกงาน' : 'การย้อนสถานะงาน');
    }

    var patch = {};
    var proposed = caseRecord;
    if (extraPatch) {
      Validation.validate('Cases', extraPatch, { partial: true });
      proposed = {};
      Object.keys(caseRecord).forEach(function (k) { proposed[k] = caseRecord[k]; });
      Object.keys(extraPatch).forEach(function (k) {
        patch[k] = extraPatch[k];
        proposed[k] = extraPatch[k];
      });
    }

    // Module rules run last, so a rule only ever sees a move that is otherwise legal.
    rulesFor(from, toStatus).forEach(function (rule) {
      rule.fn({ user: user, caseRecord: proposed, from: from, to: toStatus, reason: explained });
    });

    patch.Status = toStatus;
    if (target.isTerminal) patch.Closed_At = Utils.now();

    var updated = Repository.update('Cases', caseId, patch, version, {
      actor: user.email,
      reason: explained,
      caseId: caseId,
      fieldActions: { Status: ChangeLog.ACTIONS.STATUS_CHANGE }
    });

    if (typeof Notification !== 'undefined') {
      Notification.onStatusChanged(updated, from, toStatus, user);
    }
    return { caseRecord: Repository.toClient(updated), from: from, to: toStatus };
  }

  /**
   * Moves a Case back without a user asking, when the data no longer supports the
   * status it is in. Logged as SYSTEM so an auditor can tell the
   * difference between a person's decision and the system's correction.
   */
  function systemRevert(caseRecord, toStatus, message) {
    var target = Config.getStatus(toStatus);
    if (!target || caseRecord.Status === toStatus) return null;

    var updated = Repository.update('Cases', caseRecord.Case_ID, { Status: toStatus }, null, {
      actor: ChangeLog.SYSTEM_USER,
      reason: message,
      caseId: caseRecord.Case_ID,
      fieldActions: { Status: ChangeLog.ACTIONS.STATUS_CHANGE }
    });

    if (typeof Notification !== 'undefined') {
      Notification.onAutoRevert(updated, caseRecord.Status, toStatus, message);
    }
    return updated;
  }

  /** Runs every registered recheck and applies the first revert one of them asks for. */
  function runRechecks(caseRecord) {
    ensureInstalled();
    var messages = [];
    var reverted = null;
    for (var i = 0; i < rechecks.length; i++) {
      var verdict = rechecks[i](caseRecord);
      if (!verdict) continue;
      messages.push(verdict.message);
      if (verdict.revertTo && !reverted) {
        reverted = systemRevert(caseRecord, verdict.revertTo, verdict.message);
      }
    }
    return { reverted: reverted ? Repository.toClient(reverted) : null, messages: messages };
  }

  return {
    registerRule: registerRule,
    registerRecheck: registerRecheck,
    ensureInstalled: ensureInstalled,
    __resetRegistry: __resetRegistry,
    rulesFor: rulesFor,
    sequenceOf: sequenceOf,
    isBackwards: isBackwards,
    canTransition: canTransition,
    allowedNextFor: allowedNextFor,
    transition: transition,
    systemRevert: systemRevert,
    runRechecks: runRechecks
  };
})();
