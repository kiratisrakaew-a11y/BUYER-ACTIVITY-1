/**
 * CaseWorkflow.js — the Case operations that are decisions rather than edits:
 * closing a Case against the PR that EPICOR issued, handing work to another
 * buyer, and reopening a Case that was already closed.
 *
 * Each one writes its own Change_Log action (STATUS_CHANGE, REASSIGN, REOPEN)
 * so an auditor can read the history of a Case without having to interpret raw
 * field changes (SPEC §5.4).
 */
var CaseWorkflow = (function () {

  /* ----------------------------------------------------------------- close */

  /**
   * Closing a Case is one decision, not an edit followed by a status change:
   * the PR number and the buying company are the evidence the closure rests on.
   * They therefore travel with the transition, which writes them and the status
   * in a single versioned update — so a Case can never be found CLOSED without
   * its PR number, nor holding a PR number it was never closed under.
   *
   * The owning buyer does this themselves: they are the one holding the number
   * EPICOR gave them.
   */
  function closeCase(user, caseId, payload) {
    var data = payload || {};
    var caseRecord = CaseService.requireCase(caseId);
    Auth.assertCanEditCase(user, caseRecord);

    var prNo = Utils.isBlank(data.prNo) ? '' : String(data.prNo).trim();
    var company = Utils.isBlank(data.companyCode) ? '' : String(data.companyCode).trim();
    var problems = [];
    if (!prNo) problems.push({ field: 'PR_No', message: 'กรุณากรอกเลข PR จากระบบ EPICOR' });
    if (!company) problems.push({ field: 'Company_Code', message: 'กรุณาเลือกบริษัทที่จัดซื้อ' });
    if (problems.length) {
      throw Err.validation(problems[0].message, { fields: problems });
    }

    var patch = { PR_No: prNo, Company_Code: company };
    if (!Utils.isBlank(data.prDate)) patch.PR_Date = data.prDate;

    var reason = Utils.isBlank(data.reason) ? '' : String(data.reason).trim();
    var result = StatusEngine.transition(user, caseId, 'CLOSED', data.version, reason, patch);
    return { caseRecord: result.caseRecord, from: result.from, to: result.to, warnings: [] };
  }

  /* -------------------------------------------------------------- reassign */

  /** SPEC §7 — only HEAD moves work between buyers, and always with a reason. */
  function reassign(user, caseId, newOwnerEmail, reason) {
    Auth.requireRole(user, [Auth.ROLES.HEAD]);
    var caseRecord = CaseService.requireCase(caseId);
    var explained = Validation.requireReason(reason, 'การโอนงาน');

    if (Auth.isTerminal(caseRecord.Status)) {
      throw Err.forbidden('งานที่ปิดแล้วโอนให้ผู้อื่นไม่ได้', { caseId: caseId });
    }
    var target = String(newOwnerEmail || '').trim().toLowerCase();
    var newOwner = Auth.findUser(target);
    if (!newOwner || newOwner.Is_Active !== true) {
      throw Err.validation('ไม่พบผู้ใช้งานที่ใช้งานอยู่ตามอีเมล ' + newOwnerEmail, { field: 'newOwnerEmail' });
    }
    var role = String(newOwner.Role).toUpperCase();
    if (role !== Auth.ROLES.BUYER && role !== Auth.ROLES.HEAD) {
      throw Err.validation('โอนงานให้ได้เฉพาะผู้ที่มีบทบาท BUYER หรือ HEAD เท่านั้น', { role: role });
    }
    if (target === String(caseRecord.Buyer_Owner).toLowerCase()) {
      throw Err.validation('ผู้รับโอนเป็นเจ้าของงานนี้อยู่แล้ว');
    }

    var previousOwner = caseRecord.Buyer_Owner;
    var updated = Repository.update('Cases', caseId, { Buyer_Owner: target }, null, {
      actor: user.email,
      reason: explained,
      caseId: caseId,
      action: ChangeLog.ACTIONS.REASSIGN
    });

    if (typeof Notification !== 'undefined') {
      Notification.onReassigned(updated, previousOwner, target, user, explained);
    }
    return { caseRecord: Repository.toClient(updated), previousOwner: previousOwner };
  }

  /* ---------------------------------------------------------------- reopen */

  /**
   * SPEC §6.3 — a closed or cancelled Case is frozen until HEAD reopens it.
   * The status to return to is read back out of the Change_Log, so reopening
   * restores where the work actually was rather than a guess.
   */
  function reopen(user, caseId, reason) {
    Auth.requireRole(user, [Auth.ROLES.HEAD]);
    var caseRecord = CaseService.requireCase(caseId);
    var explained = Validation.requireReason(reason, 'การเปิดงานที่ปิดไปแล้ว');

    if (!Auth.isTerminal(caseRecord.Status)) {
      throw Err.validation('งานนี้ยังไม่ได้ปิด จึงไม่ต้อง Reopen', { status: caseRecord.Status });
    }

    var previous = ChangeLog.previousStatusBefore(caseId, caseRecord.Status);
    if (!previous || !Config.getStatus(previous) || Auth.isTerminal(previous)) {
      previous = CaseService.INITIAL_STATUS;
    }

    var updated = Repository.update('Cases', caseId, {
      Status: previous,
      Closed_At: ''
    }, null, {
      actor: user.email,
      reason: explained,
      caseId: caseId,
      // Only the status move is the REOPEN event; clearing Closed_At is a
      // consequence of it and reads as an ordinary field change.
      fieldActions: { Status: ChangeLog.ACTIONS.REOPEN }
    });

    var recheck = Rules.recheckCaseRules(caseId);
    if (typeof Notification !== 'undefined') {
      Notification.onReopened(updated, caseRecord.Status, previous, user, explained);
    }
    return {
      caseRecord: Repository.toClient(recheck.reverted || updated),
      restoredTo: previous,
      warnings: recheck.messages
    };
  }

  return {
    closeCase: closeCase,
    reassign: reassign,
    reopen: reopen
  };
})();
