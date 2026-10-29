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

  /**
   * The Config_Lists COMPANY code standing for "bought for more than one company
   * in the group". The company field holds one code, so this is how a shared
   * purchase is recorded — and the names of the companies then live only in the
   * closing note, which is why picking this code makes that note compulsory.
   *
   * This is the one place the literal is written down, and it is a literal: an
   * administrator who renames this code in the sheet turns the rule below into
   * a no-op, and shared purchases go back to being recorded as "several, unnamed".
   * Label_TH can be reworded freely; `MULTI` itself must stay.
   */
  var MULTI_COMPANY_CODE = 'MULTI';

  /* ----------------------------------------------------------------- close */

  /**
   * Closing a Case is one decision, not an edit followed by a status change: the
   * PR number, the buying company, what the purchase cost and who it was bought
   * from are the evidence the closure rests on. They therefore travel with the
   * transition, which writes them and the status in a single versioned update —
   * so a Case can never be found CLOSED without its PR number, nor holding a PR
   * number it was never closed under.
   *
   * The owning buyer does this themselves: they are the one holding the number
   * EPICOR gave them.
   *
   * All four requirements are checked here rather than in Rules.evaluate(), and
   * that placement is the whole design. Rules.recheck() reopens any closed Case
   * that no longer satisfies evaluate(), and dailyReminderJob() sweeps every
   * closed Case through it once a day. Putting the amount or the vendor there
   * would mean every Case closed before those columns existed stopped satisfying
   * the rule, and the system would quietly reopen finished purchases under the
   * SYSTEM name within 24 hours. Checking at the gate binds closures from today
   * without rewriting the past.
   */
  function closeCase(user, caseId, payload) {
    var data = payload || {};
    var caseRecord = CaseService.requireCase(caseId);
    Auth.assertCanEditCase(user, caseRecord);

    var prNo = Utils.isBlank(data.prNo) ? '' : String(data.prNo).trim();
    var company = Utils.isBlank(data.companyCode) ? '' : String(data.companyCode).trim();
    var reason = Utils.isBlank(data.reason) ? '' : String(data.reason).trim();
    var problems = [];
    if (!prNo) problems.push({ field: 'PR_No', message: 'กรุณากรอกเลข PR จากระบบ EPICOR' });
    if (!company) problems.push({ field: 'Company_Code', message: 'กรุณาเลือกบริษัทที่จัดซื้อ' });

    // Closing a Case is normally allowed with no note at all, and that stays
    // true — except for the one answer that is not an answer on its own. A Case
    // closed as MULTI with an empty note records that several companies were
    // involved and loses which, and nothing can recover that afterwards: the note
    // is only written at the moment of closing.
    if (company === MULTI_COMPANY_CODE && !reason) {
      problems.push({
        field: 'reason',
        message: 'เลือกหลายบริษัทแล้ว กรุณาระบุในหมายเหตุว่าเป็นบริษัทใดบ้าง'
      });
    }

    // Blank and zero are different answers. A purchase really can close at nothing
    // — a replacement under warranty, an insurance claim, goods thrown in free —
    // and Utils.isBlank(0) is false, so zero passes while an empty box does not.
    var amount = null;
    if (Utils.isBlank(data.finalAmount)) {
      problems.push({
        field: 'Final_Amount',
        message: 'กรุณากรอกมูลค่ารวมสุดท้ายของโครงการ'
      });
    } else {
      amount = Number(data.finalAmount);
      if (isNaN(amount)) {
        problems.push({ field: 'Final_Amount', message: 'มูลค่ารวมสุดท้ายต้องเป็นตัวเลข' });
      } else if (amount < 0) {
        problems.push({ field: 'Final_Amount', message: 'มูลค่ารวมสุดท้ายต้องไม่ติดลบ' });
      }
    }

    // The winning vendor has to be one on the register, for the same reason an
    // activity's vendor does: a free-text name makes the purchase unsearchable.
    // It is not restricted to vendors this Case logged activities against —
    // somebody who forgot to log the negotiation would otherwise be unable to
    // close their Case at all, and blocking a closure is worse than a stray id.
    var vendorId = Utils.isBlank(data.awardedVendorId) ? '' : String(data.awardedVendorId).trim();
    if (!vendorId) {
      problems.push({ field: 'Awarded_Vendor_ID', message: 'กรุณาเลือกผู้ขายที่ได้งาน' });
    } else if (!Repository.findById('Vendors', vendorId)) {
      problems.push({
        field: 'Awarded_Vendor_ID',
        message: 'ไม่พบผู้ขายรายนี้ในทะเบียนผู้ขาย'
      });
    }

    if (problems.length) {
      throw Err.validation(problems[0].message, { fields: problems });
    }

    var patch = {
      PR_No: prNo,
      Company_Code: company,
      Final_Amount: amount,
      Awarded_Vendor_ID: vendorId
    };
    if (!Utils.isBlank(data.prDate)) patch.PR_Date = data.prDate;

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
