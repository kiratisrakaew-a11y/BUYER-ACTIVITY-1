/**
 * Rules.js — the one business rule this system enforces: a Case may only be
 * closed once it carries the PR number EPICOR issued and the company that
 * actually bought the goods.
 *
 * Everything that talks about "is this Case ready to close?" reads from
 * evaluate(): the badge on My Cases, the panel on Case Detail, the check that
 * guards the transition, and the recheck that pulls a Case back open when the
 * ground shifts under it. That is deliberate — if those four ever disagreed,
 * the audit trail would be worthless.
 */
var Rules = (function () {

  var OPEN = 'OPEN';
  var CLOSED = 'CLOSED';

  /* ------------------------------------------------------------ evaluation */

  /**
   * The single answer to "may this Case be closed?".
   *
   * `caseRecord` may be the row as it stands or the row as it would stand after
   * a patch, which is what lets closeCase() check the PR number it is about to
   * write rather than the blank one already in the sheet.
   */
  function evaluate(caseRecord) {
    var blockers = [];
    var prNo = Utils.isBlank(caseRecord.PR_No) ? '' : String(caseRecord.PR_No).trim();
    var company = Utils.isBlank(caseRecord.Company_Code) ? '' : String(caseRecord.Company_Code).trim();

    if (!prNo) {
      blockers.push('ยังไม่ได้ระบุเลข PR จากระบบ EPICOR');
    }
    if (!company) {
      blockers.push('ยังไม่ได้ระบุบริษัทที่จัดซื้อ');
    } else if (!Config.isValidCode('COMPANY', company)) {
      blockers.push('บริษัทที่จัดซื้อ "' + company + '" ไม่อยู่ในรายการที่ผู้ดูแลระบบกำหนดไว้');
    }

    return {
      ok: blockers.length === 0,
      blockers: blockers,
      prNo: prNo,
      company: company,
      companyLabel: company ? Config.labelOf('COMPANY', company) : ''
    };
  }

  /**
   * A PR number identifies one purchase requisition in EPICOR, so two Cases
   * claiming the same one means one of them is wrong. Blocking is the point:
   * a warning here would let the mistake reach the month-end reconciliation.
   */
  function assertPrNoUnique(prNo, exceptCaseId) {
    var wanted = String(prNo || '').trim().toUpperCase();
    if (!wanted) return;
    var clash = null;
    Repository.query('Cases').forEach(function (c) {
      if (clash || c.Case_ID === exceptCaseId) return;
      if (String(c.PR_No || '').trim().toUpperCase() === wanted) clash = c;
    });
    if (clash) {
      throw Err.duplicate('เลข PR "' + String(prNo).trim() + '" ถูกใช้กับงาน ' + clash.Case_ID + ' แล้ว',
        { field: 'PR_No', caseId: clash.Case_ID });
    }
  }

  /**
   * Per-Case readiness for the My Cases and Team View badges.
   *
   * Everything the rule needs already lives on the Case row the list just read,
   * so this costs no extra sheet access and — unlike the quote counting it
   * replaces — never has to approximate to stay inside the execution limit.
   */
  function summarizeCases(cases) {
    var out = {};
    (cases || []).forEach(function (c) {
      var verdict = evaluate(c);
      out[c.Case_ID] = {
        ok: verdict.ok,
        blockers: verdict.blockers,
        prNo: verdict.prNo,
        company: verdict.company,
        companyLabel: verdict.companyLabel
      };
    });
    return out;
  }

  /* ----------------------------------------------------- the transition rule */

  /** Registered on OPEN -> CLOSED; throws to block the move. */
  function check(ctx) {
    var verdict = evaluate(ctx.caseRecord);
    if (!verdict.ok) {
      throw Err.ruleViolation('ปิดงานไม่ได้: ' + verdict.blockers.join(' · '),
        { blockers: verdict.blockers });
    }
    assertPrNoUnique(verdict.prNo, ctx.caseRecord.Case_ID);
  }

  /**
   * A closed Case whose PR number was cleared no longer meets the rule it was
   * closed under, so it goes back to OPEN under the SYSTEM name rather than
   * sitting in the reports as a finished purchase that cannot be traced.
   */
  function recheck(caseRecord) {
    if (caseRecord.Status !== CLOSED) return null;
    var verdict = evaluate(caseRecord);
    if (verdict.ok) return null;
    return {
      revertTo: OPEN,
      message: 'ระบบเปิดงานกลับอัตโนมัติ: ' + verdict.blockers.join(' · ')
    };
  }

  /**
   * The single entry point every service calls after a write that could change
   * the answer. Returns the messages to surface, and the reverted Case if any.
   */
  function recheckCaseRules(caseId) {
    var caseRecord = Repository.findById('Cases', caseId);
    if (!caseRecord) return { reverted: null, messages: [] };
    return StatusEngine.runRechecks(caseRecord);
  }

  /**
   * Called by Bootstrap at run time, never while this file is being loaded —
   * Apps Script picks its own file order and a load-time call works only by luck.
   */
  function install() {
    StatusEngine.registerRule(OPEN, CLOSED, check);
    StatusEngine.registerRecheck(recheck);
  }

  return {
    OPEN: OPEN,
    CLOSED: CLOSED,
    evaluate: evaluate,
    summarizeCases: summarizeCases,
    assertPrNoUnique: assertPrNoUnique,
    recheckCaseRules: recheckCaseRules,
    check: check,
    recheck: recheck,
    install: install
  };
})();
