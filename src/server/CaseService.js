/**
 * CaseService.js — the Case is the spine of the system (SPEC §2.1).
 *
 * Every other table hangs off Case_ID, which the system issues itself, because
 * the PR number only exists once EPICOR has issued it — at the end of the work,
 * not the beginning — and so cannot identify the work while it is being done.
 */
var CaseService = (function () {

  var INITIAL_STATUS = 'OPEN';

  /**
   * The one budget type that has no sub types under it, because "other" is what
   * you pick when none of them fit. A Case on it carries a written description
   * instead, which the requesting department fills in.
   *
   * This is a Code in the Config_Lists sheet, so an administrator can rename its
   * Label_TH freely — but renaming the Code itself would leave this rule matching
   * nothing, and the requirement would disappear without a word. That caution is
   * written down in README §5.
   */
  var OTHER_BUDGET_TYPE = 'OTHER';

  /** Fields a buyer fills in. Status, owner, and the PR number a Case is closed
   *  with move through their own operations so each carries its own rule and
   *  audit action. */
  var EDITABLE_FIELDS = [
    'Request_Date', 'Request_Ref', 'Requester_Name', 'Requester_Email',
    'Department_Code', 'Method', 'Budget_Type', 'Sub_Type', 'Budget_Type_Other', 'Description',
    'Required_Date', 'Intake_Complete', 'Intake_Note'
  ];

  var DEFAULT_LIST_LIMIT = 300;

  /* ----------------------------------------------------------------- create */

  function create(user, payload) {
    Auth.requireRole(user, [Auth.ROLES.BUYER, Auth.ROLES.HEAD]);

    var values = pick(payload, EDITABLE_FIELDS);
    values.Buyer_Owner = user.email;
    values.Status = INITIAL_STATUS;
    values.Intake_Complete = Utils.toBool(values.Intake_Complete);
    applyBudgetType(values, values);
    Validation.validate('Cases', values, { partial: false });

    var warnings = duplicateCaseWarnings(values.Request_Ref);

    // The folder is named after the Case, so the id is reserved first and the
    // record is then written once: one version, one CREATE entry in the log.
    var caseId = IdGenerator.next('Cases');
    // SPEC §3 — a Drive folder per Case, created at the moment the Case is opened.
    var folderId = DriveService.createCaseFolder(caseId, values.Description);
    if (!folderId) {
      warnings.push('สร้างโฟลเดอร์ใน Google Drive ไม่สำเร็จ ระบบจะสร้างให้อีกครั้งเมื่อมีการอัปโหลดไฟล์');
    }
    values.Drive_Folder_ID = folderId || '';

    var created = Repository.insert('Cases', values, { actor: user.email, id: caseId });
    return { caseRecord: Repository.toClient(created), warnings: warnings };
  }

  /**
   * A Case carries a Sub_Type or a written budget description — never both, and
   * never neither.
   *
   * `patch` is what will be written and `proposed` is the Case as it will read
   * afterwards; on a create they are the same object. Clearing the field that no
   * longer applies is part of the rule rather than a separate tidy-up: a Case
   * switched from CAPEX to OTHER would otherwise keep its old Sub_Type sitting in
   * the sheet, and every report reading that column would still count it as one.
   */
  function applyBudgetType(patch, proposed) {
    // Neither field is being touched and the Case is already consistent.
    if (!Object.prototype.hasOwnProperty.call(patch, 'Budget_Type') &&
        !Object.prototype.hasOwnProperty.call(patch, 'Sub_Type') &&
        !Object.prototype.hasOwnProperty.call(patch, 'Budget_Type_Other')) {
      return patch;
    }

    if (proposed.Budget_Type === OTHER_BUDGET_TYPE) {
      if (Utils.isBlank(proposed.Budget_Type_Other)) {
        throw Err.validation('เมื่อเลือกประเภทงบ "' + Config.labelOf('BUDGET_TYPE', OTHER_BUDGET_TYPE) +
          '" ต้องระบุด้วยว่าเป็นงบประเภทใด', { field: 'Budget_Type_Other' });
      }
      patch.Sub_Type = '';
    } else {
      if (Utils.isBlank(proposed.Sub_Type)) {
        throw Err.validation('กรุณาเลือกประเภทย่อย', { field: 'Sub_Type' });
      }
      patch.Budget_Type_Other = '';
    }
    return patch;
  }

  /** SPEC §6.2 — opening a Case whose Request_Ref matches a live one is a warning,
   *  never a block: the same memo legitimately spawns more than one purchase. */
  function duplicateCaseWarnings(requestRef) {
    if (Utils.isBlank(requestRef)) return [];
    var matches = Repository.query('Cases', {
      where: function (c) {
        return Utils.normalizeForCompare(c.Request_Ref) === Utils.normalizeForCompare(requestRef) &&
          !Auth.isTerminal(c.Status);
      }
    });
    if (matches.length === 0) return [];
    return ['มีงานที่ยังเปิดอยู่และใช้เอกสารอ้างอิงเดียวกัน: ' +
      matches.map(function (c) { return c.Case_ID; }).join(', ')];
  }

  /* ------------------------------------------------------------------- read */

  function requireCase(caseId) {
    return Repository.requireById('Cases', caseId);
  }

  function getForView(user, caseId) {
    return Auth.assertCanViewCase(user, requireCase(caseId));
  }

  function getForEdit(user, caseId) {
    return Auth.assertCanEditCase(user, requireCase(caseId));
  }

  /**
   * Everything Case Detail needs in one round trip (SPEC §9 api_getCase):
   * the Case, its items, the vendors invited with their quote lines, activities,
   * document references, and the totals computed from quantity x unit price.
   */
  function getBundle(user, caseId) {
    var caseRecord = getForView(user, caseId);
    var activities = Repository.queryByCase('Activities', caseId);

    activities.sort(function (a, b) {
      return Utils.toDate(b.Activity_Date).getTime() - Utils.toDate(a.Activity_Date).getTime();
    });

    return {
      caseRecord: Repository.toClient(caseRecord),
      activities: activities.map(Repository.toClient),
      vendorNames: vendorNamesFor(activities),
      driveFolderUrl: DriveService.folderUrl(caseRecord.Drive_Folder_ID),
      permissions: {
        canEdit: Auth.canEditCase(user, caseRecord),
        canLogActivity: Auth.canLogActivity(user, caseRecord),
        isOwner: Auth.isOwner(user, caseRecord)
      },
      rules: typeof Rules === 'undefined' ? null : Rules.evaluate(caseRecord),
      nextStatuses: typeof StatusEngine === 'undefined' ? [] : StatusEngine.allowedNextFor(user, caseRecord)
    };
  }

  function update(user, caseId, patch, version, reason) {
    var caseRecord = getForEdit(user, caseId);
    var clean = pick(patch, EDITABLE_FIELDS);
    if (Object.keys(clean).length === 0) {
      throw Err.validation('ไม่มีข้อมูลที่แก้ไขได้ในคำขอนี้');
    }
    if (Object.prototype.hasOwnProperty.call(clean, 'Intake_Complete')) {
      clean.Intake_Complete = Utils.toBool(clean.Intake_Complete);
    }
    // The rule spans two fields, so it has to see the Case as it will be, not the
    // handful of fields this request happens to carry.
    applyBudgetType(clean, Object.assign({}, caseRecord, clean));
    Validation.validate('Cases', clean, { partial: true, existing: caseRecord });

    var updated = Repository.update('Cases', caseId, clean, version, {
      actor: user.email,
      reason: Utils.isBlank(reason) ? '' : String(reason).trim(),
      caseId: caseId
    });

    var warnings = [];
    if (typeof Rules !== 'undefined') {
      warnings = warnings.concat(Rules.recheckCaseRules(caseId).messages);
    }
    return { caseRecord: Repository.toClient(updated), warnings: warnings };
  }

  /* ------------------------------------------------------------------- list */

  /**
   * The Case list, filtered and scoped by what the caller is allowed to see.
   * filter: { scope: 'mine'|'all', status, budgetType, from, to, q, owner, limit }
   */
  function list(user, filter) {
    var f = filter || {};
    var scope = f.scope || 'mine';
    var from = Utils.startOfDay(f.from);
    var to = Utils.startOfDay(f.to);
    var q = Utils.isBlank(f.q) ? '' : String(f.q).trim().toLowerCase();

    var cases = Repository.query('Cases', {
      where: function (c) {
        if (scope === 'mine' && String(c.Buyer_Owner).toLowerCase() !== user.email) return false;
        if (!Utils.isBlank(f.owner) && String(c.Buyer_Owner).toLowerCase() !== String(f.owner).toLowerCase()) return false;
        if (!Utils.isBlank(f.status) && c.Status !== f.status) return false;
        if (!Utils.isBlank(f.budgetType) && c.Budget_Type !== f.budgetType) return false;
        var requested = Utils.startOfDay(c.Request_Date);
        if (from && (!requested || requested.getTime() < from.getTime())) return false;
        if (to && (!requested || requested.getTime() > to.getTime())) return false;
        if (q && !matchesText(c, q)) return false;
        return Auth.canViewCase(user, c);
      }
    });

    cases.sort(function (a, b) {
      var left = Utils.toDate(a.Created_At);
      var right = Utils.toDate(b.Created_At);
      return (right ? right.getTime() : 0) - (left ? left.getTime() : 0);
    });

    var limit = Number(f.limit) > 0 ? Number(f.limit) : DEFAULT_LIST_LIMIT;
    var truncated = cases.length > limit;
    var page = cases.slice(0, limit);
    var nextActions = nextActionByCase(page.map(function (c) { return c.Case_ID; }));
    var readiness = typeof Rules === 'undefined' ? {} : Rules.summarizeCases(page);

    return {
      total: cases.length,
      truncated: truncated,
      cases: page.map(function (c) {
        return Object.assign(Repository.toClient(c), {
          nextAction: nextActions[c.Case_ID] || null,
          readiness: readiness[c.Case_ID] || null,
          canEdit: Auth.canEditCase(user, c)
        });
      })
    };
  }

  function matchesText(c, q) {
    return [c.Case_ID, c.Description, c.Request_Ref, c.Requester_Name, c.Buyer_Owner, c.PR_No]
      .some(function (v) { return String(v || '').toLowerCase().indexOf(q) !== -1; });
  }

  /**
   * The most pressing outstanding Next_Action per Case, which is what My Cases
   * highlights when it is overdue (SPEC §8.1).
   *
   * A follow-up with no due date is real work, but it is not the one a buyer
   * needs reminding of, so it only ever shows when the Case has nothing dated
   * outstanding. Ranking the other way round would let an undated note hide a
   * deadline that has already passed.
   */
  function nextActionByCase(caseIds) {
    if (caseIds.length === 0) return {};
    var wanted = {};
    caseIds.forEach(function (id) { wanted[id] = true; });

    var out = {};
    var today = Utils.today();
    Repository.query('Activities', {
      where: function (a) {
        return wanted[a.Case_ID] === true && !Utils.isBlank(a.Next_Action) && a.Next_Action_Done !== true;
      }
    }).forEach(function (a) {
      var due = Utils.startOfDay(a.Next_Action_Date);
      var current = out[a.Case_ID];
      if (current && !isMorePressing(due, current.dueTime)) return;
      out[a.Case_ID] = {
        activityId: a.Activity_ID,
        text: a.Next_Action,
        date: due ? due.toISOString() : '',
        dueTime: due ? due.getTime() : null,
        overdue: !!due && due.getTime() < today.getTime()
      };
    });
    return out;
  }

  /** Dated beats undated; between two dated ones the earlier due date wins. */
  function isMorePressing(due, heldDueTime) {
    if (heldDueTime === null) return true;
    if (due === null) return false;
    return due.getTime() < heldDueTime;
  }

  /**
   * Beyond this many distinct vendors on one Case, looking each one up costs more
   * in round trips than reading the register once and picking from it.
   */
  var MAX_VENDOR_LOOKUPS = 10;

  /**
   * Vendor_ID -> name for the vendors named in these activities, so the timeline
   * can show who was contacted without the client fetching the register.
   *
   * The ids are known before anything is read, so each one is fetched by id
   * rather than by reading the register and discarding nearly all of it. That
   * keeps the cost proportional to the Case — a Case naming three vendors costs
   * the same whether the register holds fifty or five thousand.
   *
   * Deleted vendors are included on purpose: a vendor removed from the register
   * is still the vendor that old activity was about, and the timeline should say
   * so by name rather than falling back to a bare id.
   */
  function vendorNamesFor(activities) {
    var ids = Utils.unique(activities
      .map(function (a) { return a.Vendor_ID; })
      .filter(function (id) { return !Utils.isBlank(id); }));
    if (ids.length === 0) return {};

    var names = {};
    if (ids.length > MAX_VENDOR_LOOKUPS) {
      var wanted = {};
      ids.forEach(function (id) { wanted[id] = true; });
      Repository.query('Vendors', { includeDeleted: true }).forEach(function (v) {
        if (wanted[v.Vendor_ID]) names[v.Vendor_ID] = v.Vendor_Name;
      });
      return names;
    }

    ids.forEach(function (id) {
      // query() rather than findById(): a duplicated id is a problem for
      // verifyDeployment to report, not a reason to take the Case page down over
      // a display name.
      var found = Repository.query('Vendors', {
        indexColumn: 'Vendor_ID', indexValue: id, includeDeleted: true
      });
      if (found.length) names[id] = found[0].Vendor_Name;
    });
    return names;
  }

  /* ---------------------------------------------------------------- helpers */

  function pick(source, fields) {
    var out = {};
    fields.forEach(function (field) {
      if (source && Object.prototype.hasOwnProperty.call(source, field)) out[field] = source[field];
    });
    return out;
  }

  function sortBy(rows, field) {
    return rows.slice().sort(function (a, b) {
      var left = a[field];
      var right = b[field];
      if (left instanceof Date || right instanceof Date) {
        return (Utils.toDate(left) ? Utils.toDate(left).getTime() : 0) -
          (Utils.toDate(right) ? Utils.toDate(right).getTime() : 0);
      }
      return (Number(left) || 0) - (Number(right) || 0);
    });
  }

  return {
    INITIAL_STATUS: INITIAL_STATUS,
    OTHER_BUDGET_TYPE: OTHER_BUDGET_TYPE,
    MAX_VENDOR_LOOKUPS: MAX_VENDOR_LOOKUPS,
    EDITABLE_FIELDS: EDITABLE_FIELDS,
    create: create,
    requireCase: requireCase,
    getForView: getForView,
    getForEdit: getForEdit,
    getBundle: getBundle,
    update: update,
    list: list,
    duplicateCaseWarnings: duplicateCaseWarnings,
    nextActionByCase: nextActionByCase
  };
})();
