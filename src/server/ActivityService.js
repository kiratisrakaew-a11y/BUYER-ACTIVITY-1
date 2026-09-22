/**
 * ActivityService.js — the record of what the buyer actually did (SPEC §5.1).
 *
 * Two details carry the audit weight:
 *   Performed_By is always the person who pressed save, never the Case owner,
 *   because buyers cover for each other and the log has to say who really acted.
 *   Activity_Date is the time the work happened, which can differ from the time
 *   it was typed in; Created_At keeps the latter.
 */
var ActivityService = (function () {

  var MODULE = 'M1';

  var EDITABLE_FIELDS = [
    'Vendor_ID', 'Activity_Date', 'Activity_Type', 'Channel', 'Activity_Description',
    'Next_Action', 'Next_Action_Date', 'Next_Action_Done', 'Attachment_URL'
  ];

  function listForCase(caseId) {
    return Repository.queryByCase('Activities', caseId).sort(function (a, b) {
      return Utils.toDate(b.Activity_Date).getTime() - Utils.toDate(a.Activity_Date).getTime();
    });
  }

  function save(user, caseId, payload, version) {
    var caseRecord = CaseService.requireCase(caseId);

    if (Utils.isBlank(payload.Activity_ID)) {
      // SPEC §7 — recording an activity on a colleague's Case is allowed.
      Auth.assertCanLogActivity(user, caseRecord);
      return { activity: Repository.toClient(insert(user, caseRecord, payload)) };
    }

    var existing = Repository.requireById('Activities', payload.Activity_ID);
    if (existing.Case_ID !== caseRecord.Case_ID) {
      throw Err.validation('กิจกรรมนี้ไม่ได้อยู่ในงาน ' + caseId);
    }
    assertCanEditActivity(user, caseRecord, existing);
    return { activity: Repository.toClient(update(user, caseRecord, existing, payload, version)) };
  }

  /**
   * The Case owner and HEAD may edit any activity on the Case; a buyer who
   * recorded an entry on someone else's Case may correct their own, which is the
   * natural companion to being allowed to write it in the first place.
   */
  function assertCanEditActivity(user, caseRecord, activity) {
    if (Auth.canEditCase(user, caseRecord)) return;
    var isAuthor = String(activity.Performed_By || '').toLowerCase() === user.email;
    if (isAuthor && Auth.canLogActivity(user, caseRecord)) return;
    throw Err.forbidden('แก้ไขได้เฉพาะกิจกรรมที่คุณบันทึกเอง หรือกิจกรรมในงานที่คุณเป็นเจ้าของ', {
      caseId: caseRecord.Case_ID,
      performedBy: activity.Performed_By
    });
  }

  function insert(user, caseRecord, payload) {
    var values = pick(payload, EDITABLE_FIELDS);
    values.Case_ID = caseRecord.Case_ID;
    values.Module = MODULE;
    values.Performed_By = user.email;                  // never taken from the client
    if (Utils.isBlank(values.Activity_Date)) values.Activity_Date = Utils.now();
    values.Next_Action_Done = Utils.toBool(values.Next_Action_Done);

    assertNextActionConsistent(values);
    assertVendorOnCase(caseRecord.Case_ID, values.Vendor_ID);
    Validation.validate('Activities', values, { partial: false });

    return Repository.insert('Activities', values, { actor: user.email, caseId: caseRecord.Case_ID });
  }

  function update(user, caseRecord, existing, payload, version) {
    var values = pick(payload, EDITABLE_FIELDS);
    if (Object.prototype.hasOwnProperty.call(values, 'Next_Action_Done')) {
      values.Next_Action_Done = Utils.toBool(values.Next_Action_Done);
    }
    assertNextActionConsistent(Object.assign({}, existing, values));
    assertVendorOnCase(caseRecord.Case_ID, values.Vendor_ID);
    Validation.validate('Activities', values, { partial: true, existing: existing });

    return Repository.update('Activities', existing.Activity_ID, values, version, {
      actor: user.email,
      caseId: caseRecord.Case_ID
    });
  }

  /** SPEC §5.1 — a Next_Action without a due date can never be chased. */
  function assertNextActionConsistent(values) {
    if (Utils.isBlank(values.Next_Action)) return;
    if (Utils.isBlank(values.Next_Action_Date)) {
      throw Err.validation('เมื่อระบุสิ่งที่ต้องทำต่อ ต้องระบุวันที่ที่ต้องทำด้วย',
        { field: 'Next_Action_Date' });
    }
  }

  /**
   * An activity may name the vendor it was about, and that vendor has to be one
   * on the register — a free-text name would make the timeline unsearchable.
   */
  function assertVendorOnCase(caseId, vendorId) {
    if (Utils.isBlank(vendorId)) return;
    if (!Repository.findById('Vendors', vendorId)) {
      throw Err.validation('ไม่พบผู้ขายรายนี้ในทะเบียนผู้ขาย', { field: 'Vendor_ID' });
    }
  }

  /** The tick box on the timeline. Kept separate so it needs no full form round trip. */
  function setNextActionDone(user, activityId, done, version) {
    var activity = Repository.requireById('Activities', activityId);
    var caseRecord = CaseService.requireCase(activity.Case_ID);
    assertCanEditActivity(user, caseRecord, activity);

    var updated = Repository.update('Activities', activityId,
      { Next_Action_Done: Utils.toBool(done) }, version,
      { actor: user.email, caseId: activity.Case_ID });
    return { activity: Repository.toClient(updated) };
  }

  function remove(user, activityId, version, reason) {
    var activity = Repository.requireById('Activities', activityId);
    var caseRecord = CaseService.requireCase(activity.Case_ID);
    assertCanEditActivity(user, caseRecord, activity);
    var explained = Validation.requireReason(reason, 'การลบกิจกรรม');

    Repository.softDelete('Activities', activityId, version, {
      actor: user.email, reason: explained, caseId: activity.Case_ID
    });
    return { deleted: true };
  }

  /** Outstanding next actions across Cases, used by the daily reminder (SPEC §11). */
  function outstandingNextActions(filterFn) {
    return Repository.query('Activities', {
      where: function (a) {
        if (Utils.isBlank(a.Next_Action) || a.Next_Action_Done === true) return false;
        return !filterFn || filterFn(a);
      }
    });
  }

  function pick(source, fields) {
    var out = {};
    fields.forEach(function (field) {
      if (source && Object.prototype.hasOwnProperty.call(source, field)) out[field] = source[field];
    });
    return out;
  }

  return {
    MODULE: MODULE,
    EDITABLE_FIELDS: EDITABLE_FIELDS,
    listForCase: listForCase,
    save: save,
    setNextActionDone: setNextActionDone,
    remove: remove,
    outstandingNextActions: outstandingNextActions
  };
})();
