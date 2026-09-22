/**
 * Validation.js — server-side checks on everything that arrives from a browser.
 *
 * Client-side validation is a convenience; this file is the control (SPEC §11).
 * Rules are expressed as data wherever possible so a new module extends them by
 * adding an entry rather than by editing a function.
 */
var Validation = (function () {

  /**
   * SPEC §6.3 — changing one of these fields is only allowed with a written reason.
   * Status reversals, reassignment and reopen carry their own checks in
   * StatusEngine and CaseWorkflow, because "reason" there depends on direction.
   */
  var REASON_REQUIRED_FIELDS = {
    Cases: ['Buyer_Owner']
  };

  var EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  /**
   * Validates a payload against the schema.
   * opts: { partial } — when true, only the supplied fields are checked, which is
   * what an update needs; a create checks every required column.
   * Throws VALIDATION listing every problem at once, so the user fixes the form in one pass.
   */
  function validate(tableName, payload, opts) {
    var options = opts || {};
    var problems = [];
    var columns = Schema.getColumns(tableName);

    columns.forEach(function (column) {
      if (Schema.isAuditColumn(column.name)) return;
      var supplied = Object.prototype.hasOwnProperty.call(payload, column.name);
      if (options.partial && !supplied) return;

      var value = payload[column.name];
      var table = Schema.getTable(tableName);
      var systemIssued = column.name === table.pk && !!table.id;

      if (column.required && !systemIssued && Utils.isBlank(value)) {
        if (column.type === 'bool' && value === false) return;   // false is a real answer
        problems.push({ field: column.name, message: 'กรุณากรอก ' + column.name });
        return;
      }
      if (Utils.isBlank(value)) return;

      checkType(column, value, problems);
      checkLength(column, value, problems);
      checkList(tableName, column, value, payload, options, problems);
    });

    if (problems.length) {
      throw Err.validation(problems[0].message, { fields: problems });
    }
    return payload;
  }

  function checkType(column, value, problems) {
    switch (column.type) {
      case 'number':
      case 'int': {
        var n = Number(value);
        if (isNaN(n)) {
          problems.push({ field: column.name, message: column.name + ' ต้องเป็นตัวเลข' });
          return;
        }
        if (column.min !== undefined) {
          var tooSmall = column.exclusiveMin ? n <= column.min : n < column.min;
          if (tooSmall) {
            problems.push({
              field: column.name,
              message: column.name + ' ต้องมากกว่า' + (column.exclusiveMin ? '' : 'หรือเท่ากับ') + ' ' + column.min
            });
          }
        }
        break;
      }
      case 'date':
      case 'datetime':
        if (!Utils.toDate(value)) {
          problems.push({ field: column.name, message: column.name + ' ไม่ใช่วันที่ที่ถูกต้อง' });
        }
        break;
      case 'email':
        if (!EMAIL_PATTERN.test(String(value).trim())) {
          problems.push({ field: column.name, message: column.name + ' ไม่ใช่อีเมลที่ถูกต้อง' });
        }
        break;
      case 'url':
        if (!/^https?:\/\//i.test(String(value).trim())) {
          problems.push({ field: column.name, message: column.name + ' ต้องขึ้นต้นด้วย http:// หรือ https://' });
        }
        break;
      default:
        break;
    }
  }

  function checkLength(column, value, problems) {
    if (!column.max) return;
    if (String(value).length > column.max) {
      problems.push({ field: column.name, message: column.name + ' ยาวเกิน ' + column.max + ' ตัวอักษร' });
    }
  }

  /** A `code` column must hold a code that exists in its Config_Lists list. */
  function checkList(tableName, column, value, payload, options, problems) {
    if (column.type !== 'code' || !column.list) return;
    var parentValue = null;
    if (column.parent) {
      parentValue = Object.prototype.hasOwnProperty.call(payload, column.parent)
        ? payload[column.parent]
        : (options.existing ? options.existing[column.parent] : null);
    }
    if (!Config.isValidCode(column.list, value, parentValue)) {
      problems.push({
        field: column.name,
        message: column.name + ' "' + value + '" ไม่อยู่ในรายการที่กำหนด' +
          (parentValue ? ' สำหรับ ' + parentValue : '')
      });
    }
  }

  /** Codes that are not driven by Config_Lists still have to be one of a known set. */
  function assertOneOf(fieldName, value, allowed) {
    if (allowed.indexOf(value) === -1) {
      throw Err.validation(fieldName + ' ต้องเป็นค่าใดค่าหนึ่งใน: ' + allowed.join(', '), { field: fieldName });
    }
    return value;
  }

  function requireReason(reason, what) {
    if (Utils.isBlank(reason)) {
      throw Err.validation('กรุณาระบุเหตุผลสำหรับ' + (what || 'การทำรายการนี้'), { field: 'reason' });
    }
    return String(reason).trim();
  }

  /**
   * Enforces SPEC §6.3 for edits: if the patch really changes a field listed in
   * REASON_REQUIRED_FIELDS, a reason is mandatory. A field set to the value it
   * already has is not a change and needs nothing.
   */
  function assertReasonForPatch(tableName, existing, patch, reason) {
    var watched = REASON_REQUIRED_FIELDS[tableName] || [];
    var changed = watched.filter(function (field) {
      if (!Object.prototype.hasOwnProperty.call(patch, field)) return false;
      return Utils.normalizeForCompare(existing[field]) !== Utils.normalizeForCompare(patch[field]);
    });
    if (changed.length === 0) return '';
    if (Utils.isBlank(reason)) {
      throw Err.validation('การแก้ไข ' + changed.join(', ') + ' ต้องระบุเหตุผล',
        { field: 'reason', fields: changed });
    }
    return String(reason).trim();
  }

  /** Rejects a patch that tries to write a column no M1 screen owns yet. */
  function stripReserved(tableName, patch) {
    var clean = {};
    Object.keys(patch || {}).forEach(function (key) {
      var column = Schema.getColumn(tableName, key);
      if (!column || column.reserved) return;
      clean[key] = patch[key];
    });
    return clean;
  }

  return {
    REASON_REQUIRED_FIELDS: REASON_REQUIRED_FIELDS,
    validate: validate,
    assertOneOf: assertOneOf,
    requireReason: requireReason,
    assertReasonForPatch: assertReasonForPatch,
    stripReserved: stripReserved
  };
})();
