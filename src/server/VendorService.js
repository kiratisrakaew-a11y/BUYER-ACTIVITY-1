/**
 * VendorService.js — the vendor master: the department's own register of the
 * suppliers it deals with, kept here rather than in a spreadsheet of its own.
 *
 * The duplicate rules in SPEC §6.2 exist to catch two different problems:
 * a Tax_ID already on file is a data-entry mistake and is blocked, while a shared
 * phone number, e-mail or address between two vendors is a red flag that a buyer
 * must see but may legitimately override.
 */
var VendorService = (function () {

  var VENDOR_STATUSES = ['NEW', 'APPROVED', 'BLACKLIST', 'INACTIVE'];

  /** Only an administrator may move a vendor into or out of these (SPEC §7). */
  var ADMIN_ONLY_STATUSES = ['APPROVED', 'BLACKLIST'];

  var MASTER_FIELDS = [
    'Vendor_No', 'Vendor_Name', 'Tax_ID', 'Address',
    'Contact_Name', 'Contact_Phone', 'Contact_Email', 'Categories', 'Remark'
  ];

  /* ------------------------------------------------------------ master data */

  function search(query) {
    var q = Utils.isBlank(query) ? '' : String(query).trim().toLowerCase();
    var vendors = Repository.query('Vendors', {
      where: function (v) {
        if (!q) return true;
        return [v.Vendor_Name, v.Tax_ID, v.Vendor_No, v.Categories]
          .some(function (field) { return String(field || '').toLowerCase().indexOf(q) !== -1; });
      }
    });
    vendors.sort(function (a, b) { return String(a.Vendor_Name).localeCompare(String(b.Vendor_Name), 'th'); });
    return vendors.slice(0, 200).map(Repository.toClient);
  }

  function create(user, payload) {
    Auth.requireRole(user, [Auth.ROLES.BUYER, Auth.ROLES.HEAD, Auth.ROLES.ADMIN]);

    var values = pick(payload, MASTER_FIELDS);
    values.Tax_ID = normalizeTaxId(values.Tax_ID);
    // A buyer adding a vendor mid-sourcing creates it as NEW; approval is the
    // administrator's decision, not the buyer's (SPEC §7).
    values.Vendor_Status = 'NEW';
    Validation.validate('Vendors', values, { partial: false });
    assertTaxIdFormat(values.Tax_ID);

    var clash = findByTaxId(values.Tax_ID);
    if (clash) {
      throw Err.duplicate('เลขประจำตัวผู้เสียภาษีนี้มีอยู่แล้วในระบบ: ' + clash.Vendor_Name, {
        existing: Repository.toClient(clash)
      });
    }

    var warnings = contactCollisionWarnings(values, null);
    var created = Repository.insert('Vendors', values, { actor: user.email });
    return { vendor: Repository.toClient(created), warnings: warnings };
  }

  function update(user, vendorId, patch, version, reason) {
    Auth.requireRole(user, [Auth.ROLES.BUYER, Auth.ROLES.HEAD, Auth.ROLES.ADMIN]);
    var existing = Repository.requireById('Vendors', vendorId);

    var values = pick(patch, MASTER_FIELDS);
    if (Object.prototype.hasOwnProperty.call(patch, 'Vendor_Status')) {
      values.Vendor_Status = assertStatusChangeAllowed(user, existing, patch.Vendor_Status);
    }
    if (Object.prototype.hasOwnProperty.call(values, 'Tax_ID')) {
      values.Tax_ID = normalizeTaxId(values.Tax_ID);
      assertTaxIdFormat(values.Tax_ID);
      var clash = findByTaxId(values.Tax_ID);
      if (clash && clash.Vendor_ID !== vendorId) {
        throw Err.duplicate('เลขประจำตัวผู้เสียภาษีนี้เป็นของ ' + clash.Vendor_Name + ' อยู่แล้ว', {
          existing: Repository.toClient(clash)
        });
      }
    }
    Validation.validate('Vendors', values, { partial: true, existing: existing });

    var warnings = contactCollisionWarnings(Object.assign({}, existing, values), vendorId);
    var updated = Repository.update('Vendors', vendorId, values, version, {
      actor: user.email,
      reason: Utils.isBlank(reason) ? '' : String(reason).trim(),
      fieldActions: { Vendor_Status: ChangeLog.ACTIONS.STATUS_CHANGE }
    });
    return { vendor: Repository.toClient(updated), warnings: warnings };
  }

  function assertStatusChangeAllowed(user, existing, nextStatus) {
    Validation.assertOneOf('Vendor_Status', nextStatus, VENDOR_STATUSES);
    if (nextStatus === existing.Vendor_Status) return nextStatus;
    var touchesRestricted = ADMIN_ONLY_STATUSES.indexOf(nextStatus) !== -1 ||
      ADMIN_ONLY_STATUSES.indexOf(existing.Vendor_Status) !== -1;
    if (touchesRestricted && !Auth.isAdmin(user)) {
      throw Err.forbidden('เฉพาะผู้ดูแลระบบเท่านั้นที่เปลี่ยนสถานะผู้ขายเป็น APPROVED หรือ BLACKLIST ได้',
        { role: user.role });
    }
    return nextStatus;
  }

  function normalizeTaxId(value) {
    return Utils.isBlank(value) ? '' : String(value).replace(/[\s-]/g, '');
  }

  function assertTaxIdFormat(taxId) {
    if (!/^\d{13}$/.test(taxId)) {
      throw Err.validation('เลขประจำตัวผู้เสียภาษีต้องเป็นตัวเลข 13 หลัก', { field: 'Tax_ID' });
    }
  }

  function findByTaxId(taxId) {
    if (Utils.isBlank(taxId)) return null;
    var matches = Repository.query('Vendors', {
      where: function (v) { return normalizeTaxId(v.Tax_ID) === taxId; }
    });
    return matches.length ? matches[0] : null;
  }

  /**
   * SPEC §6.2 — a phone, e-mail or address shared with another vendor is the
   * classic sign of related bidders. Warn loudly, but let the buyer proceed.
   */
  function contactCollisionWarnings(values, selfVendorId) {
    var checks = [
      { field: 'Contact_Phone', label: 'เบอร์โทรศัพท์' },
      { field: 'Contact_Email', label: 'อีเมลผู้ติดต่อ' },
      { field: 'Address', label: 'ที่อยู่' }
    ];
    var warnings = [];
    var all = Repository.query('Vendors');

    checks.forEach(function (check) {
      var value = Utils.normalizeForCompare(values[check.field]);
      if (!value) return;
      var others = all.filter(function (v) {
        return v.Vendor_ID !== selfVendorId && Utils.normalizeForCompare(v[check.field]) === value;
      });
      if (others.length === 0) return;
      warnings.push(check.label + 'ตรงกับผู้ขายรายอื่น: ' +
        others.map(function (v) { return v.Vendor_Name; }).join(', ') +
        ' — โปรดตรวจสอบความเกี่ยวข้องกันก่อนเปรียบเทียบราคา');
    });
    return warnings;
  }

  function pick(source, fields) {
    var out = {};
    fields.forEach(function (field) {
      if (source && Object.prototype.hasOwnProperty.call(source, field)) out[field] = source[field];
    });
    return out;
  }

  return {
    VENDOR_STATUSES: VENDOR_STATUSES,
    MASTER_FIELDS: MASTER_FIELDS,
    search: search,
    create: create,
    update: update,
    findByTaxId: findByTaxId,
    normalizeTaxId: normalizeTaxId,
    contactCollisionWarnings: contactCollisionWarnings
  };
})();
