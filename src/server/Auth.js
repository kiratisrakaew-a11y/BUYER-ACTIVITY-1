/**
 * Auth.js — who is calling, and what they are allowed to do (SPEC §7).
 *
 * The web app is deployed "execute as me", so the browser identity is the only
 * thing that distinguishes users. Every api_* function therefore resolves the
 * caller here and checks permission on the server; hiding a button in the client
 * is a convenience, never a control (SPEC §7, §11).
 */
var Auth = (function () {

  var ROLES = { BUYER: 'BUYER', HEAD: 'HEAD', AUDITOR: 'AUDITOR', ADMIN: 'ADMIN' };

  /** Roles that may act on any Case rather than only their own. */
  var ALL_CASE_EDITORS = [ROLES.HEAD];
  var ALL_CASE_VIEWERS = [ROLES.HEAD, ROLES.AUDITOR, ROLES.ADMIN];

  var userOverride = null;                // set by Tests.js only; no api_* exposes it

  /** Tests call this to act as a given user. Not reachable from the client. */
  function __setUserOverride(email) {
    userOverride = email || null;
  }

  function getActiveEmail() {
    if (userOverride) return userOverride;
    var email = '';
    try {
      email = Session.getActiveUser().getEmail();
    } catch (e) {
      email = '';
    }
    return String(email || '').trim().toLowerCase();
  }

  /**
   * The caller as a Users row. Throws UNAUTHORIZED when they are absent from the
   * sheet or marked inactive, which is what produces the "ไม่มีสิทธิ์เข้าใช้งาน" screen.
   */
  function getCurrentUser() {
    var email = getActiveEmail();
    if (!email) {
      throw Err.unauthorized('ระบบไม่สามารถระบุตัวตนของคุณได้ กรุณาเข้าสู่ระบบด้วยบัญชีของบริษัท');
    }
    var row = findUser(email);
    if (!row) {
      throw Err.unauthorized('ไม่พบบัญชี ' + email + ' ในระบบ กรุณาติดต่อผู้ดูแลระบบ');
    }
    if (row.Is_Active !== true) {
      throw Err.unauthorized('บัญชี ' + email + ' ถูกปิดการใช้งาน กรุณาติดต่อผู้ดูแลระบบ');
    }
    return {
      email: email,
      name: row.Name || email,
      role: String(row.Role || '').trim().toUpperCase(),
      scope: row.Responsible_Scope || '',
      isActive: true
    };
  }

  function findUser(email) {
    var wanted = String(email).trim().toLowerCase();
    var users = Repository.readAll('Users');
    for (var i = 0; i < users.length; i++) {
      if (String(users[i].Email || '').trim().toLowerCase() === wanted) return users[i];
    }
    return null;
  }

  function hasRole(user, roles) {
    if (!roles || roles.length === 0) return true;
    return roles.indexOf(user.role) !== -1;
  }

  function requireRole(user, roles) {
    if (!hasRole(user, roles)) {
      throw Err.forbidden('บทบาท ' + user.role + ' ไม่มีสิทธิ์ทำรายการนี้',
        { role: user.role, allowed: roles });
    }
    return user;
  }

  function isHead(user) { return user.role === ROLES.HEAD; }
  function isAdmin(user) { return user.role === ROLES.ADMIN; }
  function isOwner(user, caseRow) {
    return !!caseRow && String(caseRow.Buyer_Owner || '').trim().toLowerCase() === user.email;
  }

  /** SPEC §7 — BUYER sees other people's Cases read-only when BUYER_CAN_VIEW_ALL is on. */
  function canViewCase(user, caseRow) {
    if (ALL_CASE_VIEWERS.indexOf(user.role) !== -1) return true;
    if (isOwner(user, caseRow)) return true;
    if (user.role === ROLES.BUYER) return Config.getBool('BUYER_CAN_VIEW_ALL', true);
    return false;
  }

  function assertCanViewCase(user, caseRow) {
    if (!canViewCase(user, caseRow)) {
      throw Err.forbidden('คุณไม่มีสิทธิ์ดูงานนี้', { caseId: caseRow && caseRow.Case_ID });
    }
    return caseRow;
  }

  /**
   * SPEC §6.3 / §7 — data is editable until the Case reaches a terminal status,
   * and then only HEAD can bring it back through Reopen.
   */
  function canEditCase(user, caseRow) {
    if (!caseRow) return false;
    if (isTerminal(caseRow.Status)) return false;
    if (ALL_CASE_EDITORS.indexOf(user.role) !== -1) return true;
    return user.role === ROLES.BUYER && isOwner(user, caseRow);
  }

  function assertCanEditCase(user, caseRow) {
    if (!caseRow) throw Err.notFound('ไม่พบงานที่ต้องการแก้ไข');
    if (isTerminal(caseRow.Status)) {
      throw Err.forbidden('งานนี้อยู่ในสถานะ ' + statusLabel(caseRow.Status) + ' จึงแก้ไขไม่ได้ ' +
        '(หัวหน้าฝ่ายจัดซื้อสามารถกด Reopen พร้อมระบุเหตุผลได้)', { caseId: caseRow.Case_ID });
    }
    if (!canEditCase(user, caseRow)) {
      throw Err.forbidden('คุณไม่ใช่เจ้าของงานนี้ จึงแก้ไขข้อมูลไม่ได้', {
        caseId: caseRow.Case_ID,
        owner: caseRow.Buyer_Owner
      });
    }
    return caseRow;
  }

  /**
   * SPEC §7 — recording an Activity on someone else's Case is allowed, because
   * buyers cover for each other. Performed_By still records who actually did it.
   */
  function canLogActivity(user, caseRow) {
    if (!caseRow || isTerminal(caseRow.Status)) return false;
    return user.role === ROLES.BUYER || user.role === ROLES.HEAD;
  }

  function assertCanLogActivity(user, caseRow) {
    if (!caseRow) throw Err.notFound('ไม่พบงานที่ต้องการบันทึกกิจกรรม');
    if (isTerminal(caseRow.Status)) {
      throw Err.forbidden('งานนี้ปิดแล้ว จึงบันทึกกิจกรรมเพิ่มไม่ได้', { caseId: caseRow.Case_ID });
    }
    if (!canLogActivity(user, caseRow)) {
      throw Err.forbidden('บทบาท ' + user.role + ' ไม่มีสิทธิ์บันทึกกิจกรรม', { role: user.role });
    }
    return caseRow;
  }

  function isTerminal(statusCode) {
    var status = Config.getStatus(statusCode);
    return !!status && status.isTerminal;
  }

  function statusLabel(statusCode) {
    var status = Config.getStatus(statusCode);
    return status ? status.label : statusCode;
  }

  /** What the client uses to decide which buttons to render. Never a substitute for the checks above. */
  function permissions(user) {
    return {
      canCreateCase: user.role === ROLES.BUYER || user.role === ROLES.HEAD,
      canReassign: user.role === ROLES.HEAD,
      canReopen: user.role === ROLES.HEAD,
      // The owning buyer closes their own Case: they are the one holding the
      // PR number EPICOR issued. StatusEngine still checks ownership per Case.
      canCloseCase: user.role === ROLES.BUYER || user.role === ROLES.HEAD,
      canCreateVendor: [ROLES.BUYER, ROLES.HEAD, ROLES.ADMIN].indexOf(user.role) !== -1,
      canSetVendorApproval: user.role === ROLES.ADMIN,
      canSeeTeamView: ALL_CASE_VIEWERS.indexOf(user.role) !== -1,
      canEditAnyCase: ALL_CASE_EDITORS.indexOf(user.role) !== -1,
      canViewAllCases: ALL_CASE_VIEWERS.indexOf(user.role) !== -1 || Config.getBool('BUYER_CAN_VIEW_ALL', true)
    };
  }

  /** All active users, for the reassign picker. */
  function listActiveUsers() {
    return Repository.readAll('Users')
      .filter(function (u) { return u.Is_Active === true; })
      .map(function (u) {
        return {
          email: String(u.Email).trim().toLowerCase(),
          name: u.Name || u.Email,
          role: String(u.Role || '').trim().toUpperCase()
        };
      });
  }

  function listBuyers() {
    return listActiveUsers().filter(function (u) {
      return u.role === ROLES.BUYER || u.role === ROLES.HEAD;
    });
  }

  return {
    ROLES: ROLES,
    __setUserOverride: __setUserOverride,
    getActiveEmail: getActiveEmail,
    getCurrentUser: getCurrentUser,
    findUser: findUser,
    hasRole: hasRole,
    requireRole: requireRole,
    isHead: isHead,
    isAdmin: isAdmin,
    isOwner: isOwner,
    isTerminal: isTerminal,
    statusLabel: statusLabel,
    canViewCase: canViewCase,
    assertCanViewCase: assertCanViewCase,
    canEditCase: canEditCase,
    assertCanEditCase: assertCanEditCase,
    canLogActivity: canLogActivity,
    assertCanLogActivity: assertCanLogActivity,
    permissions: permissions,
    listActiveUsers: listActiveUsers,
    listBuyers: listBuyers
  };
})();
