/**
 * Api.js — every function the client may call, and nothing else.
 *
 * Contract (SPEC §9):
 *   success -> { ok: true,  data }
 *   failure -> { ok: false, error: { code, message, details } }
 * Non-blocking warnings travel inside data.warnings (SPEC §6.2).
 *
 * handle() resolves the caller, enforces the role list and converts any thrown
 * error into the envelope, so no api_* function can accidentally skip the check
 * or leak a stack trace to the browser.
 */

/** Roles are named here rather than inline so the permission table stays readable. */
var API_ROLES = {
  ANY: null,
  BUYER_HEAD: ['BUYER', 'HEAD'],
  HEAD_ONLY: ['HEAD'],
  ADMIN_ONLY: ['ADMIN'],
  VENDOR_EDITORS: ['BUYER', 'HEAD', 'ADMIN'],
  TEAM_VIEWERS: ['HEAD', 'AUDITOR', 'ADMIN']
};

function handle(name, allowedRoles, fn) {
  try {
    var user = Auth.getCurrentUser();
    if (allowedRoles) Auth.requireRole(user, allowedRoles);
    return { ok: true, data: fn(user) };
  } catch (e) {
    return toErrorResponse(name, e);
  }
}

/** Logs the detail for the developer, returns only what is safe for the user. */
function toErrorResponse(name, e) {
  if (isAppError(e)) {
    if (e.code === ERROR_CODES.INTERNAL) {
      console.error(name + ' failed: ' + e.message + '\n' + (e.stack || ''));
    }
    return { ok: false, error: { code: e.code, message: e.message, details: e.details } };
  }
  console.error(name + ' failed unexpectedly: ' + ((e && e.message) || e) + '\n' + ((e && e.stack) || ''));
  return {
    ok: false,
    error: { code: ERROR_CODES.INTERNAL, message: 'เกิดข้อผิดพลาดภายในระบบ กรุณาลองใหม่อีกครั้ง', details: null }
  };
}

/* ------------------------------------------------------------------ bootstrap */

/**
 * Everything the single-page client needs before it can render anything:
 * the current user, their permissions, all dropdown lists and the status graph.
 */
function api_bootstrap() {
  return handle('api_bootstrap', API_ROLES.ANY, function (user) {
    return {
      user: { email: user.email, name: user.name, role: user.role, scope: user.scope },
      permissions: Auth.permissions(user),
      lists: Config.getAllLists(),
      statuses: Config.getStatusMaster(),
      settings: {
        BUYER_CAN_VIEW_ALL: Config.getBool('BUYER_CAN_VIEW_ALL', true),
        APP_TIMEZONE: Config.getTimezone()
      },
      buyers: Auth.listBuyers()
    };
  });
}

/** Admins clear the config cache after editing the Config sheets (SPEC §11). */
function api_clearCache() {
  return handle('api_clearCache', API_ROLES.ADMIN_ONLY, function () {
    Config.clearCache();
    Repository.resetCache();
    return { cleared: true };
  });
}

/* ----------------------------------------------------------------- cases */

function api_listCases(filter) {
  return handle('api_listCases', API_ROLES.ANY, function (user) {
    return CaseService.list(user, filter || {});
  });
}

function api_getCase(caseId) {
  return handle('api_getCase', API_ROLES.ANY, function (user) {
    return CaseService.getBundle(user, caseId);
  });
}

function api_createCase(payload) {
  return handle('api_createCase', API_ROLES.BUYER_HEAD, function (user) {
    return CaseService.create(user, payload || {});
  });
}

function api_updateCase(caseId, patch, version, reason) {
  return handle('api_updateCase', API_ROLES.BUYER_HEAD, function (user) {
    return CaseService.update(user, caseId, patch || {}, version, reason);
  });
}

/* ---------------------------------------------------------------- delete */

/**
 * One delete entry point for every table (SPEC §9). The handler for a table owns
 * its cascade, so a caller cannot delete a parent and orphan its children.
 * A reason is always required (SPEC §6.3).
 */
var DELETE_HANDLERS = {
  Activities: function (user, record, version, reason) {
    return ActivityService.remove(user, record.Activity_ID, version, reason);
  }
};

function api_deleteRecord(tableName, recordId, version, reason) {
  return handle('api_deleteRecord', API_ROLES.BUYER_HEAD, function (user) {
    var deleter = DELETE_HANDLERS[tableName];
    if (!deleter) {
      throw Err.validation('ไม่รองรับการลบข้อมูลในตาราง ' + tableName, { table: tableName });
    }
    var record = Repository.requireById(tableName, recordId);
    return deleter(user, record, version, reason);
  });
}

/* --------------------------------------------------------------- vendors */

function api_searchVendors(query) {
  return handle('api_searchVendors', API_ROLES.ANY, function () {
    return { vendors: VendorService.search(query) };
  });
}

function api_createVendor(payload) {
  return handle('api_createVendor', API_ROLES.VENDOR_EDITORS, function (user) {
    return VendorService.create(user, payload || {});
  });
}

function api_updateVendor(vendorId, patch, version, reason) {
  return handle('api_updateVendor', API_ROLES.VENDOR_EDITORS, function (user) {
    return VendorService.update(user, vendorId, patch || {}, version, reason);
  });
}

/* --------------------------------------------------------------- uploads */

function api_uploadFile(caseId, fileName, mimeType, base64) {
  return handle('api_uploadFile', API_ROLES.BUYER_HEAD, function (user) {
    var caseRecord = CaseService.getForEdit(user, caseId);
    return DriveService.uploadToCase(caseRecord, fileName, mimeType, base64);
  });
}

/* ------------------------------------------------------------- activities */

function api_saveActivity(caseId, activity, version) {
  return handle('api_saveActivity', API_ROLES.BUYER_HEAD, function (user) {
    return ActivityService.save(user, caseId, activity || {}, version);
  });
}

function api_setNextActionDone(activityId, done, version) {
  return handle('api_setNextActionDone', API_ROLES.BUYER_HEAD, function (user) {
    return ActivityService.setNextActionDone(user, activityId, done, version);
  });
}

/* ---------------------------------------------------- status and workflow */

function api_changeStatus(caseId, toStatus, version, reason) {
  return handle('api_changeStatus', API_ROLES.BUYER_HEAD, function (user) {
    var result = StatusEngine.transition(user, caseId, toStatus, version, reason);
    return Object.assign(result, { warnings: Rules.recheckCaseRules(caseId).messages });
  });
}

/**
 * Closing carries its own endpoint rather than going through api_changeStatus,
 * because the PR number and the buying company are part of the decision and
 * have to be written in the same versioned update as the status.
 */
function api_closeCase(caseId, payload) {
  return handle('api_closeCase', API_ROLES.BUYER_HEAD, function (user) {
    return CaseWorkflow.closeCase(user, caseId, payload || {});
  });
}

function api_reassignCase(caseId, newOwnerEmail, reason) {
  return handle('api_reassignCase', API_ROLES.HEAD_ONLY, function (user) {
    return CaseWorkflow.reassign(user, caseId, newOwnerEmail, reason);
  });
}

function api_reopenCase(caseId, reason) {
  return handle('api_reopenCase', API_ROLES.HEAD_ONLY, function (user) {
    return CaseWorkflow.reopen(user, caseId, reason);
  });
}

/* -------------------------------------------------------------- audit log */

function api_getChangeLog(caseId) {
  return handle('api_getChangeLog', API_ROLES.ANY, function (user) {
    CaseService.getForView(user, caseId);        // you may read the log of a Case you may read
    return { entries: ChangeLog.forCase(caseId).map(toLogClient) };
  });
}

function toLogClient(entry) {
  return {
    Log_ID: entry.Log_ID,
    Timestamp: entry.Timestamp instanceof Date ? entry.Timestamp.toISOString() : String(entry.Timestamp),
    User: entry.User,
    Table_Name: entry.Table_Name,
    Record_ID: entry.Record_ID,
    Action: entry.Action,
    Field: entry.Field,
    Old_Value: entry.Old_Value,
    New_Value: entry.New_Value,
    Reason: entry.Reason
  };
}

/* -------------------------------------------------------------- team view */

function api_teamView() {
  return handle('api_teamView', API_ROLES.TEAM_VIEWERS, function (user) {
    return TeamService.overview(user);
  });
}

/** The post-deployment checklist, also reachable from the app for an administrator. */
function api_verifyDeployment() {
  return handle('api_verifyDeployment', API_ROLES.ADMIN_ONLY, function () {
    return Verify.run();
  });
}

/** Lets an administrator run the daily job by hand, e.g. after fixing a trigger. */
function api_runDailyJob() {
  return handle('api_runDailyJob', API_ROLES.ADMIN_ONLY, function () {
    return dailyReminderJob();
  });
}
