/**
 * ChangeLog.js — the audit trail (SPEC §5.4).
 *
 * Append-only by construction: this file offers no update and no delete, and no
 * api_* function is allowed to reach the Change_Log sheet by any other route.
 * Repository writes every entry, so a service cannot forget to log a change.
 */
var ChangeLog = (function () {

  var SHEET = 'Change_Log';

  /** SPEC §5.4 — the closed set of actions an auditor will see. */
  var ACTIONS = {
    CREATE: 'CREATE',
    UPDATE: 'UPDATE',
    DELETE: 'DELETE',
    RESTORE: 'RESTORE',
    STATUS_CHANGE: 'STATUS_CHANGE',
    REASSIGN: 'REASSIGN',
    REOPEN: 'REOPEN',
    EXCEPTION: 'EXCEPTION'
  };

  /** Used when the system itself changes data, e.g. an auto-revert (SPEC §6.1). */
  var SYSTEM_USER = 'SYSTEM';

  var MAX_VALUE_LENGTH = 2000;            // keeps a runaway text field out of the sheet

  /**
   * Appends entries in one write.
   * entry = { user, table, recordId, caseId, action, field, oldValue, newValue, reason }
   */
  function write(entries) {
    if (!entries || entries.length === 0) return [];

    var ids = IdGenerator.reserve('Change_Log', entries.length);
    var timestamp = Utils.now();
    var rows = entries.map(function (e, i) {
      return [
        ids[i],
        timestamp,
        e.user || SYSTEM_USER,
        e.table || '',
        e.recordId || '',
        e.caseId || '',
        e.action,
        e.field || '',
        Utils.truncate(e.oldValue === undefined || e.oldValue === null ? '' : e.oldValue, MAX_VALUE_LENGTH),
        Utils.truncate(e.newValue === undefined || e.newValue === null ? '' : e.newValue, MAX_VALUE_LENGTH),
        Utils.truncate(e.reason === undefined || e.reason === null ? '' : e.reason, MAX_VALUE_LENGTH)
      ];
    });

    var sheet = Config.getSheet(SHEET);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    // Anything holding a cached read of Change_Log is now stale.
    if (typeof Repository !== 'undefined') Repository.resetCache(SHEET);
    return ids;
  }

  /** Convenience for a single entry. */
  function writeOne(entry) {
    return write([entry])[0];
  }

  function toRecord(row) {
    return {
      Log_ID: row[0],
      Timestamp: row[1],
      User: row[2],
      Table_Name: row[3],
      Record_ID: row[4],
      Case_ID: row[5],
      Action: row[6],
      Field: row[7],
      Old_Value: row[8],
      New_Value: row[9],
      Reason: row[10]
    };
  }

  /** Every entry for a Case, newest first. Backs the "ประวัติการแก้ไข" tab. */
  function forCase(caseId) {
    var sheet = Config.getSheet(SHEET);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var matches = sheet.getRange(2, 6, lastRow - 1, 1)
      .createTextFinder(caseId).matchEntireCell(true).findAll()
      .map(function (r) { return r.getRow(); });
    if (matches.length === 0) return [];

    var entries = readRows(sheet, matches).map(toRecord);
    entries.sort(function (a, b) {
      return Utils.toDate(b.Timestamp).getTime() - Utils.toDate(a.Timestamp).getTime();
    });
    return entries;
  }

  /**
   * The status a Case held before it moved to `toStatus`.
   * Reopen uses this to restore the status a Case had before it was closed (SPEC §6.3).
   */
  function previousStatusBefore(caseId, toStatus) {
    var entries = forCase(caseId).filter(function (e) {
      return e.Action === ACTIONS.STATUS_CHANGE && e.New_Value === toStatus;
    });
    if (entries.length === 0) return null;
    return entries[0].Old_Value || null;      // forCase is newest first
  }

  /** Reads scattered row numbers in as few calls as possible by grouping runs. */
  function readRows(sheet, rowNumbers) {
    var width = sheet.getLastColumn();
    var out = [];
    Utils.groupRuns(rowNumbers).forEach(function (run) {
      sheet.getRange(run[0], 1, run[1] - run[0] + 1, width).getValues()
        .forEach(function (v) { out.push(v); });
    });
    return out;
  }

  return {
    SHEET: SHEET,
    ACTIONS: ACTIONS,
    SYSTEM_USER: SYSTEM_USER,
    write: write,
    writeOne: writeOne,
    forCase: forCase,
    previousStatusBefore: previousStatusBefore
  };
})();
