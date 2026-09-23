/**
 * Tests.js — the one acceptance-test suite, shared by both runners.
 *
 *   Node:         node tests/node/run.js
 *   Apps Script:  open the editor and run runAllTests()
 *
 * Numbers in test names map to the Acceptance Test Cases in SPEC §13.
 * Each test gets a brand-new DB spreadsheet built by the real setup(), so the
 * tests exercise the same code path an administrator does on day one.
 */

/* ------------------------------------------------------------- mini framework */

var TEST_REGISTRY = [];

function test(name, fn) {
  TEST_REGISTRY.push({ name: name, fn: fn });
}

function assert(condition, message) {
  if (!condition) throw new Error('assert failed: ' + (message || ''));
}

function assertEquals(actual, expected, message) {
  var a = actual instanceof Date ? actual.toISOString() : actual;
  var e = expected instanceof Date ? expected.toISOString() : expected;
  if (a !== e) {
    throw new Error((message || 'values differ') + ' — expected ' + JSON.stringify(e) + ', got ' + JSON.stringify(a));
  }
}

function assertDeepEquals(actual, expected, message) {
  var a = JSON.stringify(actual);
  var e = JSON.stringify(expected);
  if (a !== e) throw new Error((message || 'values differ') + ' — expected ' + e + ', got ' + a);
}

function assertContains(haystack, needle, message) {
  if (String(haystack).indexOf(needle) === -1) {
    throw new Error((message || 'missing substring') + ' — ' + JSON.stringify(needle) + ' not in ' + JSON.stringify(String(haystack)));
  }
}

/** Asserts that fn() throws an AppError carrying the given code. */
function assertThrowsCode(expectedCode, fn, message) {
  try {
    fn();
  } catch (e) {
    if (e && e.code === expectedCode) return e;
    throw new Error((message || 'wrong error') + ' — expected ' + expectedCode + ', got ' +
      (e && e.code ? e.code : (e && e.message) || String(e)));
  }
  throw new Error((message || 'expected an error') + ' — expected ' + expectedCode + ', nothing was thrown');
}

/** Asserts that an api_* envelope failed with the given code. */
function assertApiError(response, expectedCode, message) {
  assert(response && response.ok === false, (message || 'expected failure') + ' — got ' + JSON.stringify(response));
  assertEquals(response.error.code, expectedCode, message || 'error code');
  return response.error;
}

/** Asserts that an api_* envelope succeeded, and returns its data. */
function assertApiOk(response, message) {
  if (!response || response.ok !== true) {
    throw new Error((message || 'expected success') + ' — got ' + JSON.stringify(response));
  }
  return response.data;
}

/* ------------------------------------------------------------------ fixtures */

/**
 * Builds a fresh database with the real setup(), runs fn(report), then cleans up.
 * Deleting the script property first guarantees setup() creates a new spreadsheet
 * rather than reusing the previous test's.
 */
function withFreshDatabase(fn) {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(Config.DB_PROPERTY_KEY);
  Config.__clearDbOverride();
  Config.clearCache();

  // Triggers and sent mail live outside the spreadsheet, so a fresh database is
  // not by itself a fresh world. Clearing triggers is safe only under the mock:
  // on Apps Script the project's real reminder trigger belongs to production.
  if (typeof __test !== 'undefined') {
    __test.clearMail();
    ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  }

  var report = setup();
  try {
    return fn(report);
  } finally {
    Config.clearCache();
    Config.__clearDbOverride();
    try {
      DriveApp.getFileById(report.spreadsheetId).setTrashed(true);
    } catch (e) {
      // A scratch spreadsheet that cannot be trashed is not a test failure.
    }
    props.deleteProperty(Config.DB_PROPERTY_KEY);
  }
}

/* ----------------------------------------------------------------- the runner */

function runAllTests(options) {
  var opts = options || {};
  var selected = TEST_REGISTRY.filter(function (t) {
    if (opts.caseNumber) {
      var prefix = String(opts.caseNumber).trim();
      if (t.name.indexOf('T' + prefix + ' ') !== 0 && t.name.indexOf('T' + prefix + ':') !== 0) return false;
    }
    if (opts.grep && t.name.toLowerCase().indexOf(String(opts.grep).toLowerCase()) === -1) return false;
    return true;
  });

  var results = [];
  var passed = 0;
  var failed = 0;

  selected.forEach(function (t) {
    var started = new Date().getTime();
    try {
      t.fn();
      passed++;
      results.push({ name: t.name, ok: true, ms: new Date().getTime() - started });
    } catch (e) {
      failed++;
      results.push({
        name: t.name,
        ok: false,
        ms: new Date().getTime() - started,
        error: (e && e.code ? '[' + e.code + '] ' : '') + ((e && e.message) || String(e)),
        stack: e && e.stack ? e.stack : ''
      });
    }
  });

  var summary = { passed: passed, failed: failed, results: results };
  console.log('Tests: ' + passed + ' passed, ' + failed + ' failed');
  results.forEach(function (r) {
    if (!r.ok) console.error('FAIL ' + r.name + ': ' + r.error);
  });
  return summary;
}

/* ============================================================================
 * Phase 1 — Schema, Setup, Counters, IdGenerator
 * ==========================================================================*/

test('setup creates every sheet declared in Schema', function () {
  withFreshDatabase(function () {
    var ss = Config.getDb();
    Schema.tableNames().forEach(function (tableName) {
      var sheetName = Schema.getTable(tableName).sheet;
      assert(!!ss.getSheetByName(sheetName), 'missing sheet ' + sheetName);
    });
  });
});

test('setup writes the exact schema headers into row 1', function () {
  withFreshDatabase(function () {
    Schema.tableNames().forEach(function (tableName) {
      var expected = Schema.getColumnNames(tableName);
      var sheet = Config.getSheet(Schema.getTable(tableName).sheet);
      var actual = sheet.getRange(1, 1, 1, expected.length).getValues()[0];
      assertDeepEquals(actual, expected, 'headers of ' + tableName);
    });
  });
});

test('setup seeds Config_Settings, Config_Lists and Status_Master', function () {
  withFreshDatabase(function () {
    assertEquals(Config.getBool('BUYER_CAN_VIEW_ALL'), true, 'BUYER_CAN_VIEW_ALL default');
    assertEquals(Config.getTimezone(), 'Asia/Bangkok', 'timezone');

    assertEquals(Config.getList('BUDGET_TYPE').length, 2, 'BUDGET_TYPE entries');
    assertEquals(Config.getList('SUB_TYPE', 'CAPEX').length, 3, 'SUB_TYPE under CAPEX');
    assertEquals(Config.getList('SUB_TYPE', 'OPEX').length, 1, 'SUB_TYPE under OPEX');
    assert(Config.isValidCode('ACTIVITY_TYPE', 'EPICOR_PR'), 'EPICOR_PR is a valid activity type');
    assert(Config.isValidCode('COMPANY', 'PLANB'), 'the buying company list is seeded');

    var statuses = Config.getStatusMaster();
    assertEquals(statuses.length, 3, 'open, closed and cancelled');
    assertDeepEquals(Config.getStatus('OPEN').allowedNext,
      ['CLOSED', 'CANCELLED'], 'where an open Case may go');
    assertEquals(Config.getStatus('CLOSED').isTerminal, true, 'CLOSED is terminal');
    assertEquals(Config.getStatus('CANCELLED').isTerminal, true, 'CANCELLED is terminal');
  });
});

test('T18 rerunning setup appends new schema columns and keeps existing data', function () {
  withFreshDatabase(function () {
    // Put a row in Cases so we can prove the data survives.
    var sheet = Config.getSheet('Cases');
    var headers = Schema.getColumnNames('Cases');
    var row = headers.map(function (h) { return h === 'Case_ID' ? 'SRC-2026-0001' : 'keep-' + h; });
    sheet.getRange(2, 1, 1, headers.length).setValues([row]);

    var listRowsBefore = Config.getSheet('Config_Lists').getLastRow();
    var settingRowsBefore = Config.getSheet('Config_Settings').getLastRow();

    // Simulate a future module adding a column to an existing table.
    var casesTable = Schema.getTable('Cases');
    casesTable.columns.push({ name: 'M2_Benchmark_Done', type: 'bool', reserved: true });
    try {
      var report = setup();

      assertEquals(report.sheetsCreated.length, 0, 'no sheet recreated on a rerun');
      assert(report.columnsAdded.indexOf('Cases.M2_Benchmark_Done') !== -1, 'new column reported');

      var after = Config.getSheet('Cases');
      var afterHeaders = after.getRange(1, 1, 1, after.getLastColumn()).getValues()[0];
      assertEquals(afterHeaders[afterHeaders.length - 1], 'M2_Benchmark_Done', 'new column appended last');

      // Every original header kept its original position.
      headers.forEach(function (h, i) {
        assertEquals(afterHeaders[i], h, 'header ' + h + ' stayed at column ' + (i + 1));
      });
      // Every original value is still there.
      var afterRow = after.getRange(2, 1, 1, headers.length).getValues()[0];
      assertDeepEquals(afterRow, row, 'existing Cases row untouched');

      // Seed data was not duplicated.
      assertEquals(Config.getSheet('Config_Lists').getLastRow(), listRowsBefore, 'lists not re-seeded');
      assertEquals(Config.getSheet('Config_Settings').getLastRow(), settingRowsBefore, 'settings not re-seeded');
    } finally {
      casesTable.columns.pop();
    }
  });
});

test('setup does not overwrite an administrator edit to Config_Settings', function () {
  withFreshDatabase(function () {
    var sheet = Config.getSheet('Config_Settings');
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] === 'REMINDER_HOUR') sheet.getRange(i + 2, 2).setValue('5');
    }
    Config.clearCache();
    assertEquals(Config.getNumber('REMINDER_HOUR'), 5, 'admin value read back');

    setup();
    Config.clearCache();
    assertEquals(Config.getNumber('REMINDER_HOUR'), 5, 'admin value survives a rerun');
  });
});

test('IdGenerator mints sequential, correctly formatted ids', function () {
  withFreshDatabase(function () {
    var year = IdGenerator.currentYear();
    assertEquals(IdGenerator.next('Cases'), 'SRC-' + year + '-0001', 'first Case id');
    assertEquals(IdGenerator.next('Cases'), 'SRC-' + year + '-0002', 'second Case id');
    assertEquals(IdGenerator.next('Activities'), 'ACT-000001', 'activity id has no year');
    assertEquals(IdGenerator.next('Vendors'), 'VEN-00001', 'vendor id is padded to 5');

    var batch = IdGenerator.reserve('Vendors', 3);
    assertDeepEquals(batch, ['VEN-00002', 'VEN-00003', 'VEN-00004'], 'batch reservation');
    assertEquals(IdGenerator.next('Vendors'), 'VEN-00005', 'counter continues after a batch');
  });
});

test('T2 concurrent id allocation never repeats a number', function () {
  withFreshDatabase(function () {
    var seen = {};
    for (var i = 0; i < 50; i++) {
      var id = IdGenerator.next('Cases');
      assert(!seen[id], 'duplicate id issued: ' + id);
      seen[id] = true;
    }
    assertEquals(IdGenerator.peek('Cases', IdGenerator.currentYear()), 50, 'counter matches issued ids');
    assert(!Utils.isLockHeld(), 'script lock released after allocation');
  });
});

/* ============================================================================
 * Phase 2 — Repository, Change_Log, optimistic locking, soft delete
 * ==========================================================================*/

/** Change_Log rows for one record, oldest first. */
function logsFor(tableName, recordId) {
  return Repository.readAll('Change_Log').filter(function (l) {
    return l.Table_Name === tableName && l.Record_ID === recordId;
  });
}

function makeVendor(overrides) {
  var payload = {
    Vendor_Name: 'บริษัท ทดสอบ จำกัด',
    Tax_ID: '0105500000001',
    Contact_Phone: '021112222',
    Contact_Email: 'sales@vendor.example',
    Vendor_Status: 'NEW'
  };
  Object.keys(overrides || {}).forEach(function (k) { payload[k] = overrides[k]; });
  return Repository.insert('Vendors', payload, { actor: 'buyer.a@example.com' });
}

test('Repository.insert fills audit columns and logs a CREATE', function () {
  withFreshDatabase(function () {
    var vendor = makeVendor();

    assertEquals(vendor.Vendor_ID, 'VEN-00001', 'generated id');
    assertEquals(vendor.Version, 1, 'version starts at 1');
    assertEquals(vendor.Is_Deleted, false, 'not deleted');
    assertEquals(vendor.Created_By, 'buyer.a@example.com', 'Created_By');
    assertEquals(vendor.Updated_By, 'buyer.a@example.com', 'Updated_By');
    assert(vendor.Created_At instanceof Date, 'Created_At is a date');

    var logs = logsFor('Vendors', 'VEN-00001');
    assertEquals(logs.length, 1, 'one CREATE entry');
    assertEquals(logs[0].Action, 'CREATE', 'action');
    assertEquals(logs[0].Field, '', 'CREATE has no field (SPEC 5.4)');
    assertContains(logs[0].New_Value, 'Tax_ID=0105500000001', 'CREATE snapshot');

    // Reading it back by id returns the same record.
    var reread = Repository.requireById('Vendors', 'VEN-00001');
    assertEquals(reread.Vendor_Name, 'บริษัท ทดสอบ จำกัด', 'round trip');
  });
});

test('Repository.insert ignores client-supplied ids and audit columns', function () {
  withFreshDatabase(function () {
    var vendor = Repository.insert('Vendors', {
      Vendor_ID: 'VEN-99999',
      Vendor_Name: 'ผู้ขายปลอม',
      Tax_ID: '0105500000002',
      Vendor_Status: 'APPROVED',
      Version: 42,
      Is_Deleted: true,
      Created_By: 'attacker@example.com'
    }, { actor: 'buyer.a@example.com' });

    assertEquals(vendor.Vendor_ID, 'VEN-00001', 'system issues the id');
    assertEquals(vendor.Version, 1, 'client Version ignored');
    assertEquals(vendor.Is_Deleted, false, 'client Is_Deleted ignored');
    assertEquals(vendor.Created_By, 'buyer.a@example.com', 'client Created_By ignored');
  });
});

test('T12 update logs one row per changed field with old, new and reason', function () {
  withFreshDatabase(function () {
    makeVendor();
    var updated = Repository.update('Vendors', 'VEN-00001', {
      Vendor_Name: 'บริษัท ทดสอบ (แก้ไข) จำกัด',
      Contact_Phone: '029998888',
      Vendor_Status: 'NEW'                      // unchanged — must not be logged
    }, 1, { actor: 'head@example.com', reason: 'แก้ชื่อตามหนังสือรับรอง' });

    assertEquals(updated.Version, 2, 'version bumped once for the whole update');
    assertEquals(updated.Updated_By, 'head@example.com', 'Updated_By');

    var changes = logsFor('Vendors', 'VEN-00001').filter(function (l) { return l.Action === 'UPDATE'; });
    assertEquals(changes.length, 2, 'only the two fields that actually changed');

    var byField = {};
    changes.forEach(function (c) { byField[c.Field] = c; });
    assert(!byField.Vendor_Status, 'an unchanged field is not logged');
    assertEquals(byField.Vendor_Name.Old_Value, 'บริษัท ทดสอบ จำกัด', 'old value');
    assertEquals(byField.Vendor_Name.New_Value, 'บริษัท ทดสอบ (แก้ไข) จำกัด', 'new value');
    assertEquals(byField.Vendor_Name.Reason, 'แก้ชื่อตามหนังสือรับรอง', 'reason');
    assertEquals(byField.Contact_Phone.New_Value, '029998888', 'second field');

    // Audit columns never appear as their own log rows.
    changes.forEach(function (c) {
      assert(!Schema.isAuditColumn(c.Field), 'audit column ' + c.Field + ' must not be logged');
    });
  });
});

test('T11 a stale version is rejected with CONFLICT', function () {
  withFreshDatabase(function () {
    makeVendor();

    // Two users opened the same form; both hold version 1.
    Repository.update('Vendors', 'VEN-00001', { Vendor_Name: 'บันทึกโดยคนแรก' }, 1,
      { actor: 'user1@example.com' });

    var error = assertThrowsCode('CONFLICT', function () {
      Repository.update('Vendors', 'VEN-00001', { Vendor_Name: 'บันทึกโดยคนที่สอง' }, 1,
        { actor: 'user2@example.com' });
    }, 'second save must conflict');
    assertEquals(error.details.expected, 1, 'conflict reports the version sent');
    assertEquals(error.details.actual, 2, 'conflict reports the stored version');

    assertEquals(Repository.requireById('Vendors', 'VEN-00001').Vendor_Name, 'บันทึกโดยคนแรก',
      'the losing write changed nothing');

    // Reloading and retrying with the current version succeeds.
    Repository.update('Vendors', 'VEN-00001', { Vendor_Name: 'บันทึกโดยคนที่สอง' }, 2,
      { actor: 'user2@example.com' });
    assertEquals(Repository.requireById('Vendors', 'VEN-00001').Version, 3, 'retry applied');
  });
});

test('update with no real change leaves the version alone', function () {
  withFreshDatabase(function () {
    makeVendor();
    var same = Repository.update('Vendors', 'VEN-00001', { Vendor_Name: 'บริษัท ทดสอบ จำกัด' }, 1,
      { actor: 'buyer.a@example.com' });
    assertEquals(same.Version, 1, 'version unchanged');
    assertEquals(logsFor('Vendors', 'VEN-00001').length, 1, 'only the CREATE entry exists');
  });
});

test('softDelete hides the row, keeps it in the sheet and logs DELETE', function () {
  withFreshDatabase(function () {
    makeVendor();
    var sheetRowsBefore = Config.getSheet('Vendors').getLastRow();

    Repository.softDelete('Vendors', 'VEN-00001', 1, { actor: 'admin@example.com', reason: 'สร้างซ้ำ' });

    assertEquals(Config.getSheet('Vendors').getLastRow(), sheetRowsBefore, 'no row was removed');
    assertEquals(Repository.findById('Vendors', 'VEN-00001'), null, 'hidden from normal reads');
    assert(!!Repository.findById('Vendors', 'VEN-00001', { includeDeleted: true }), 'still reachable for audit');
    assertEquals(Repository.query('Vendors').length, 0, 'excluded from queries');
    assertEquals(Repository.query('Vendors', { includeDeleted: true }).length, 1, 'included when asked');

    var del = logsFor('Vendors', 'VEN-00001').filter(function (l) { return l.Action === 'DELETE'; });
    assertEquals(del.length, 1, 'one DELETE entry');
    assertEquals(del[0].Reason, 'สร้างซ้ำ', 'reason recorded');

    Repository.restore('Vendors', 'VEN-00001', null, { actor: 'admin@example.com', reason: 'ลบผิด' });
    assert(!!Repository.findById('Vendors', 'VEN-00001'), 'restored');
    assertEquals(logsFor('Vendors', 'VEN-00001').filter(function (l) { return l.Action === 'RESTORE'; }).length,
      1, 'one RESTORE entry');
  });
});

test('fieldActions let a caller label a change STATUS_CHANGE instead of UPDATE', function () {
  withFreshDatabase(function () {
    makeVendor();
    Repository.update('Vendors', 'VEN-00001', { Vendor_Status: 'BLACKLIST' }, 1, {
      actor: 'admin@example.com',
      reason: 'พบพฤติกรรมสมยอมราคา',
      fieldActions: { Vendor_Status: ChangeLog.ACTIONS.STATUS_CHANGE }
    });
    var logs = logsFor('Vendors', 'VEN-00001');
    assertEquals(logs[logs.length - 1].Action, 'STATUS_CHANGE', 'action overridden per field');
  });
});

test('updateMany applies a batch and writes all entries at once', function () {
  withFreshDatabase(function () {
    var a = makeVendor({ Tax_ID: '0105500000011', Vendor_Name: 'ผู้ขาย A' });
    var b = makeVendor({ Tax_ID: '0105500000012', Vendor_Name: 'ผู้ขาย B' });

    var result = Repository.updateMany('Vendors', [
      { id: a.Vendor_ID, patch: { Vendor_Status: 'APPROVED' }, version: 1 },
      { id: b.Vendor_ID, patch: { Vendor_Status: 'APPROVED' }, version: 1 }
    ], { actor: 'admin@example.com', reason: 'ผ่านการตรวจสอบ' });

    assertEquals(result.length, 2, 'both returned');
    assertEquals(Repository.requireById('Vendors', a.Vendor_ID).Vendor_Status, 'APPROVED', 'first applied');
    assertEquals(Repository.requireById('Vendors', b.Vendor_ID).Version, 2, 'second version bumped');

    assertThrowsCode('CONFLICT', function () {
      Repository.updateMany('Vendors', [{ id: a.Vendor_ID, patch: { Vendor_Status: 'NEW' }, version: 1 }],
        { actor: 'admin@example.com' });
    }, 'batch honours optimistic locking too');
  });
});

test('Repository coerces types on the way in and out of the sheet', function () {
  withFreshDatabase(function () {
    var stored = Repository.insert('Activities', {
      Case_ID: 'SRC-2026-0001',
      Module: 'M1',
      Activity_Date: '2026-01-15T03:00:00.000Z',
      Activity_Type: 'CALL',
      Activity_Description: 'โทรติดตามหน่วยงาน',
      Performed_By: 'buyer.a@example.com',
      Next_Action_Done: 'TRUE'
    }, { actor: 'buyer.a@example.com' });

    assert(stored.Activity_Date instanceof Date, 'datetime coerced');
    assertEquals(stored.Next_Action_Done, true, 'bool coerced');

    var read = Repository.requireById('Activities', stored.Activity_ID);
    assertEquals(typeof read.Activity_Description, 'string', 'text reads back as a string');
    assertEquals(read.Is_Deleted, false, 'bool reads back as a boolean');

    assertThrowsCode('VALIDATION', function () {
      Repository.update('Activities', stored.Activity_ID, { Activity_Date: 'เมื่อวานนี้' }, 1,
        { actor: 'buyer.a@example.com' });
    }, 'non-date activity date rejected');
  });
});

test('a value for a column the sheet has not got yet is refused, not dropped', function () {
  withFreshDatabase(function () {
    // A column added to the schema by an update whose setup() has not been run.
    var table = Schema.getTable('Activities');
    table.columns.push({ name: 'Follow_Up_Owner', type: 'string', max: 100 });
    Repository.resetCache('Activities');
    try {
      var base = {
        Case_ID: 'SRC-2026-0001', Module: 'M1', Activity_Date: new Date(),
        Activity_Type: 'CALL', Activity_Description: 'โทรติดตาม',
        Performed_By: 'buyer.a@example.com', Vendor_ID: 'VEN-00001'
      };

      // Nothing to store in it: the older sheet keeps working.
      var ok = Repository.insert('Activities', Object.assign({}, base, { Follow_Up_Owner: '' }),
        { actor: 'buyer.a@example.com' });
      assert(!!ok.Activity_ID, 'a blank in the missing column is harmless');

      var error = assertThrowsCode('INTERNAL', function () {
        Repository.insert('Activities', Object.assign({}, base, { Follow_Up_Owner: 'คุณเอ' }),
          { actor: 'buyer.a@example.com' });
      }, 'a real value cannot be stored');
      assertContains(error.message, 'Follow_Up_Owner', 'the message names the column');
      assertContains(error.message, 'setup()', 'and what to do about it');
    } finally {
      table.columns.pop();
      Repository.resetCache('Activities');
    }
  });
});

test('queryByCase finds only the rows of that Case', function () {
  withFreshDatabase(function () {
    ['SRC-2026-0001', 'SRC-2026-0001', 'SRC-2026-0002'].forEach(function (caseId, i) {
      Repository.insert('Activities', {
        Case_ID: caseId, Module: 'M1', Activity_Date: new Date(), Activity_Type: 'CALL',
        Activity_Description: 'กิจกรรม ' + (i + 1), Performed_By: 'buyer.a@example.com'
      }, { actor: 'buyer.a@example.com' });
    });

    assertEquals(Repository.queryByCase('Activities', 'SRC-2026-0001').length, 2, 'two rows for case 1');
    assertEquals(Repository.queryByCase('Activities', 'SRC-2026-0002').length, 1, 'one row for case 2');
    assertEquals(Repository.queryByCase('Activities', 'SRC-2026-0009').length, 0, 'none for an unknown case');
  });
});

test('softDeleteWhere cascades and every removal is logged', function () {
  withFreshDatabase(function () {
    for (var i = 0; i < 3; i++) {
      Repository.insert('Activities', {
        Case_ID: 'SRC-2026-0001', Module: 'M1', Activity_Date: new Date(),
        Activity_Type: 'CALL', Activity_Description: 'กิจกรรม ' + i,
        Performed_By: 'buyer.a@example.com'
      }, { actor: 'buyer.a@example.com' });
    }
    var removed = Repository.softDeleteWhere('Activities', 'Case_ID', 'SRC-2026-0001',
      { actor: 'buyer.a@example.com', reason: 'ยกเลิกงานทั้งใบ' });

    assertEquals(removed, 3, 'all three cascaded');
    assertEquals(Repository.queryByCase('Activities', 'SRC-2026-0001').length, 0, 'none visible');
    var deletes = Repository.readAll('Change_Log').filter(function (l) {
      return l.Table_Name === 'Activities' && l.Action === 'DELETE';
    });
    assertEquals(deletes.length, 3, 'one DELETE entry per row');
  });
});

test('Change_Log has no update or delete API', function () {
  assertEquals(typeof ChangeLog.update, 'undefined', 'no ChangeLog.update');
  assertEquals(typeof ChangeLog.remove, 'undefined', 'no ChangeLog.remove');
  assertEquals(Schema.getTable('Change_Log').appendOnly, true, 'schema marks it append-only');
  withFreshDatabase(function () {
    assertThrowsCode('INTERNAL', function () {
      Repository.update('Change_Log', 'LOG-00000001', { Reason: 'แก้ประวัติ' }, null, { actor: 'x@example.com' });
    }, 'Repository refuses to update an append-only table');
  });
});

/* ============================================================================
 * Phase 3 — Auth, roles, api envelope
 * ==========================================================================*/

var USERS = {
  buyerA: 'buyer.a@example.com',
  buyerB: 'buyer.b@example.com',
  head: 'head@example.com',
  auditor: 'auditor@example.com',
  admin: 'admin@example.com',
  outsider: 'nobody@example.com'
};

/** Adds the five standard people to the Users sheet of the current test database. */
function seedUsers() {
  [
    [USERS.buyerA, 'บายเออร์ เอ', 'BUYER', 'CAPEX', true],
    [USERS.buyerB, 'บายเออร์ บี', 'BUYER', 'OPEX', true],
    [USERS.head, 'หัวหน้าจัดซื้อ', 'HEAD', '', true],
    [USERS.auditor, 'ผู้ตรวจสอบ', 'AUDITOR', '', true],
    [USERS.admin, 'ผู้ดูแลระบบ', 'ADMIN', '', true]
  ].forEach(function (u) {
    Repository.insert('Users', {
      Email: u[0], Name: u[1], Role: u[2], Responsible_Scope: u[3], Is_Active: u[4]
    }, { actor: 'setup' });
  });
}

/** Runs fn while the server sees `email` as the caller. */
function asUser(email, fn) {
  Auth.__setUserOverride(email);
  try {
    return fn();
  } finally {
    Auth.__setUserOverride(null);
  }
}

/** A fresh database that already has the five standard users. */
/**
 * One supplier on the register, written straight through Repository so the
 * fixture needs no signed-in user. Every activity has to name a vendor, so a
 * database with people in it but nothing to buy from cannot record any work.
 * Its Tax_ID sits well clear of the ones the vendor tests mint for themselves.
 */
var FIXTURE_VENDOR_ID = null;

function seedFixtureVendor() {
  var vendor = Repository.insert('Vendors', {
    Vendor_Name: 'บจก. ผู้ขายประจำชุดทดสอบ',
    Tax_ID: '0999900000001',
    Vendor_Status: 'APPROVED'
  }, { actor: 'setup' });
  FIXTURE_VENDOR_ID = vendor.Vendor_ID;
  return vendor;
}

/** A fresh database that already has the five standard users and one vendor. */
function withUsers(fn) {
  return withFreshDatabase(function (report) {
    seedUsers();
    seedFixtureVendor();
    return fn(report);
  });
}

test('T17 a user who is not in the Users sheet is refused', function () {
  withUsers(function () {
    asUser(USERS.outsider, function () {
      var error = assertApiError(api_bootstrap(), 'UNAUTHORIZED', 'unknown account');
      assertContains(error.message, USERS.outsider, 'the message names the account');
    });
    // An inactive account is refused in the same way.
    var users = Repository.readAll('Users');
    var head = users.filter(function (u) { return u.Email === USERS.head; })[0];
    Repository.update('Users', head.Email, { Is_Active: false }, null, { actor: 'admin' });
    asUser(USERS.head, function () {
      assertApiError(api_bootstrap(), 'UNAUTHORIZED', 'deactivated account');
    });
  });
});

test('api_bootstrap gives each role its own permission set', function () {
  withUsers(function () {
    var buyer = asUser(USERS.buyerA, function () { return assertApiOk(api_bootstrap()); });
    assertEquals(buyer.user.role, 'BUYER', 'role resolved');
    assertEquals(buyer.permissions.canCreateCase, true, 'buyer opens cases');
    assertEquals(buyer.permissions.canCloseCase, true, 'the owning buyer closes their own Case');
    assertEquals(buyer.permissions.canSeeTeamView, false, 'buyer has no team view');
    assertEquals(buyer.permissions.canSetVendorApproval, false, 'buyer cannot approve vendors');
    assert(buyer.lists.BUDGET_TYPE.length > 0, 'lists are delivered for the dropdowns');
    assert(buyer.statuses.length > 0, 'status master is delivered');
    assertEquals(buyer.settings.BUYER_CAN_VIEW_ALL, true, 'settings are delivered');

    var head = asUser(USERS.head, function () { return assertApiOk(api_bootstrap()); });
    assertEquals(head.permissions.canReopen, true, 'head reopens');
    assertEquals(head.permissions.canReassign, true, 'head reassigns');
    assertEquals(head.permissions.canEditAnyCase, true, 'head edits any case');

    var auditor = asUser(USERS.auditor, function () { return assertApiOk(api_bootstrap()); });
    assertEquals(auditor.permissions.canCreateCase, false, 'auditor does not open cases');
    assertEquals(auditor.permissions.canEditAnyCase, false, 'auditor never edits');
    assertEquals(auditor.permissions.canSeeTeamView, true, 'auditor sees the team view');

    var admin = asUser(USERS.admin, function () { return assertApiOk(api_bootstrap()); });
    assertEquals(admin.permissions.canSetVendorApproval, true, 'only admin approves vendors');
    assertEquals(admin.permissions.canEditAnyCase, false, 'admin does not edit cases (SPEC 7)');
  });
});

test('api_clearCache is restricted to ADMIN', function () {
  withUsers(function () {
    asUser(USERS.buyerA, function () {
      assertApiError(api_clearCache(), 'FORBIDDEN', 'buyer may not clear the cache');
    });
    asUser(USERS.head, function () {
      assertApiError(api_clearCache(), 'FORBIDDEN', 'head may not clear the cache');
    });
    asUser(USERS.admin, function () {
      assertEquals(assertApiOk(api_clearCache()).cleared, true, 'admin may');
    });
  });
});

test('the api envelope never leaks a stack trace', function () {
  withUsers(function () {
    asUser(USERS.buyerA, function () {
      var response = handle('api_boom', null, function () {
        throw new Error('ENOENT: secret/internal/path.js line 42');
      });
      assertApiError(response, 'INTERNAL', 'unexpected errors become INTERNAL');
      assertEquals(response.error.message, 'เกิดข้อผิดพลาดภายในระบบ กรุณาลองใหม่อีกครั้ง', 'generic message');
      assertEquals(response.error.details, null, 'no details');
      assertEquals(JSON.stringify(response).indexOf('path.js'), -1, 'nothing internal escapes');
    });
  });
});

test('doGet renders the access-denied page for an unknown account', function () {
  withUsers(function () {
    asUser(USERS.outsider, function () {
      var html = doGet().getContent();
      assertContains(html, 'ไม่มีสิทธิ์เข้าใช้งาน', 'denied heading');
      assertContains(html, USERS.outsider, 'the account is shown');
    });
  });
});

test('escapeHtml neutralises markup in the denied page', function () {
  assertEquals(escapeHtml('<script>alert("x")</script>'),
    '&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;', 'tags and quotes escaped');
  assertEquals(escapeHtml("O'Brien & Co"), 'O&#39;Brien &amp; Co', 'quote and ampersand escaped');
});

/* ============================================================================
 * Phase 4 — Cases, Items, Drive folders
 * ==========================================================================*/

function newCasePayload(overrides) {
  var payload = {
    Request_Date: '2026-01-15',
    Request_Ref: 'MEMO-2026-001',
    Requester_Name: 'คุณสมชาย ใจดี',
    Requester_Email: 'somchai@example.com',
    Department_Code: 'OPS',
    Method: 'NORMAL',
    Budget_Type: 'CAPEX',
    Sub_Type: 'NEW_LOCATION',
    Description: 'ติดตั้งป้ายบิลบอร์ดจุดใหม่ ถนนพระราม 9',
    Required_Date: '2026-03-01',
    Intake_Complete: true,
    Intake_Note: ''
  };
  Object.keys(overrides || {}).forEach(function (k) { payload[k] = overrides[k]; });
  return payload;
}

/** Opens a Case as buyer A and returns its id. */
function createCaseAs(email, overrides) {
  return asUser(email, function () {
    return assertApiOk(api_createCase(newCasePayload(overrides))).caseRecord.Case_ID;
  });
}

test('T1 opening a Case issues SRC-YYYY-0001, sets OPEN, makes a folder and logs CREATE', function () {
  withUsers(function () {
    var result = asUser(USERS.buyerA, function () {
      return assertApiOk(api_createCase(newCasePayload()));
    });
    var c = result.caseRecord;

    assertEquals(c.Case_ID, 'SRC-' + IdGenerator.currentYear() + '-0001', 'first case id of the year');
    assertEquals(c.Status, 'OPEN', 'starts open');
    assertEquals(c.Buyer_Owner, USERS.buyerA, 'the creator owns it');
    assertEquals(c.Version, 1, 'version starts at 1');
    assert(!!c.Drive_Folder_ID, 'a Drive folder was created');
    assertDeepEquals(result.warnings, [], 'no warnings for the first case');

    var folder = DriveApp.getFolderById(c.Drive_Folder_ID);
    assertEquals(folder.getName(), c.Case_ID + ' - ติดตั้งป้ายบิลบอร์ดจุดใหม่ ถนนพระราม 9', 'folder name');
    assertEquals(folder.sharing.access, 'DOMAIN', 'folder shared inside the domain only');

    var creates = logsFor('Cases', c.Case_ID).filter(function (l) { return l.Action === 'CREATE'; });
    assertEquals(creates.length, 1, 'one CREATE entry');
    assertEquals(creates[0].User, USERS.buyerA, 'logged against the buyer');
    assertEquals(creates[0].Case_ID, c.Case_ID, 'log row carries the Case_ID');
  });
});

test('creating a Case rejects codes that are not in Config_Lists', function () {
  withUsers(function () {
    asUser(USERS.buyerA, function () {
      assertApiError(api_createCase(newCasePayload({ Department_Code: 'NOT_A_DEPT' })), 'VALIDATION', 'bad department');
      // SUB_TYPE is filtered by Budget_Type, so an OPEX sub-type under CAPEX is invalid.
      assertApiError(api_createCase(newCasePayload({ Budget_Type: 'CAPEX', Sub_Type: 'GENERAL' })),
        'VALIDATION', 'sub-type must belong to the budget type');
      assertApiError(api_createCase(newCasePayload({ Description: '' })), 'VALIDATION', 'description required');
      assertApiError(api_createCase(newCasePayload({ Requester_Email: 'not-an-email' })), 'VALIDATION', 'bad email');
    });
  });
});

test('an auditor cannot open a Case', function () {
  withUsers(function () {
    asUser(USERS.auditor, function () {
      assertApiError(api_createCase(newCasePayload()), 'FORBIDDEN', 'auditors only read');
    });
  });
});

test('a second Case with the same Request_Ref warns but is still created', function () {
  withUsers(function () {
    var first = createCaseAs(USERS.buyerA);
    var second = asUser(USERS.buyerA, function () {
      return assertApiOk(api_createCase(newCasePayload()));
    });
    assertEquals(second.warnings.length, 1, 'one warning');
    assertContains(second.warnings[0], first, 'names the earlier Case');
    assert(!!second.caseRecord.Case_ID, 'the Case was still created');
  });
});

test('T10 a buyer cannot edit another buyer\'s Case', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);

    asUser(USERS.buyerB, function () {
      var error = assertApiError(api_updateCase(caseId, { Description: 'แก้โดยคนอื่น' }, 1), 'FORBIDDEN',
        'buyer B may not edit');
      assertEquals(error.details.owner, USERS.buyerA, 'the error names the owner');

      // Reading it is fine while BUYER_CAN_VIEW_ALL is on.
      var bundle = assertApiOk(api_getCase(caseId));
      assertEquals(bundle.permissions.canEdit, false, 'read-only for buyer B');
    });

    // HEAD may edit any Case.
    asUser(USERS.head, function () {
      assertApiOk(api_updateCase(caseId, { Description: 'แก้โดยหัวหน้า' }, 1));
    });
    assertEquals(Repository.requireById('Cases', caseId).Description, 'แก้โดยหัวหน้า', 'head edit applied');
  });
});

test('BUYER_CAN_VIEW_ALL = FALSE hides other buyers\' Cases', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    Config.setSetting('BUYER_CAN_VIEW_ALL', 'FALSE');

    asUser(USERS.buyerB, function () {
      assertApiError(api_getCase(caseId), 'FORBIDDEN', 'hidden from other buyers');
    });
    asUser(USERS.auditor, function () {
      assertApiOk(api_getCase(caseId), 'auditors always see everything');
    });
  });
});

test('api_updateCase refuses to change owner or status through the back door', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    asUser(USERS.buyerA, function () {
      assertApiError(api_updateCase(caseId, { Buyer_Owner: USERS.buyerB, Status: 'CLOSED' }, 1),
        'VALIDATION', 'neither field is editable here');
    });
    var stored = Repository.requireById('Cases', caseId);
    assertEquals(stored.Buyer_Owner, USERS.buyerA, 'owner unchanged');
    assertEquals(stored.Status, 'OPEN', 'status unchanged');
  });
});

/** The administrator's own addition: a budget type with no sub types under it. */
function addOtherBudgetType() {
  Config.getSheet('Config_Lists').appendRow(['BUDGET_TYPE', 'OTHER', 'งบอื่นๆ', '', 90, true]);
  Config.clearCache();
}

test('a Case on the OTHER budget type describes itself instead of picking a sub type', function () {
  withUsers(function () {
    addOtherBudgetType();

    asUser(USERS.buyerA, function () {
      assertApiError(api_createCase(newCasePayload({
        Budget_Type: 'OTHER', Sub_Type: '', Budget_Type_Other: ''
      })), 'VALIDATION', 'OTHER without a description');

      var created = assertApiOk(api_createCase(newCasePayload({
        Budget_Type: 'OTHER', Sub_Type: '', Budget_Type_Other: 'งบส่วนกลางฝ่ายปฏิบัติการ'
      }))).caseRecord;
      assertEquals(created.Budget_Type_Other, 'งบส่วนกลางฝ่ายปฏิบัติการ', 'the description is kept');
      assertEquals(Utils.isBlank(created.Sub_Type), true, 'and no sub type is stored');
    });
  });
});

test('every other budget type still has to name a sub type', function () {
  withUsers(function () {
    asUser(USERS.buyerA, function () {
      assertApiError(api_createCase(newCasePayload({ Sub_Type: '' })),
        'VALIDATION', 'CAPEX without a sub type');
    });
  });
});

test('switching the budget type clears the field that no longer applies', function () {
  withUsers(function () {
    addOtherBudgetType();
    var caseId = asUser(USERS.buyerA, function () {
      return assertApiOk(api_createCase(newCasePayload({
        Budget_Type: 'OTHER', Sub_Type: '', Budget_Type_Other: 'งบส่วนกลาง'
      }))).caseRecord.Case_ID;
    });

    asUser(USERS.buyerA, function () {
      // Moving off OTHER without naming a sub type leaves the Case describing
      // nothing at all, so it is refused.
      assertApiError(api_updateCase(caseId, { Budget_Type: 'CAPEX' },
        currentCase(caseId).Version), 'VALIDATION', 'a sub type is needed now');

      assertApiOk(api_updateCase(caseId,
        { Budget_Type: 'CAPEX', Sub_Type: 'RENOVATE' }, currentCase(caseId).Version));
    });

    var after = currentCase(caseId);
    assertEquals(after.Sub_Type, 'RENOVATE', 'the sub type took');
    assertEquals(Utils.isBlank(after.Budget_Type_Other), true,
      'and the OTHER description did not linger in the sheet');
  });
});

test('an activity has to name the vendor it was about', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    asUser(USERS.buyerA, function () {
      assertApiError(api_saveActivity(caseId, activityPayload({ Vendor_ID: '' })),
        'VALIDATION', 'no vendor named');
      assertApiOk(api_saveActivity(caseId, activityPayload()), 'the fixture vendor is fine');
    });
  });
});

test('My Cases lists only what the caller may see and flags overdue next actions', function () {
  withUsers(function () {
    var mine = createCaseAs(USERS.buyerA, { Description: 'งานของเอ' });
    createCaseAs(USERS.buyerB, { Request_Ref: 'MEMO-B', Description: 'งานของบี' });

    Repository.insert('Activities', {
      Case_ID: mine, Module: 'M1', Activity_Date: new Date(), Activity_Type: 'COORDINATE',
      Activity_Description: 'ประสานงานกับหน่วยงานผู้ขอ', Performed_By: USERS.buyerA,
      Next_Action: 'ตามข้อมูลเพิ่มเติม', Next_Action_Date: '2020-01-01', Next_Action_Done: false
    }, { actor: USERS.buyerA, caseId: mine });

    var own = asUser(USERS.buyerA, function () { return assertApiOk(api_listCases({ scope: 'mine' })); });
    assertEquals(own.cases.length, 1, 'only my own Cases by default');
    assertEquals(own.cases[0].Case_ID, mine, 'the right one');
    assertEquals(own.cases[0].nextAction.overdue, true, 'the overdue next action is flagged');
    assertEquals(own.cases[0].nextAction.text, 'ตามข้อมูลเพิ่มเติม', 'next action text');

    var all = asUser(USERS.buyerA, function () { return assertApiOk(api_listCases({ scope: 'all' })); });
    assertEquals(all.cases.length, 2, 'both Cases when asking for all');

    var filtered = asUser(USERS.buyerA, function () {
      return assertApiOk(api_listCases({ scope: 'all', q: 'งานของบี' }));
    });
    assertEquals(filtered.cases.length, 1, 'text search');
    assertEquals(filtered.cases[0].canEdit, false, 'buyer A cannot edit buyer B\'s Case');

    var byStatus = asUser(USERS.buyerA, function () {
      return assertApiOk(api_listCases({ scope: 'all', status: 'CLOSED' }));
    });
    assertEquals(byStatus.cases.length, 0, 'status filter');
  });
});

/* ============================================================================
 * Phase 5 — Vendor master, vendors on a Case, the price matrix
 * ==========================================================================*/

function vendorPayload(overrides) {
  return Object.assign({
    Vendor_Name: 'บริษัท ป้ายไทย จำกัด',
    Tax_ID: '0105500000001',
    Address: '99 ถนนพระราม 9 กรุงเทพฯ',
    Contact_Name: 'คุณมานี',
    Contact_Phone: '021112222',
    Contact_Email: 'sales@paithai.example',
    Categories: 'BILLBOARD,STEEL'
  }, overrides || {});
}

function createVendorAs(email, overrides) {
  return asUser(email, function () {
    return assertApiOk(api_createVendor(vendorPayload(overrides)));
  });
}

/** A Case with two items, ready for vendors to be priced against. */
/** Invites a vendor and prices every item, i.e. produces a valid quotation. */
test('T9 a duplicate Tax_ID is blocked and a duplicate phone only warns', function () {
  withUsers(function () {
    var first = createVendorAs(USERS.buyerA);
    assertDeepEquals(first.warnings, [], 'the first vendor is clean');
    assertEquals(first.vendor.Vendor_Status, 'NEW', 'buyers create vendors as NEW');

    asUser(USERS.buyerA, function () {
      var error = assertApiError(api_createVendor(vendorPayload({ Vendor_Name: 'ชื่ออื่น' })),
        'DUPLICATE', 'same Tax_ID is blocked');
      assertEquals(error.details.existing.Vendor_ID, first.vendor.Vendor_ID, 'the existing vendor is offered');
    });

    // Same phone, different Tax_ID: a red flag the buyer must see, not a block.
    var second = createVendorAs(USERS.buyerA, {
      Vendor_Name: 'บริษัท ป้ายไทย 2 จำกัด',
      Tax_ID: '0105500000002',
      Address: 'ที่อยู่อื่น',
      Contact_Email: 'other@paithai.example'
    });
    assertEquals(second.warnings.length, 1, 'one warning');
    assertContains(second.warnings[0], 'เบอร์โทรศัพท์', 'about the phone number');
    assertContains(second.warnings[0], 'บริษัท ป้ายไทย จำกัด', 'naming the other vendor');
    assert(!!second.vendor.Vendor_ID, 'but the vendor was created');

    asUser(USERS.buyerA, function () {
      assertApiError(api_createVendor(vendorPayload({ Tax_ID: '123' })), 'VALIDATION', 'Tax_ID must be 13 digits');
    });
    // Formatting characters are stripped before the uniqueness check.
    asUser(USERS.buyerA, function () {
      assertApiError(api_createVendor(vendorPayload({ Tax_ID: '0-105-500-000001' })),
        'DUPLICATE', 'dashes do not create a new vendor');
    });
  });
});

test('only an administrator may approve or blacklist a vendor', function () {
  withUsers(function () {
    var vendor = createVendorAs(USERS.buyerA).vendor;

    asUser(USERS.buyerA, function () {
      assertApiError(api_updateVendor(vendor.Vendor_ID, { Vendor_Status: 'APPROVED' }, 1), 'FORBIDDEN', 'buyer');
    });
    asUser(USERS.head, function () {
      assertApiError(api_updateVendor(vendor.Vendor_ID, { Vendor_Status: 'BLACKLIST' }, 1), 'FORBIDDEN', 'head');
    });
    asUser(USERS.admin, function () {
      assertApiOk(api_updateVendor(vendor.Vendor_ID, { Vendor_Status: 'APPROVED' }, 1, 'ตรวจเอกสารครบ'));
    });
    assertEquals(Repository.requireById('Vendors', vendor.Vendor_ID).Vendor_Status, 'APPROVED', 'applied');

    // A buyer may still correct ordinary master data.
    asUser(USERS.buyerA, function () {
      assertApiOk(api_updateVendor(vendor.Vendor_ID, { Contact_Name: 'คุณสมหญิง' }, 2));
    });
  });
});

test('the vendor register is searchable by name and by tax id', function () {
  withUsers(function () {
    createVendorAs(USERS.buyerA, { Vendor_Name: 'บริษัท เมกะไซน์ จำกัด' });

    asUser(USERS.buyerA, function () {
      assertEquals(assertApiOk(api_searchVendors('เมกะ')).vendors.length, 1, 'by name');
      assertEquals(assertApiOk(api_searchVendors('0105500000001')).vendors.length, 1, 'by tax id');
      assertEquals(assertApiOk(api_searchVendors('ไม่มีอยู่จริง')).vendors.length, 0, 'no match');
      assertEquals(assertApiOk(api_searchVendors('')).vendors.length, 2,
        'an empty query lists them all, fixture vendor included');
    });
  });
});

test('an upload lands in the Case folder and is shared inside the domain only', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    var base64 = Utilities.base64Encode([80, 68, 70]);

    var file = asUser(USERS.buyerA, function () {
      return assertApiOk(api_uploadFile(caseId, 'ใบเสนอราคา QT-1.pdf', 'application/pdf', base64));
    });
    assert(!!file.url, 'a URL comes back for Quote_File_URL');
    assertEquals(DriveApp.getFileById(file.fileId).sharing.access, 'DOMAIN', 'domain sharing');

    asUser(USERS.buyerB, function () {
      assertApiError(api_uploadFile(caseId, 'x.pdf', 'application/pdf', base64),
        'FORBIDDEN', 'not the owner');
    });
  });
});

/* ============================================================================
 * Phase 6 — Activities
 * ==========================================================================*/

function activityPayload(overrides) {
  return Object.assign({
    Vendor_ID: FIXTURE_VENDOR_ID,
    Activity_Type: 'COORDINATE',
    Channel: 'EMAIL',
    Activity_Description: 'ส่งอีเมลสรุปความต้องการให้หน่วยงานผู้ขอยืนยัน',
    Activity_Date: '2026-01-20T09:30:00.000Z'
  }, overrides || {});
}

test('T10b a buyer may record an activity on a colleague\'s Case, as themselves', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);

    var saved = asUser(USERS.buyerB, function () {
      return assertApiOk(api_saveActivity(caseId, activityPayload())).activity;
    });

    assertEquals(saved.Performed_By, USERS.buyerB, 'recorded against the person who did it');
    assertEquals(saved.Module, 'M1', 'tagged with the module');
    assertEquals(saved.Created_By, USERS.buyerB, 'and in the audit columns');

    // Buyer B still cannot edit the Case itself.
    asUser(USERS.buyerB, function () {
      assertApiError(api_updateCase(caseId, { Description: 'แก้' }, 1), 'FORBIDDEN', 'still no case edit');
    });
  });
});

test('Performed_By cannot be spoofed from the client', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    var saved = asUser(USERS.buyerB, function () {
      return assertApiOk(api_saveActivity(caseId, activityPayload({
        Performed_By: USERS.head, Module: 'M6', Case_ID: 'SRC-9999-0001'
      }))).activity;
    });
    assertEquals(saved.Performed_By, USERS.buyerB, 'the client value is ignored');
    assertEquals(saved.Module, 'M1', 'module is set by the server');
    assertEquals(saved.Case_ID, caseId, 'the Case comes from the URL, not the payload');
  });
});

test('a next action may be recorded without a due date', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    asUser(USERS.buyerA, function () {
      var saved = assertApiOk(api_saveActivity(caseId, activityPayload({
        Next_Action: 'รอหน่วยงานยืนยันสเปก'
      }))).activity;
      assertEquals(saved.Next_Action, 'รอหน่วยงานยืนยันสเปก', 'the follow-up is kept');
      var stored = Repository.requireById('Activities', saved.Activity_ID);
      assertEquals(Utils.isBlank(stored.Next_Action_Date), true, 'and the due date cell stays empty');

      assertApiOk(api_saveActivity(caseId, activityPayload({
        Next_Action: 'โทรตามหน่วยงาน', Next_Action_Date: '2026-02-01'
      })));
      assertApiError(api_saveActivity(caseId, activityPayload({ Activity_Type: 'NOT_A_TYPE' })),
        'VALIDATION', 'activity type must be in Config_Lists');
    });
  });
});

test('an undated follow-up never hides a dated one on the Case list', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    var overdueDate = Utils.formatDateForTest(Utils.addDays(Utils.today(), -2));

    asUser(USERS.buyerA, function () {
      assertApiOk(api_saveActivity(caseId, activityPayload({
        Activity_Date: '2026-01-10T03:00:00.000Z',
        Next_Action: 'ตามเอกสารที่เลยกำหนด', Next_Action_Date: overdueDate
      })));
      // Written later, so it would win on insertion order alone.
      assertApiOk(api_saveActivity(caseId, activityPayload({
        Activity_Date: '2026-01-20T03:00:00.000Z',
        Next_Action: 'เรื่องที่ยังไม่มีกำหนด'
      })));
    });

    var listed = asUser(USERS.buyerA, function () {
      return assertApiOk(api_listCases({ scope: 'mine' })).cases[0];
    });
    assertEquals(listed.nextAction.text, 'ตามเอกสารที่เลยกำหนด', 'the dated one is the pressing one');
    assertEquals(listed.nextAction.overdue, true, 'and it is still flagged overdue');

    // Once the dated one is ticked off, the undated one is all that is left.
    var dated = ActivityService.listForCase(caseId).filter(function (a) {
      return a.Next_Action === 'ตามเอกสารที่เลยกำหนด';
    })[0];
    asUser(USERS.buyerA, function () {
      assertApiOk(api_setNextActionDone(dated.Activity_ID, true, dated.Version));
    });

    var after = asUser(USERS.buyerA, function () {
      return assertApiOk(api_listCases({ scope: 'mine' })).cases[0];
    });
    assertEquals(after.nextAction.text, 'เรื่องที่ยังไม่มีกำหนด', 'the undated one shows when nothing else is due');
    assertEquals(after.nextAction.overdue, false, 'an undated follow-up is never overdue');
  });
});

test('an undated follow-up is never chased by the daily reminder', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    asUser(USERS.buyerA, function () {
      assertApiOk(api_saveActivity(caseId, activityPayload({
        Next_Action: 'เรื่องที่ไม่มีกำหนด'
      })));
    });
    assertDeepEquals(Notification.buildDailyDigests(), {}, 'nothing to remind anyone about');

    asUser(USERS.buyerA, function () {
      assertApiOk(api_saveActivity(caseId, activityPayload({
        Activity_Date: '2026-01-21T03:00:00.000Z',
        Next_Action: 'เรื่องที่เลยกำหนด',
        Next_Action_Date: Utils.formatDateForTest(Utils.addDays(Utils.today(), -1))
      })));
    });

    var digests = Notification.buildDailyDigests();
    var mine = digests[USERS.buyerA];
    assertEquals(mine.overdue.length, 1, 'the dated one is chased');
    assertEquals(mine.overdue[0].text, 'เรื่องที่เลยกำหนด', 'and it is the right one');
    assertEquals(mine.upcoming.length, 0, 'the undated one is not quietly filed as upcoming');
  });
});

test('an activity may only name a vendor that is on the register', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    var known = createVendorAs(USERS.buyerA).vendor;

    asUser(USERS.buyerA, function () {
      assertApiOk(api_saveActivity(caseId, activityPayload({
        Vendor_ID: known.Vendor_ID, Activity_Type: 'CALL'
      })));
      assertApiError(api_saveActivity(caseId, activityPayload({ Vendor_ID: 'VEN-99999' })),
        'VALIDATION', 'no such vendor');
    });

    // The Case page carries the name, so the timeline needs no second request.
    var bundle = asUser(USERS.buyerA, function () { return assertApiOk(api_getCase(caseId)); });
    assertEquals(bundle.vendorNames[known.Vendor_ID], known.Vendor_Name, 'name resolved for the timeline');
  });
});

/** Writes a vendor straight through Repository, so no API rule or user is needed. */
function makeRegisteredVendor(name, taxId) {
  return Repository.insert('Vendors', {
    Vendor_Name: name, Tax_ID: taxId, Vendor_Status: 'APPROVED'
  }, { actor: 'setup' });
}

test('the timeline resolves only the vendors its own activities name', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    var a = makeRegisteredVendor('บจก. ผู้ขาย ก', '0999900000010');
    var b = makeRegisteredVendor('บจก. ผู้ขาย ข', '0999900000011');
    makeRegisteredVendor('บจก. ผู้ขายที่ไม่เกี่ยวข้อง', '0999900000012');

    asUser(USERS.buyerA, function () {
      assertApiOk(api_saveActivity(caseId, activityPayload({ Vendor_ID: a.Vendor_ID })));
      assertApiOk(api_saveActivity(caseId, activityPayload({
        Vendor_ID: b.Vendor_ID, Activity_Date: '2026-01-21T03:00:00.000Z'
      })));
    });

    var names = asUser(USERS.buyerA, function () {
      return assertApiOk(api_getCase(caseId)).vendorNames;
    });
    assertEquals(names[a.Vendor_ID], 'บจก. ผู้ขาย ก', 'first vendor resolved');
    assertEquals(names[b.Vendor_ID], 'บจก. ผู้ขาย ข', 'second vendor resolved');
    assertEquals(Object.keys(names).length, 2, 'and nothing else was dragged along');
  });
});

test('a vendor removed from the register still shows by name on the timeline', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    var vendor = makeRegisteredVendor('บจก. ที่ถูกลบภายหลัง', '0999900000020');

    asUser(USERS.buyerA, function () {
      assertApiOk(api_saveActivity(caseId, activityPayload({ Vendor_ID: vendor.Vendor_ID })));
    });
    Repository.softDelete('Vendors', vendor.Vendor_ID, null,
      { actor: USERS.admin, reason: 'เลิกใช้ผู้ขายรายนี้' });

    var names = asUser(USERS.buyerA, function () {
      return assertApiOk(api_getCase(caseId)).vendorNames;
    });
    assertEquals(names[vendor.Vendor_ID], 'บจก. ที่ถูกลบภายหลัง',
      'the activity still says who it was about');
  });
});

test('both vendor lookup paths return the same names', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    var expected = {};

    // Past the threshold on purpose, so this Case takes the read-everything path
    // rather than looking each vendor up. Read from the constant so the test keeps
    // exercising that path if the threshold is ever retuned.
    var count = CaseService.MAX_VENDOR_LOOKUPS + 2;
    for (var i = 0; i < count; i++) {
      var vendor = makeRegisteredVendor('บจก. ผู้ขายที่ ' + i, '09999001000' + (10 + i));
      expected[vendor.Vendor_ID] = vendor.Vendor_Name;
      asUser(USERS.buyerA, function () {
        assertApiOk(api_saveActivity(caseId, activityPayload({
          Vendor_ID: vendor.Vendor_ID,
          Activity_Description: 'ติดต่อผู้ขายรายที่ ' + i
        })));
      });
    }

    var names = asUser(USERS.buyerA, function () {
      return assertApiOk(api_getCase(caseId)).vendorNames;
    });
    assertEquals(Object.keys(names).length, count, 'every vendor on the Case resolved');
    Object.keys(expected).forEach(function (id) {
      assertEquals(names[id], expected[id], 'name for ' + id);
    });
  });
});

test('the timeline is newest first and next actions can be ticked off', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);

    asUser(USERS.buyerA, function () {
      assertApiOk(api_saveActivity(caseId, activityPayload({
        Activity_Date: '2026-01-10T03:00:00.000Z', Activity_Description: 'กิจกรรมแรก'
      })));
      assertApiOk(api_saveActivity(caseId, activityPayload({
        Activity_Date: '2026-01-22T03:00:00.000Z', Activity_Description: 'กิจกรรมล่าสุด',
        Next_Action: 'ตามใบเสนอราคา', Next_Action_Date: '2026-02-05'
      })));
    });

    var bundle = asUser(USERS.buyerA, function () { return assertApiOk(api_getCase(caseId)); });
    assertEquals(bundle.activities.length, 2, 'both activities');
    assertEquals(bundle.activities[0].Activity_Description, 'กิจกรรมล่าสุด', 'newest first');

    var latest = bundle.activities[0];
    asUser(USERS.buyerA, function () {
      var done = assertApiOk(api_setNextActionDone(latest.Activity_ID, true, latest.Version)).activity;
      assertEquals(done.Next_Action_Done, true, 'ticked');
    });

    // A completed next action no longer appears on My Cases.
    var list = asUser(USERS.buyerA, function () { return assertApiOk(api_listCases({ scope: 'mine' })); });
    assertEquals(list.cases[0].nextAction, null, 'nothing outstanding');
  });
});

test('a buyer may correct their own activity but not someone else\'s', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    var byB = asUser(USERS.buyerB, function () {
      return assertApiOk(api_saveActivity(caseId, activityPayload())).activity;
    });
    var byA = asUser(USERS.buyerA, function () {
      return assertApiOk(api_saveActivity(caseId, activityPayload({ Activity_Description: 'ของเอ' }))).activity;
    });

    // B fixes their own typo, on someone else's Case.
    asUser(USERS.buyerB, function () {
      assertApiOk(api_saveActivity(caseId, {
        Activity_ID: byB.Activity_ID, Activity_Description: 'แก้คำผิด'
      }, byB.Version));
      assertApiError(api_saveActivity(caseId, {
        Activity_ID: byA.Activity_ID, Activity_Description: 'แก้ของคนอื่น'
      }, byA.Version), 'FORBIDDEN', 'not B\'s entry and not B\'s Case');
    });

    // The Case owner may edit anything on their own Case.
    asUser(USERS.buyerA, function () {
      assertApiOk(api_saveActivity(caseId, {
        Activity_ID: byB.Activity_ID, Channel: 'PHONE'
      }, byB.Version + 1));
    });
    assertEquals(Repository.requireById('Activities', byB.Activity_ID).Performed_By, USERS.buyerB,
      'editing never rewrites who performed it');
  });
});

test('deleting an activity is a soft delete with a reason', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    var activity = asUser(USERS.buyerA, function () {
      return assertApiOk(api_saveActivity(caseId, activityPayload())).activity;
    });

    asUser(USERS.buyerA, function () {
      assertApiError(api_deleteRecord('Activities', activity.Activity_ID, activity.Version, ''),
        'VALIDATION', 'reason required');
      assertApiOk(api_deleteRecord('Activities', activity.Activity_ID, activity.Version, 'บันทึกผิดงาน'));
    });

    assertEquals(Repository.queryByCase('Activities', caseId).length, 0, 'hidden');
    assert(!!Repository.findById('Activities', activity.Activity_ID, { includeDeleted: true }),
      'but still in the sheet for audit');
  });
});

/* ============================================================================
 * Phase 7 — StatusEngine, closing against an EPICOR PR, reopen
 * ==========================================================================*/

function currentCase(caseId) {
  return Repository.requireById('Cases', caseId);
}

/** Moves a Case, always sending the version the sheet currently holds. */
function moveTo(email, caseId, toStatus, reason) {
  return asUser(email, function () {
    return api_changeStatus(caseId, toStatus, currentCase(caseId).Version, reason);
  });
}

/** Closes a Case the way the screen does, with the version the sheet holds. */
function closeCase(email, caseId, overrides) {
  var payload = {
    prNo: 'PR-2026-0001',
    prDate: '2026-02-20',
    companyCode: 'PLANB',
    reason: ''
  };
  Object.keys(overrides || {}).forEach(function (k) { payload[k] = overrides[k]; });
  if (!Object.prototype.hasOwnProperty.call(payload, 'version')) {
    payload.version = currentCase(caseId).Version;
  }
  return asUser(email, function () { return api_closeCase(caseId, payload); });
}

test('T3 closing needs the PR number and the buying company, and writes both with the status', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);

    assertApiError(closeCase(USERS.buyerA, caseId, { prNo: '   ' }), 'VALIDATION', 'no PR number');
    assertApiError(closeCase(USERS.buyerA, caseId, { companyCode: '' }), 'VALIDATION', 'no company');
    assertEquals(currentCase(caseId).Status, 'OPEN', 'neither attempt moved the Case');

    var before = currentCase(caseId).Version;
    var closed = assertApiOk(closeCase(USERS.buyerA, caseId)).caseRecord;

    assertEquals(closed.Status, 'CLOSED', 'closed');
    assertEquals(closed.PR_No, 'PR-2026-0001', 'the PR number was written');
    assertEquals(closed.Company_Code, 'PLANB', 'and the buying company');
    assert(!!closed.Closed_At, 'Closed_At is stamped');
    assertEquals(closed.Version, before + 1, 'one update, not two');

    var logged = logsFor('Cases', caseId).filter(function (l) { return l.Field === 'PR_No'; });
    assertEquals(logged.length, 1, 'the PR number is in the history');
    assertEquals(logged[0].New_Value, 'PR-2026-0001', 'with its value');
  });
});

test('T4 a company that is not in Config_Lists cannot close a Case', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    assertApiError(closeCase(USERS.buyerA, caseId, { companyCode: 'NOT_A_COMPANY' }),
      'VALIDATION', 'the company must come from the register');
    assertEquals(currentCase(caseId).Status, 'OPEN', 'still open');
    assertEquals(currentCase(caseId).PR_No, '', 'and nothing was written');
  });
});

test('T5 the same PR number cannot close two Cases', function () {
  withUsers(function () {
    var first = createCaseAs(USERS.buyerA);
    var second = createCaseAs(USERS.buyerA, { Request_Ref: 'MEMO-2026-002' });

    assertApiOk(closeCase(USERS.buyerA, first));
    var error = assertApiError(closeCase(USERS.buyerA, second, { prNo: 'pr-2026-0001' }),
      'DUPLICATE', 'the same number, in any case');
    assertContains(error.message, first, 'the message names the Case already holding it');
    assertEquals(currentCase(second).Status, 'OPEN', 'the second Case stayed open');

    assertApiOk(closeCase(USERS.buyerA, second, { prNo: 'PR-2026-0002' }), 'a different number is fine');
  });
});

test('T11b two people closing the same Case at once: the second gets CONFLICT', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    var stale = currentCase(caseId).Version;

    assertApiOk(closeCase(USERS.buyerA, caseId, { version: stale }));
    assertApiError(closeCase(USERS.head, caseId, { prNo: 'PR-2026-0009', version: stale }),
      'FORBIDDEN', 'a closed Case is frozen before the version is even looked at');

    // And on a Case that is still open, a stale version is what stops the second write.
    var other = createCaseAs(USERS.buyerA, { Request_Ref: 'MEMO-2026-003' });
    var otherStale = currentCase(other).Version;
    asUser(USERS.buyerA, function () {
      assertApiOk(api_updateCase(other, { Description: 'แก้ไขก่อนปิดงาน' }, otherStale));
    });
    assertApiError(closeCase(USERS.buyerA, other, { prNo: 'PR-2026-0010', version: otherStale }),
      'CONFLICT', 'the version moved on');
  });
});

test('the readiness verdict is the same on the list, the Case page and the transition', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);

    var listed = asUser(USERS.buyerA, function () {
      return assertApiOk(api_listCases({ scope: 'mine' })).cases[0];
    });
    var bundle = asUser(USERS.buyerA, function () { return assertApiOk(api_getCase(caseId)); });

    assertEquals(listed.readiness.ok, false, 'the list says it is not ready');
    assertEquals(bundle.rules.ok, false, 'and so does the Case page');
    assertDeepEquals(listed.readiness.blockers, bundle.rules.blockers, 'word for word the same reasons');
    assertEquals(bundle.rules.blockers.length, 2, 'the PR number and the company');

    assertApiOk(closeCase(USERS.buyerA, caseId));

    var after = asUser(USERS.buyerA, function () { return assertApiOk(api_getCase(caseId)); });
    assertEquals(after.rules.ok, true, 'nothing left to do');
    assertDeepEquals(after.rules.blockers, [], 'and nothing left to say');
  });
});

test('a PR number cleared straight in the sheet reopens the Case on the next recheck', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    assertApiOk(closeCase(USERS.buyerA, caseId));

    // The one path that bypasses the rules: an administrator editing the sheet.
    Repository.update('Cases', caseId, { PR_No: '' }, null, { actor: 'admin@example.com' });

    var result = Rules.recheckCaseRules(caseId);
    assert(!!result.reverted, 'the Case was pulled back open');
    assertEquals(currentCase(caseId).Status, 'OPEN', 'back to OPEN');
    assertContains(result.messages[0], 'เลข PR', 'and the message says why');

    var reverts = logsFor('Cases', caseId).filter(function (l) {
      return l.Field === 'Status' && l.New_Value === 'OPEN';
    });
    assertEquals(reverts.length, 1, 'logged once');
    assertEquals(reverts[0].User, ChangeLog.SYSTEM_USER, 'as the system, not as a person');
  });
});

test('a buyer may cancel or close their own Case but not a colleague\'s', function () {
  withUsers(function () {
    var mine = createCaseAs(USERS.buyerA);
    var theirs = createCaseAs(USERS.buyerB, { Request_Ref: 'MEMO-B' });

    assertApiError(moveTo(USERS.buyerA, mine, 'CANCELLED'), 'VALIDATION', 'a reason is required');
    assertApiOk(moveTo(USERS.buyerA, mine, 'CANCELLED', 'ผู้ขอยกเลิกคำขอ'));
    assertEquals(currentCase(mine).Status, 'CANCELLED', 'cancelled');

    assertApiError(moveTo(USERS.buyerA, theirs, 'CANCELLED', 'ไม่เอาแล้ว'), 'FORBIDDEN',
      'not buyer A\'s Case to cancel');
    assertApiError(closeCase(USERS.buyerA, theirs), 'FORBIDDEN', 'nor to close');
    assertApiOk(closeCase(USERS.head, theirs), 'but HEAD may close any Case');
  });
});

test('an illegal transition is refused even for HEAD', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    assertApiOk(closeCase(USERS.buyerA, caseId));
    var error = assertApiError(moveTo(USERS.head, caseId, 'CANCELLED', 'เปลี่ยนใจ'), 'FORBIDDEN',
      'a closed Case is frozen');
    assertContains(error.message, 'Reopen', 'and the message says what to do instead');

    var open = createCaseAs(USERS.buyerA, { Request_Ref: 'MEMO-2026-004' });
    var refused = assertApiError(moveTo(USERS.head, open, 'NOT_A_STATUS'), 'VALIDATION',
      'an unknown status');
    assertContains(refused.message, 'NOT_A_STATUS', 'named in the message');
  });
});

test('T14 a closed Case is frozen until HEAD reopens it', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    assertApiOk(closeCase(USERS.buyerA, caseId, { reason: 'ส่งมอบเรียบร้อย' }));
    assert(!!currentCase(caseId).Closed_At, 'Closed_At is stamped');

    asUser(USERS.buyerA, function () {
      assertApiError(api_updateCase(caseId, { Description: 'แก้หลังปิดงาน' },
        currentCase(caseId).Version), 'FORBIDDEN', 'no edits');
      assertApiError(api_saveActivity(caseId, activityPayload()), 'FORBIDDEN', 'no activities either');
      assertApiError(api_reopenCase(caseId, 'ขอเปิดใหม่'), 'FORBIDDEN', 'buyers cannot reopen');
    });

    asUser(USERS.head, function () {
      assertApiError(api_reopenCase(caseId, ''), 'VALIDATION', 'reopen needs a reason');
      var result = assertApiOk(api_reopenCase(caseId, 'หน่วยงานขอแก้ไขรายการเพิ่ม'));
      assertEquals(result.restoredTo, 'OPEN', 'restored to the status it held before closing');
    });

    assertEquals(currentCase(caseId).Status, 'OPEN', 'back where it was');
    assertEquals(currentCase(caseId).Closed_At, null, 'Closed_At cleared');
    assertEquals(currentCase(caseId).PR_No, 'PR-2026-0001',
      'the PR number it was closed under stays on the record');

    var reopens = logsFor('Cases', caseId).filter(function (l) { return l.Action === 'REOPEN'; });
    assertEquals(reopens.length, 1, 'logged as REOPEN');
    assertEquals(reopens[0].Reason, 'หน่วยงานขอแก้ไขรายการเพิ่ม', 'with the reason');

    // And it is editable again.
    asUser(USERS.buyerA, function () {
      assertApiOk(api_updateCase(caseId, { Description: 'แก้ไขหลัง reopen' }, currentCase(caseId).Version));
    });
  });
});

test('reopening a cancelled Case that never moved returns it to OPEN', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    assertApiOk(moveTo(USERS.buyerA, caseId, 'CANCELLED', 'ผู้ขอยกเลิก'));
    asUser(USERS.head, function () {
      assertEquals(assertApiOk(api_reopenCase(caseId, 'ผู้ขอกลับมายืนยันว่าต้องการ')).restoredTo,
        'OPEN', 'back to the start');
    });
    asUser(USERS.head, function () {
      assertApiError(api_reopenCase(caseId, 'อีกครั้ง'), 'VALIDATION', 'an open Case needs no reopen');
    });
  });
});

test('the Change_Log of a Case is readable by anyone who may read the Case', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    asUser(USERS.buyerA, function () {
      assertApiOk(api_saveActivity(caseId, activityPayload()));
      assertApiOk(api_updateCase(caseId, { Description: 'แก้รายละเอียด' }, currentCase(caseId).Version));
    });
    assertApiOk(closeCase(USERS.buyerA, caseId));

    var entries = asUser(USERS.auditor, function () {
      return assertApiOk(api_getChangeLog(caseId)).entries;
    });
    assert(entries.length > 5, 'a full history is there');
    assertEquals(entries[0].Timestamp > entries[entries.length - 1].Timestamp, true, 'newest first');
    assert(entries.every(function (e) { return !!e.Action && !!e.User; }), 'every entry is attributed');

    Config.setSetting('BUYER_CAN_VIEW_ALL', 'FALSE');
    asUser(USERS.buyerB, function () {
      assertApiError(api_getChangeLog(caseId), 'FORBIDDEN', 'not readable if the Case is not');
    });
  });
});

test('a module can register its own transition rule without touching Rules.js', function () {
  withUsers(function () {
    var calls = [];
    StatusEngine.registerRule('OPEN', 'CANCELLED', function (ctx) {
      calls.push(ctx.to);
      throw Err.ruleViolation('กฎของโมดูลสมมติ: ยังยกเลิกงานไม่ได้');
    });
    try {
      var caseId = createCaseAs(USERS.buyerA);
      var error = assertApiError(moveTo(USERS.buyerA, caseId, 'CANCELLED', 'ลองดู'),
        'RULE_VIOLATION', 'the new rule ran');
      assertContains(error.message, 'โมดูลสมมติ', 'and its message is what the user sees');
      assertEquals(calls.length, 1, 'called once');
      assertEquals(currentCase(caseId).Status, 'OPEN', 'the Case did not move');
    } finally {
      // Emptying the registry is enough: the defaults reinstall themselves
      // on the next use, through Bootstrap.
      StatusEngine.__resetRegistry();
    }
  });
});

/* ============================================================================
 * Phase 8 — Notifications, reassignment, Team View
 * ==========================================================================*/

/** Mail sent during fn(), as an array of MailApp payloads. */
function mailSentDuring(fn) {
  __test.clearMail();
  fn();
  return __test.sentMail.slice();
}

function mailTo(messages, email) {
  return messages.filter(function (m) { return String(m.to).indexOf(email) !== -1; });
}

test('T15 reassigning a Case moves ownership, logs REASSIGN and e-mails both buyers', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);

    asUser(USERS.buyerA, function () {
      assertApiError(api_reassignCase(caseId, USERS.buyerB, 'ลองโอนเอง'), 'FORBIDDEN', 'buyers do not reassign');
    });

    var messages = mailSentDuring(function () {
      asUser(USERS.head, function () {
        assertApiError(api_reassignCase(caseId, USERS.buyerB, ''), 'VALIDATION', 'a reason is required');
        assertApiError(api_reassignCase(caseId, 'ghost@example.com', 'โอน'), 'VALIDATION', 'unknown user');
        assertApiError(api_reassignCase(caseId, USERS.auditor, 'โอน'), 'VALIDATION', 'auditors do not own work');
        assertApiError(api_reassignCase(caseId, USERS.buyerA, 'โอน'), 'VALIDATION', 'already the owner');

        var result = assertApiOk(api_reassignCase(caseId, USERS.buyerB, 'บายเออร์ เอ ลาคลอด'));
        assertEquals(result.previousOwner, USERS.buyerA, 'previous owner reported');
      });
    });

    assertEquals(currentCase(caseId).Buyer_Owner, USERS.buyerB, 'ownership moved');

    var reassigns = logsFor('Cases', caseId).filter(function (l) { return l.Action === 'REASSIGN'; });
    assertEquals(reassigns.length, 1, 'logged as REASSIGN');
    assertEquals(reassigns[0].Old_Value, USERS.buyerA, 'from');
    assertEquals(reassigns[0].New_Value, USERS.buyerB, 'to');
    assertEquals(reassigns[0].Reason, 'บายเออร์ เอ ลาคลอด', 'with the reason');

    assertEquals(mailTo(messages, USERS.buyerB).length, 1, 'the new owner is told');
    assertEquals(mailTo(messages, USERS.buyerA).length, 1, 'and so is the previous one');
    assertContains(mailTo(messages, USERS.buyerB)[0].subject, caseId, 'the subject names the Case');

    // Buyer B can now edit; buyer A cannot.
    asUser(USERS.buyerB, function () {
      assertApiOk(api_updateCase(caseId, { Description: 'รับช่วงต่อ' }, currentCase(caseId).Version));
    });
    asUser(USERS.buyerA, function () {
      assertApiError(api_updateCase(caseId, { Description: 'ขอแก้' }, currentCase(caseId).Version),
        'FORBIDDEN', 'no longer the owner');
    });
  });
});

test('T7b an auto-revert e-mails the Case owner and never leaks vendor contact details', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    createVendorAs(USERS.buyerA);
    assertApiOk(closeCase(USERS.buyerA, caseId));

    var messages = mailSentDuring(function () {
      // The PR number disappearing from the sheet is what the daily recheck catches.
      Repository.update('Cases', caseId, { PR_No: '' }, null, { actor: 'admin@example.com' });
      Notification.recheckOpenCases();
    });

    var toOwner = mailTo(messages, USERS.buyerA);
    assertEquals(toOwner.length, 1, 'the owner is told');
    assertContains(toOwner[0].subject, 'ย้อนสถานะ', 'the subject says what happened');

    // PDPA — no vendor contact detail may appear in an e-mail (SPEC §11).
    var everything = JSON.stringify(messages);
    Repository.readAll('Vendors').forEach(function (v) {
      ['Contact_Phone', 'Contact_Email', 'Contact_Name', 'Address'].forEach(function (field) {
        if (!v[field]) return;
        assertEquals(everything.indexOf(v[field]), -1, field + ' must not appear in any e-mail');
      });
    });
  });
});

test('the daily job reminds each buyer once, about their own work only', function () {
  withUsers(function () {
    var aCase = createCaseAs(USERS.buyerA, { Description: 'งานของเอ' });
    var bCase = createCaseAs(USERS.buyerB, { Request_Ref: 'MEMO-B', Description: 'งานของบี' });

    var yesterday = Utils.formatDateForTest(Utils.addDays(Utils.today(), -1));
    var tomorrow = Utils.formatDateForTest(Utils.addDays(Utils.today(), 1));
    var farFuture = Utils.formatDateForTest(Utils.addDays(Utils.today(), 30));

    asUser(USERS.buyerA, function () {
      assertApiOk(api_saveActivity(aCase, activityPayload({
        Next_Action: 'ตามใบเสนอราคาจากผู้ขาย A', Next_Action_Date: yesterday
      })));
      assertApiOk(api_saveActivity(aCase, activityPayload({
        Next_Action: 'นัดดูหน้างาน', Next_Action_Date: tomorrow
      })));
      assertApiOk(api_saveActivity(aCase, activityPayload({
        Next_Action: 'เรื่องที่ยังอีกนาน', Next_Action_Date: farFuture
      })));
    });
    asUser(USERS.buyerB, function () {
      assertApiOk(api_saveActivity(bCase, activityPayload({
        Next_Action: 'ตามเอกสารจากหน่วยงาน', Next_Action_Date: yesterday
      })));
    });

    var digests = Notification.buildDailyDigests();
    assertEquals(digests[USERS.buyerA].overdue.length, 1, 'A has one overdue');
    assertEquals(digests[USERS.buyerA].upcoming.length, 1, 'and one due tomorrow');
    assertEquals(digests[USERS.buyerB].overdue.length, 1, 'B has their own');

    var messages = mailSentDuring(function () { dailyReminderJob(); });
    assertEquals(messages.length, 2, 'one digest per buyer, not one per action');
    assertEquals(mailTo(messages, USERS.buyerA).length, 1, 'A got exactly one');

    var aMail = mailTo(messages, USERS.buyerA)[0];
    assertContains(aMail.body, 'ตามใบเสนอราคาจากผู้ขาย A', 'listing the overdue item');
    assertContains(aMail.body, 'นัดดูหน้างาน', 'and the one due tomorrow');
    assertEquals(aMail.body.indexOf('เรื่องที่ยังอีกนาน'), -1, 'but not one outside the horizon');
    assertEquals(aMail.body.indexOf('งานของบี'), -1, 'and nothing belonging to another buyer');
  });
});

test('a finished Case never appears in a reminder', function () {
  withUsers(function () {
    var caseId = createCaseAs(USERS.buyerA);
    asUser(USERS.buyerA, function () {
      assertApiOk(api_saveActivity(caseId, activityPayload({
        Next_Action: 'สิ่งที่ค้างอยู่', Next_Action_Date: Utils.formatDateForTest(Utils.addDays(Utils.today(), -3))
      })));
    });
    assertEquals(Object.keys(Notification.buildDailyDigests()).length, 1, 'reminded while open');

    assertApiOk(closeCase(USERS.buyerA, caseId, { reason: 'จบงาน' }));
    assertDeepEquals(Notification.buildDailyDigests(), {}, 'and silent once it is closed');
  });
});

test('Team View reports workload per buyer and is closed to buyers', function () {
  withUsers(function () {
    var a1 = createCaseAs(USERS.buyerA);
    var a2 = createCaseAs(USERS.buyerA, { Request_Ref: 'MEMO-A2' });
    createCaseAs(USERS.buyerB, { Request_Ref: 'MEMO-B' });

    // A1 already has its PR number, so it is ready to be closed; A2 has slipped.
    Repository.update('Cases', a1, { PR_No: 'PR-2026-0055', Company_Code: 'PLANB' }, null,
      { actor: USERS.buyerA });
    asUser(USERS.buyerA, function () {
      assertApiOk(api_saveActivity(a2, activityPayload({
        Next_Action: 'ตามข้อมูลจากผู้ขอ',
        Next_Action_Date: Utils.formatDateForTest(Utils.addDays(Utils.today(), -2))
      })));
    });

    asUser(USERS.buyerA, function () {
      assertApiError(api_teamView(), 'FORBIDDEN', 'buyers do not get the team view');
    });

    var view = asUser(USERS.head, function () { return assertApiOk(api_teamView()); });
    var byEmail = {};
    view.buyers.forEach(function (b) { byEmail[b.email] = b; });

    assertEquals(byEmail[USERS.buyerA].openCases, 2, 'A carries two open Cases');
    assertEquals(byEmail[USERS.buyerA].overdueActions, 1, 'one of them has slipped');
    assertEquals(byEmail[USERS.buyerA].readyToClose, 1, 'and one is ready to be closed');
    assertEquals(byEmail[USERS.buyerB].readyToClose, 0, 'B has nothing ready yet');
    assertEquals(view.totals.openCases, 3, 'three open in total');
    assertEquals(view.canReassign, true, 'HEAD may reassign from here');

    var auditorView = asUser(USERS.auditor, function () { return assertApiOk(api_teamView()); });
    assertEquals(auditorView.canReassign, false, 'auditors watch, they do not move work');
  });
});

/* ============================================================================
 * HTML file name resolution
 *
 * Apps Script has no folders: clasp turns src/client/views/MyCases.html into a
 * file NAMED "client/views/MyCases". That name depends on .clasp.json keeping
 * "rootDir": "src", which this codebase cannot verify from outside Google. Every
 * page is assembled from eleven such files, so one wrong name means a blank app.
 *
 * These run under the Node mock only; on Apps Script, verifyDeployment() check 3
 * proves the same thing against the real project.
 * ==========================================================================*/

if (typeof __test !== 'undefined') {

  /** Runs fn with the HTML files renamed by `rename`, then puts everything back. */
  function withHtmlNames(rename, fn) {
    var original = __test.defaultHtmlFiles;
    var renamed = {};
    Object.keys(original).forEach(function (name) { renamed[rename(name)] = original[name]; });
    __test.setHtmlFiles(renamed);
    __resetHtmlNameCache();
    try {
      return fn();
    } finally {
      __test.setHtmlFiles(original);
      __resetHtmlNameCache();
    }
  }

  test('doGet assembles the whole page from the real client files', function () {
    withUsers(function () {
      asUser(USERS.buyerA, function () {
        var html = doGet().getContent();

        assert(html.length > 50000, 'the page is assembled, not a stub — got ' + html.length + ' bytes');
        assertContains(html, 'var App =', 'App.js was included');
        assertContains(html, 'งานของฉัน', 'the My Cases view was included');
        assertContains(html, 'เลขผู้เสียภาษี', 'the vendor register was included');
        assertContains(html, 'ประวัติการแก้ไข', 'the change-history tab was included');
        assertEquals(/<\?[!=]/.test(html), false, 'no scriptlet was left unevaluated');

        // The signed-in user is rendered into the header, escaped.
        assertContains(html, USERS.buyerA, 'the header shows who is signed in');
      });
    });
  });

  test('resolveHtmlName falls back when clasp names the files differently', function () {
    withUsers(function () {
      // rootDir lost from .clasp.json: every file answers to "src/client/..." instead.
      withHtmlNames(function (name) { return 'src/' + name; }, function () {
        assertEquals(resolveHtmlName('client/Index'), 'src/client/Index', 'found under the longer name');
        asUser(USERS.buyerA, function () {
          assert(doGet().getContent().length > 50000, 'the app works rather than going blank');
        });
      });

      // Pushed from inside src/client: the "client/" segment is gone.
      withHtmlNames(function (name) { return name.replace(/^client\//, ''); }, function () {
        assertEquals(resolveHtmlName('client/Index'), 'Index', 'resolved by dropping a segment');
        assertEquals(resolveHtmlName('client/views/MyCases'), 'views/MyCases', 'and for a nested file');
        asUser(USERS.buyerA, function () {
          assert(doGet().getContent().length > 50000, 'the page still renders');
        });
      });

      // Added by hand in the editor: flat names, no slashes at all.
      withHtmlNames(function (name) { return name.split('/').pop(); }, function () {
        assertEquals(resolveHtmlName('client/views/MyCases'), 'MyCases', 'resolved to the bare name');
        asUser(USERS.buyerA, function () {
          assert(doGet().getContent().length > 50000, 'the page still renders');
        });
      });
    });
  });

  test('a genuinely missing page file gives an error that says what to check', function () {
    withUsers(function () {
      withHtmlNames(function (name) { return 'nowhere/' + name.split('/').pop(); }, function () {
        var error = assertThrowsCode('INTERNAL', function () {
          resolveHtmlName('client/views/MyCases');
        }, 'nothing matches');
        assertContains(error.message, 'client/views/MyCases', 'it names the file it wanted');
        assertContains(error.message, 'views/MyCases', 'and lists the names it tried');
        assertContains(error.message, 'rootDir', 'and points at the likely cause');
      });
    });
  });

  test('every file Index.html includes can be resolved', function () {
    withUsers(function () {
      var expected = [
        'client/Styles', 'client/App.js',
        'client/views/MyCases', 'client/views/CaseDetail',
        'client/views/CaseStatus', 'client/views/CaseActivity', 'client/views/CaseHistory',
        'client/views/Vendors', 'client/views/TeamView'
      ];
      expected.forEach(function (name) {
        assertEquals(resolveHtmlName(name), name, name + ' resolves');
        assert(include(name).length > 0, name + ' has content');
      });
    });
  });
}

/* ============================================================================
 * Load order — Apps Script picks its own, so nothing may depend on it
 * ==========================================================================*/

test('the rule registry fills itself on first use, not at load time', function () {
  withUsers(function () {
    // This is the state a fresh Apps Script execution starts in.
    StatusEngine.__resetRegistry();

    var caseId = createCaseAs(USERS.buyerA);
    assertApiError(moveTo(USERS.buyerA, caseId, 'CLOSED'), 'RULE_VIOLATION',
      'the closing rule is in force without anyone having registered it');

    StatusEngine.__resetRegistry();
    assertApiOk(closeCase(USERS.buyerA, caseId));
    Repository.update('Cases', caseId, { PR_No: '' }, null, { actor: 'admin@example.com' });
    assert(!!Rules.recheckCaseRules(caseId).reverted, 'and the recheck side is installed too');

    StatusEngine.__resetRegistry();
    var bundle = asUser(USERS.buyerA, function () { return assertApiOk(api_getCase(caseId)); });
    assert(bundle.nextStatuses.length > 0, 'allowedNextFor installs the registry as well');
  });
});

/* ============================================================================
 * verifyDeployment — the post-deployment self-check
 * ==========================================================================*/

function checkById(report, id) {
  var found = report.checks.filter(function (c) { return c.id === id; })[0];
  assert(!!found, 'no check with id ' + id);
  return found;
}

/** Brings a freshly set up database up to "ready for production". */
function makeProductionReady() {
  seedUsers();
  installTriggers();

  var folder = DriveApp.createFolder('ไฟล์แนบงานจัดซื้อ');
  Config.setSetting('DRIVE_ROOT_FOLDER_ID', folder.getId());

  // Replace the sample dropdown values with the company's own.
  var sheet = Config.getSheet('Config_Lists');
  sheet.appendRow(['DEPARTMENT', 'OOH', 'ฝ่ายสื่อนอกบ้าน', '', 40, true]);
  sheet.appendRow(['COMPANY', 'PLANB_SUB', 'บริษัทในเครือ (ทดสอบ)', '', 20, true]);
  Config.clearCache();
}

test('verifyDeployment passes every check on a correctly configured system', function () {
  withFreshDatabase(function () {
    makeProductionReady();

    var report = verifyDeployment();
    var failing = report.checks.filter(function (c) { return c.status !== 'PASS'; });
    assertEquals(failing.map(function (c) { return c.id + '=' + c.status + ' (' + c.detail + ')'; }).join('; '),
      '', 'no check should be anything but PASS');

    assertEquals(report.ok, true, 'ok');
    assertEquals(report.readyForProduction, true, 'ready for production');
    assert(!!report.info.webAppUrl, 'reports the web app URL for the administrator');
    assert(!!report.info.databaseUrl, 'and the database URL');
    // The five seeded people plus the ADMIN that setup() adds for whoever ran it.
    assertEquals(report.info.activeUsers, 6, 'and how many people can sign in');
  });
});

test('verifyDeployment warns about the sample dropdown values setup() seeds', function () {
  withFreshDatabase(function () {
    seedUsers();
    var report = Verify.run();

    var sample = checkById(report, 'sampleData');
    assertEquals(sample.status, 'WARN', 'still the examples');
    assertContains(sample.detail, 'DEPARTMENT', 'names the list');
    assertContains(sample.detail, 'OPS', 'and shows the values that are still in place');

    assertEquals(report.ok, true, 'a warning does not make the system unusable');
    assertEquals(report.readyForProduction, false, 'but it is not ready for the whole department');
  });
});

test('verifyDeployment fails loudly when nobody can sign in or approve', function () {
  withFreshDatabase(function () {
    // Nothing but the ADMIN that setup() seeded for whoever ran it.
    var users = checkById(Verify.run(), 'users');
    assertEquals(users.status, 'FAIL', 'no HEAD means exceptions can never be approved');
    assertContains(users.detail, 'HEAD', 'and it says which role is missing');

    Config.getSheet('Users').appendRow(['head@example.com', 'หัวหน้า', 'HEAD', '', true]);
    Repository.resetCache();
    var withHead = checkById(Verify.run(), 'users');
    assertEquals(withHead.status, 'WARN', 'now only the buyers are missing');
    assertContains(withHead.detail, 'BUYER', 'and it says so');
  });
});

test('installTriggers leaves exactly one daily trigger, however often it is run', function () {
  withFreshDatabase(function () {
    seedUsers();
    installTriggers();
    installTriggers();
    installTriggers();
    var check = checkById(Verify.run(), 'trigger');
    assertEquals(check.status, 'PASS', 'no duplicates were left behind');
    assertContains(check.detail, '1 ตัว', 'exactly one');
  });
});

if (typeof __test !== 'undefined') {

  test('verifyDeployment warns when the daily trigger was never installed', function () {
    withFreshDatabase(function () {
      seedUsers();
      var check = checkById(Verify.run(), 'trigger');
      assertEquals(check.status, 'WARN', 'not installed yet');
      assertContains(check.detail, 'installTriggers', 'and says what to run');
      assertContains(check.detail, 'เลข PR', 'and why it matters beyond the e-mail');
    });
  });

  test('verifyDeployment catches HTML files that clasp named unexpectedly', function () {
    withFreshDatabase(function () {
      makeProductionReady();
      assertEquals(checkById(Verify.run(), 'htmlFiles').status, 'PASS', 'baseline is clean');

      // rootDir lost: the app still runs, but the configuration has drifted.
      withHtmlNames(function (name) { return 'src/' + name; }, function () {
        var drifted = checkById(Verify.run(), 'htmlFiles');
        assertEquals(drifted.status, 'WARN', 'works, but flagged');
        assertContains(drifted.detail, 'rootDir', 'and points at the cause');
      });

      // Push incomplete: files genuinely absent.
      withHtmlNames(function (name) { return 'unrelated/' + name.split('/').pop(); }, function () {
        var broken = checkById(Verify.run(), 'htmlFiles');
        assertEquals(broken.status, 'FAIL', 'the app cannot render at all');
        assertContains(broken.detail, 'clasp push', 'and says what to check');
      });
    });
  });
}

test('verifyDeployment reports a broken Drive folder id rather than throwing', function () {
  withFreshDatabase(function () {
    seedUsers();
    var healthy = Verify.run();
    assertEquals(checkById(healthy, 'driveFolder').status, 'WARN', 'blank means auto-create');

    Config.setSetting('DRIVE_ROOT_FOLDER_ID', 'folder_that_does_not_exist');
    var broken = checkById(Verify.run(), 'driveFolder');
    assertEquals(broken.status, 'FAIL', 'set but unreachable is a real problem');

    // One bad check must not stop the others from running. Counted against the
    // healthy run rather than a literal, so adding a check does not fail this.
    assertEquals(Verify.run().checks.length, healthy.checks.length, 'every check still ran');
    assert(healthy.checks.length >= 12, 'and there is a real checklist to run');
  });
});

/* ============================================================================
 * Duplicate primary keys
 *
 * The counter that mints ids once handed the same one out twice, because writes
 * were not flushed before the lock was released. The flush is fixed, but rows
 * written before that keep their shared ids, and the failure was invisible:
 * findById took the first match, so a vendor who had never been invited to a
 * Case came back as "already invited" — the duplicate check was comparing a real
 * id against a row belonging to a different company.
 * ==========================================================================*/

/** Rewrites one row's primary key, which is what the id collision looked like. */
function forceDuplicateKey(tableName, idToBreak, idToUse) {
  var table = Schema.getTable(tableName);
  var sheet = Config.getSheet(table.sheet);
  var meta = Repository.getHeaders(tableName);
  var column = meta.index[table.pk] + 1;
  var keys = sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === idToBreak) {
      sheet.getRange(i + 2, column, 1, 1).setValues([[idToUse]]);
      Repository.resetCache(tableName);
      return i + 2;
    }
  }
  throw new Error('no row holding ' + idToBreak + ' in ' + tableName);
}

test('findById refuses to guess when two live rows share one primary key', function () {
  withUsers(function () {
    var first = createVendorAs(USERS.admin).vendor;
    var second = createVendorAs(USERS.admin, {
      Vendor_Name: 'บริษัท อื่น จำกัด', Tax_ID: '0105500000002'
    }).vendor;
    // Name the row to break rather than assuming which number it was issued.
    var brokenRow = forceDuplicateKey('Vendors', second.Vendor_ID, first.Vendor_ID);

    var threw = null;
    try {
      Repository.findById('Vendors', first.Vendor_ID);
    } catch (e) {
      threw = e;
    }
    assert(!!threw, 'reading an ambiguous id is an error, not a coin flip');
    assertEquals(threw.code, 'INTERNAL', 'reported as a broken sheet');
    assertContains(threw.message, first.Vendor_ID, 'the message names the id');
    assertContains(threw.message, 'Vendors', 'and the sheet it is in');
    assertContains(threw.message, String(brokenRow), 'and the row to go and look at');
  });
});

test('a deleted row sharing an id is not ambiguous, so the live row still wins', function () {
  withUsers(function () {
    var live = createVendorAs(USERS.admin).vendor;
    var gone = createVendorAs(USERS.admin, {
      Vendor_Name: 'บจก. ที่ถูกลบแล้ว', Tax_ID: '0105500000002'
    }).vendor;

    Repository.softDelete('Vendors', gone.Vendor_ID, gone.Version,
      { actor: USERS.admin, reason: 'เพิ่มผิด' });
    forceDuplicateKey('Vendors', gone.Vendor_ID, live.Vendor_ID);

    var found = Repository.findById('Vendors', live.Vendor_ID);
    assertEquals(found.Vendor_Name, live.Vendor_Name, 'the one row still in use is returned');

    var check = checkById(Verify.run(), 'duplicateIds');
    assertEquals(check.status, 'WARN', 'worth cleaning up, but nothing is reading the wrong row');
  });
});

test('verifyDeployment finds every duplicated primary key in one pass', function () {
  withFreshDatabase(function () {
    seedUsers();
    assertEquals(checkById(Verify.run(), 'duplicateIds').status, 'PASS', 'a clean database');

    createVendorAs(USERS.admin);
    createVendorAs(USERS.admin, { Vendor_Name: 'บจก. สอง', Tax_ID: '0105500000002' });
    forceDuplicateKey('Vendors', 'VEN-00002', 'VEN-00001');

    var check = checkById(Verify.run(), 'duplicateIds');
    assertEquals(check.status, 'FAIL', 'a shared id among live rows is a real fault');
    assertContains(check.detail, 'Vendors', 'names the sheet');
    assertContains(check.detail, 'VEN-00001', 'names the id');
    assertContains(check.detail, 'Counters', 'and says to raise the counter as well');
  });
});

test('api_verifyDeployment is restricted to ADMIN', function () {
  withUsers(function () {
    asUser(USERS.head, function () {
      assertApiError(api_verifyDeployment(), 'FORBIDDEN', 'not for the head of procurement');
    });
    asUser(USERS.admin, function () {
      assert(assertApiOk(api_verifyDeployment()).checks.length > 0, 'admins may run it from the app');
    });
  });
});

/* ============================================================================
 * Client views — a button that only some roles see must never be bound blindly
 *
 * Eight buttons in the UI are rendered only when the signed-in role is allowed
 * to use them. Binding one without checking it exists throws, and because the
 * view renders in one pass, that throw takes the entire page down rather than
 * disabling one button. setup() makes the first installer an ADMIN, who may not
 * open Cases, so the one place this was missed broke the first screen for the
 * first user of every new installation.
 *
 * Node-only: it reads the client files, which the runner hands to the mock.
 * ==========================================================================*/

if (typeof __test !== 'undefined') {

  test('every conditionally rendered control is null-checked before binding', function () {
    var files = __test.defaultHtmlFiles;
    var offenders = [];
    var checked = 0;

    Object.keys(files).forEach(function (name) {
      var source = files[name];

      // Ids emitted from inside a ternary are the ones that may be absent.
      var conditional = {};
      var pattern = /\?\s*'<[^']*id="([A-Za-z0-9_-]+)"/g;
      var match;
      while ((match = pattern.exec(source)) !== null) {
        conditional[match[1]] = true;
      }

      Object.keys(conditional).forEach(function (id) {
        checked++;
        // Flags querySelector('#id').something — the reach-through that throws.
        var chained = new RegExp("querySelector\\(['\"]#" + id + "['\"]\\)\\s*\\.");
        if (chained.test(source)) {
          offenders.push(name + ' → #' + id);
        }
      });
    });

    assert(checked >= 4, 'the scan found the conditional controls (found ' + checked + ')');
    assertEquals(offenders.join(', '), '',
      'these controls are bound without checking they exist: ' + offenders.join(', '));
  });
}

/* ============================================================================
 * Writes must reach the sheet before the lock is handed to the next execution
 *
 * Apps Script buffers writes and chooses when to send them. Releasing the script
 * lock without forcing them out first makes the lock decorative: the next
 * execution acquires it, reads the counter, and sees the value the previous one
 * had already replaced — so two Cases receive the same Case_ID. That is exactly
 * what happened on the first real deployment, where one Case_ID came back nine
 * times after the create button was clicked in quick succession.
 *
 * The mock writes synchronously and cannot reproduce the buffering, so it checks
 * the contract instead: flush is called while the lock is still held.
 * ==========================================================================*/

if (typeof __test !== 'undefined') {

  test('T2b every locked write is flushed before the lock is released', function () {
    withFreshDatabase(function () {
      __test.clearLockEvents();
      IdGenerator.next('Cases');
      assertEquals(__test.lockEvents.join(' '), 'tryLock flush releaseLock',
        'minting an id must flush inside the lock');

      // The same has to hold for the repository, which is where every other
      // write in the system goes.
      __test.clearLockEvents();
      Repository.insert('Vendors', {
        Vendor_Name: 'ผู้ขายทดสอบ flush', Tax_ID: '0105512340001', Vendor_Status: 'NEW'
      }, { actor: 'buyer.a@example.com' });

      var events = __test.lockEvents;
      var lastFlush = events.lastIndexOf('flush');
      var lastRelease = events.lastIndexOf('releaseLock');
      assert(lastFlush !== -1, 'an insert flushes at all');
      assert(lastFlush < lastRelease, 'and the flush comes before the release — got ' + events.join(' '));
      assertEquals(events[events.length - 1], 'releaseLock', 'the release is last');
    });
  });
}
