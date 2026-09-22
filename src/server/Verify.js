/**
 * Verify.js — the post-deployment self-check.
 *
 * Everything in this project is tested against mocks, which proves the logic but
 * cannot prove that a particular Google account, spreadsheet and deployment were
 * wired together correctly. This file closes that gap: one function an
 * administrator runs from the Apps Script editor that turns a silent
 * misconfiguration into a list of things to fix.
 *
 * It only reads. Running it twice changes nothing.
 */
var Verify = (function () {

  var PASS = 'PASS';
  var WARN = 'WARN';
  var FAIL = 'FAIL';

  /** Every HTML file the single page is assembled from. */
  var HTML_FILES = [
    'client/Index', 'client/Styles', 'client/App.js',
    'client/views/MyCases', 'client/views/CaseDetail',
    'client/views/CaseStatus', 'client/views/CaseActivity', 'client/views/CaseHistory',
    'client/views/Vendors', 'client/views/TeamView'
  ];

  /** Lists the application cannot function without. */
  var REQUIRED_LISTS = [
    'BUDGET_TYPE', 'SUB_TYPE', 'METHOD', 'ACTIVITY_TYPE', 'CHANNEL',
    'DEPARTMENT', 'COMPANY', 'VENDOR_STATUS', 'ROLE'
  ];

  /**
   * The placeholder codes setup() seeds so the system starts up at all. They are
   * examples, not this company's data, and going live on them would put
   * "ฝ่ายปฏิบัติการ / ฝ่ายการตลาด / ฝ่ายการเงิน" in front of every buyer.
   */
  var SAMPLE_CODES = {
    DEPARTMENT: ['OPS', 'MKT', 'FIN'],
    // A Case is closed against the company that bought the goods, so a register
    // holding only the parent company means the affiliates were never entered.
    COMPANY: ['PLANB']
  };

  var TRANSITIONS = {
    OPEN: ['CLOSED', 'CANCELLED']
  };

  function run() {
    var checks = [];
    var info = {};

    check(checks, 'database', 'ไฟล์ฐานข้อมูล', function () {
      var id = PropertiesService.getScriptProperties().getProperty(Config.DB_PROPERTY_KEY);
      if (!id) return fail('ยังไม่ได้ตั้ง ' + Config.DB_PROPERTY_KEY + ' — กรุณารัน setup() ก่อน');
      var db = Config.getDb();
      info.databaseUrl = db.getUrl();
      return pass('เปิดไฟล์ได้: ' + db.getName());
    });

    check(checks, 'sheets', 'ชีตและหัวคอลัมน์ครบตาม Schema', function () {
      var problems = [];
      Schema.tableNames().forEach(function (tableName) {
        var sheetName = Schema.getTable(tableName).sheet;
        var sheet = Config.getDb().getSheetByName(sheetName);
        if (!sheet) { problems.push('ไม่มีชีต ' + sheetName); return; }
        var headers = Repository.getHeaders(tableName).headers;
        var missing = Schema.getColumnNames(tableName).filter(function (column) {
          return headers.indexOf(column) === -1;
        });
        if (missing.length) problems.push(sheetName + ' ขาดคอลัมน์ ' + missing.join(', '));
      });
      if (problems.length) return fail(problems.join(' · ') + ' — กรุณารัน setup() อีกครั้ง');
      return pass('ครบทั้ง ' + Schema.tableNames().length + ' ตาราง');
    });

    // The counter that mints primary keys once handed the same id out twice,
    // because writes were not being flushed before the lock was released. The
    // code is fixed; the rows it already wrote are not, and a duplicate key is
    // invisible until something reads by that key and silently gets the wrong
    // row. This finds every one of them in one pass.
    check(checks, 'duplicateIds', 'รหัสประจำตัว (primary key) ไม่ซ้ำ', function () {
      var live = [];
      var deletedOnly = [];

      Schema.tableNames().forEach(function (tableName) {
        var table = Schema.getTable(tableName);
        if (!table.pk) return;
        var counts = countKeys(tableName, table);
        if (counts === null) return;                  // empty sheet, nothing to compare
        if (counts.live.length) live.push(table.sheet + ': ' + counts.live.join(', '));
        if (counts.withDeleted.length) {
          deletedOnly.push(table.sheet + ': ' + counts.withDeleted.join(', '));
        }
      });

      if (live.length) {
        return fail('พบรหัสซ้ำในแถวที่ใช้งานอยู่ — ' + live.join(' · ') +
          ' · ระบบจะอ่านข้อมูลผิดแถว กรุณาแก้รหัสของแถวที่ยังไม่มีใครอ้างถึงให้เป็นเลขใหม่ ' +
          'แล้วตั้ง Last_No ในชีต Counters ให้สูงกว่าเลขที่ใช้อยู่ทั้งหมด');
      }
      if (deletedOnly.length) {
        return warn('รหัสซ้ำเฉพาะกับแถวที่ถูกลบไปแล้ว — ' + deletedOnly.join(' · ') +
          ' · ระบบยังอ่านถูกแถว แต่ประวัติการแก้ไขจะอ่านยาก');
      }
      return pass('ไม่มีรหัสซ้ำในทุกตาราง');
    });

    // The one thing the Node tests cannot prove: that clasp named the files the
    // way this code asks for them.
    check(checks, 'htmlFiles', 'ไฟล์หน้าเว็บ (ตรวจชื่อที่ clasp ตั้งให้จริง)', function () {
      __resetHtmlNameCache();
      var missing = [];
      var renamed = [];
      HTML_FILES.forEach(function (logical) {
        var actual;
        try {
          actual = resolveHtmlName(logical);
        } catch (e) {
          missing.push(logical);
          return;
        }
        if (actual !== logical) renamed.push(logical + ' → ' + actual);
      });
      if (missing.length) {
        return fail('หาไม่พบ ' + missing.length + ' ไฟล์: ' + missing.join(', ') +
          ' — กรุณาตรวจว่า clasp push สำเร็จครบทุกไฟล์');
      }
      if (renamed.length) {
        return warn('พบครบ ' + HTML_FILES.length + ' ไฟล์ แต่ชื่อไม่ตรงรูปแบบมาตรฐาน (' +
          renamed.join(', ') + ') — ระบบยังทำงานได้ แต่ควรตรวจว่า .clasp.json ตั้ง "rootDir": "src" ไว้');
      }
      return pass('ครบทั้ง ' + HTML_FILES.length + ' ไฟล์ และชื่อตรงตามที่ควรเป็น');
    });

    check(checks, 'settings', 'ค่าตั้งค่าหลัก', function () {
      var problems = [];
      if (Utils.isBlank(Config.getTimezone())) problems.push('APP_TIMEZONE ว่าง');
      var reminderHour = Config.getNumber('REMINDER_HOUR', -1);
      if (!(reminderHour >= 0 && reminderHour <= 23)) {
        problems.push('REMINDER_HOUR ต้องอยู่ระหว่าง 0-23 (ตอนนี้ "' + Config.get('REMINDER_HOUR', '') + '")');
      }
      if (problems.length) return fail(problems.join(' · '));
      return pass('โซนเวลา ' + Config.getTimezone() + ' · เตือนเวลา ' + reminderHour + ':00');
    });

    check(checks, 'driveFolder', 'โฟลเดอร์หลักใน Google Drive', function () {
      var id = Config.get('DRIVE_ROOT_FOLDER_ID', '');
      if (!id) return warn('ยังไม่ได้ตั้ง DRIVE_ROOT_FOLDER_ID — ระบบจะสร้างโฟลเดอร์ให้เองเมื่อเปิด Case แรก');
      var folder = DriveApp.getFolderById(id);      // throws if unreachable; check() catches it
      info.driveFolderUrl = folder.getUrl();
      return pass('เปิดได้: ' + folder.getName());
    });

    check(checks, 'lists', 'รายการ dropdown ที่จำเป็น', function () {
      var empty = REQUIRED_LISTS.filter(function (name) { return Config.getList(name).length === 0; });
      if (empty.length) {
        return fail('ไม่มีข้อมูลในรายการ ' + empty.join(', ') + ' — กรุณารัน setup() หรือเติมในชีต Config_Lists');
      }
      return pass('ครบทั้ง ' + REQUIRED_LISTS.length + ' รายการ');
    });

    check(checks, 'sampleData', 'ข้อมูลตัวอย่างที่ต้องแทนที่ด้วยของบริษัท', function () {
      var untouched = [];
      Object.keys(SAMPLE_CODES).forEach(function (listName) {
        var codes = Config.getList(listName).map(function (item) { return item.code; });
        if (codes.length === 0) return;
        var stillSample = codes.every(function (code) {
          return SAMPLE_CODES[listName].indexOf(code) !== -1;
        });
        if (stillSample) untouched.push(listName + ' (' + codes.join(', ') + ')');
      });
      if (untouched.length) {
        return warn('ยังเป็นค่าตัวอย่างที่ระบบใส่ให้: ' + untouched.join(' · ') +
          ' — กรุณาแก้ในชีต Config_Lists ให้ตรงกับบริษัทก่อนเปิดใช้งานจริง');
      }
      return pass('แก้เป็นข้อมูลของบริษัทแล้ว');
    });

    check(checks, 'statuses', 'สถานะงานและเส้นทางที่ไปต่อได้', function () {
      var all = Config.getStatusMaster();
      if (all.length < 3) return fail('มีสถานะเพียง ' + all.length + ' รายการ จากที่ควรมี 3 — กรุณารัน setup()');
      var problems = [];
      Object.keys(TRANSITIONS).forEach(function (code) {
        var status = Config.getStatus(code);
        if (!status) { problems.push('ไม่มีสถานะ ' + code); return; }
        TRANSITIONS[code].forEach(function (next) {
          if (status.allowedNext.indexOf(next) === -1) {
            problems.push(code + ' ไปต่อที่ ' + next + ' ไม่ได้');
          }
        });
      });
      if (problems.length) {
        return fail(problems.join(' · ') + ' — กรุณาตรวจคอลัมน์ Allowed_Next ในชีต Status_Master');
      }
      return pass(all.length + ' สถานะ และเส้นทางที่ไปต่อได้ครบ');
    });

    check(checks, 'users', 'ผู้ใช้งาน', function () {
      var users = Auth.listActiveUsers();
      var byRole = {};
      users.forEach(function (u) { byRole[u.role] = (byRole[u.role] || 0) + 1; });
      info.activeUsers = users.length;

      if (users.length === 0) return fail('ไม่มีผู้ใช้ที่ใช้งานอยู่เลยในชีต Users — จะไม่มีใครเข้าระบบได้');
      var missing = [];
      if (!byRole.ADMIN) missing.push('ADMIN');
      if (!byRole.HEAD) missing.push('HEAD');
      if (missing.length) {
        return fail('ยังไม่มีผู้ใช้บทบาท ' + missing.join(' และ ') +
          ' — งานที่ต้องอนุมัติหรือดูแลระบบจะทำไม่ได้');
      }
      if (!byRole.BUYER) {
        return warn('มี ADMIN และ HEAD แล้ว แต่ยังไม่มี BUYER — ยังไม่มีใครเปิดงานจัดซื้อได้');
      }
      return pass(users.length + ' คน: ' + Object.keys(byRole).map(function (role) {
        return role + ' ' + byRole[role];
      }).join(', '));
    });

    check(checks, 'trigger', 'อีเมลแจ้งเตือนรายวัน', function () {
      var installed = ScriptApp.getProjectTriggers().filter(function (t) {
        return t.getHandlerFunction() === 'dailyReminderJob';
      });
      if (installed.length === 0) {
        return warn('ยังไม่ได้ติดตั้ง trigger — กรุณารัน installTriggers() ' +
          '(ไม่มี trigger = ไม่มีอีเมลเตือน และงานที่เลข PR หายไปจากชีตจะไม่ถูกตรวจพบ)');
      }
      if (installed.length > 1) {
        return warn('มี trigger ซ้ำ ' + installed.length + ' ตัว — รัน installTriggers() อีกครั้งเพื่อล้างของเก่า');
      }
      // Apps Script does not expose the hour of an installed trigger, so the
      // configured value cannot be read back and compared here.
      return pass('ติดตั้งแล้ว 1 ตัว (ตั้งไว้ตาม REMINDER_HOUR = ' + Config.getNumber('REMINDER_HOUR', 8) + ')');
    });

    check(checks, 'webApp', 'การ Deploy เป็น Web app', function () {
      var url = '';
      try {
        url = ScriptApp.getService().getUrl() || '';
      } catch (e) {
        url = '';
      }
      info.webAppUrl = url;
      if (!url) {
        return warn('ยังไม่พบ URL ของ web app — ถ้ายังไม่ได้ Deploy กรุณา Deploy เป็น Web app ' +
          '(Execute as: Me, Who has access: โดเมนของบริษัท)');
      }
      return pass(url);
    });

    check(checks, 'mailQuota', 'โควตาอีเมลคงเหลือวันนี้', function () {
      var remaining = MailApp.getRemainingDailyQuota();
      info.mailQuota = remaining;
      if (remaining <= 0) return warn('โควตาอีเมลวันนี้หมดแล้ว อีเมลแจ้งเตือนจะไม่ถูกส่งจนกว่าจะถึงวันถัดไป');
      return pass(remaining + ' ฉบับ');
    });

    var failed = checks.filter(function (c) { return c.status === FAIL; });
    var warned = checks.filter(function (c) { return c.status === WARN; });

    return {
      ok: failed.length === 0,
      readyForProduction: failed.length === 0 && warned.length === 0,
      failed: failed.length,
      warned: warned.length,
      passed: checks.length - failed.length - warned.length,
      checks: checks,
      info: info
    };
  }

  /** Runs one check, turning any throw into a FAIL so the rest still run. */
  /**
   * Counts primary keys in one table, reading only the columns it needs.
   *
   * Returns the ids that appear more than once among live rows, and separately
   * those that only collide once deleted rows are counted too. Reads the key
   * column (and Is_Deleted, when the table has one) rather than whole rows, so
   * this stays affordable on the biggest sheets.
   */
  function countKeys(tableName, table) {
    var meta = Repository.getHeaders(tableName);
    var lastRow = meta.sheet.getLastRow();
    if (lastRow < 2) return null;

    var keyColumn = meta.index[table.pk];
    if (keyColumn === undefined) return null;         // the sheets check reports this
    var keys = meta.sheet.getRange(2, keyColumn + 1, lastRow - 1, 1).getValues();

    var deletedColumn = meta.index.Is_Deleted;
    var deleted = deletedColumn === undefined ? null
      : meta.sheet.getRange(2, deletedColumn + 1, lastRow - 1, 1).getValues();

    var liveCount = {};
    var allCount = {};
    for (var i = 0; i < keys.length; i++) {
      var id = String(keys[i][0]).trim();
      if (!id) continue;
      allCount[id] = (allCount[id] || 0) + 1;
      var isDeleted = deleted !== null &&
        (deleted[i][0] === true || String(deleted[i][0]).trim().toUpperCase() === 'TRUE');
      if (!isDeleted) liveCount[id] = (liveCount[id] || 0) + 1;
    }

    var live = Object.keys(liveCount).filter(function (id) { return liveCount[id] > 1; });
    var withDeleted = Object.keys(allCount).filter(function (id) {
      return allCount[id] > 1 && live.indexOf(id) === -1;
    });
    return { live: live.sort(), withDeleted: withDeleted.sort() };
  }

  function check(checks, id, title, fn) {
    var result;
    try {
      result = fn();
    } catch (e) {
      result = fail((e && e.message) || String(e));
    }
    checks.push({ id: id, title: title, status: result.status, detail: result.detail });
  }

  function pass(detail) { return { status: PASS, detail: detail }; }
  function warn(detail) { return { status: WARN, detail: detail }; }
  function fail(detail) { return { status: FAIL, detail: detail }; }

  return { PASS: PASS, WARN: WARN, FAIL: FAIL, HTML_FILES: HTML_FILES, SAMPLE_CODES: SAMPLE_CODES, run: run };
})();

/**
 * Run this from the Apps Script editor after deploying, and again whenever
 * something stops behaving. Read the execution log for the checklist.
 */
function verifyDeployment() {
  var report = Verify.run();
  var marks = { PASS: '[ ผ่าน ]', WARN: '[ เตือน ]', FAIL: '[ ไม่ผ่าน ]' };

  var lines = ['', '===== ตรวจความพร้อมของระบบ ====='];
  report.checks.forEach(function (c, i) {
    lines.push(marks[c.status] + ' ' + (i + 1) + '. ' + c.title);
    lines.push('          ' + c.detail);
  });
  lines.push('');
  lines.push('สรุป: ผ่าน ' + report.passed + ' · เตือน ' + report.warned + ' · ไม่ผ่าน ' + report.failed);
  lines.push(report.ok
    ? (report.readyForProduction
      ? 'พร้อมใช้งานจริง'
      : 'ใช้งานได้ แต่มีข้อเตือนที่ควรแก้ก่อนเปิดให้ทั้งฝ่ายใช้')
    : 'ยังใช้งานไม่ได้ กรุณาแก้ข้อที่ "ไม่ผ่าน" ก่อน');
  lines.push('================================');

  console.log(lines.join('\n'));
  return report;
}
