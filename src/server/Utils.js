/**
 * Utils.js — small shared helpers. No business rules live here.
 */
var Utils = (function () {

  var LOCK_TIMEOUT_MS = 10000;            // SPEC §11 — 10s, then tell the user to retry

  var lockDepth = 0;                      // an execution is single-threaded, so this is exact

  /**
   * Runs fn while holding the script lock (SPEC §11).
   * Every write path that mints an ID or must not interleave goes through here.
   *
   * Reentrant: a service that already holds the lock can call a repository method
   * that also asks for it without deadlocking or double-releasing.
   */
  function withScriptLock(fn) {
    if (lockDepth > 0) {
      lockDepth++;
      try {
        return fn();
      } finally {
        lockDepth--;
      }
    }
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
      throw Err.conflict('ระบบกำลังถูกใช้งานโดยผู้อื่น กรุณาลองใหม่อีกครั้ง');
    }
    lockDepth++;
    try {
      return fn();
    } finally {
      lockDepth--;
      // Apps Script buffers writes and decides for itself when to send them.
      // Releasing the lock without forcing them out first makes the lock
      // useless: the next execution takes the lock, reads the sheet, and sees
      // the value this one already replaced. That is how two Cases end up
      // sharing a Case_ID. Flush while the lock is still held, so whoever takes
      // it next reads what this execution actually wrote.
      try {
        SpreadsheetApp.flush();
      } catch (e) {
        console.error('SpreadsheetApp.flush() failed before releasing the lock: ' +
          ((e && e.message) || e));
      }
      lock.releaseLock();
    }
  }

  function isLockHeld() {
    return lockDepth > 0;
  }

  function now() {
    return new Date();
  }

  function isBlank(v) {
    return v === null || v === undefined || String(v).trim() === '';
  }

  function toBool(v) {
    if (v === true) return true;
    if (v === false || v === null || v === undefined || v === '') return false;
    var s = String(v).trim().toUpperCase();
    return s === 'TRUE' || s === 'YES' || s === '1';
  }

  /** Accepts a Date, an ISO string or dd/MM/yyyy. Returns null when unparseable. */
  function toDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    var s = String(v).trim();
    var thai = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
    if (thai) {
      var d = new Date(Number(thai[3]), Number(thai[2]) - 1, Number(thai[1]));
      return isNaN(d.getTime()) ? null : d;
    }
    var parsed = new Date(s);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  /** Midnight of the given date, so date-only comparisons ignore the clock. */
  function startOfDay(v) {
    var d = toDate(v);
    if (!d) return null;
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function today() {
    return startOfDay(new Date());
  }

  function addDays(date, days) {
    var d = new Date(date.getTime());
    d.setDate(d.getDate() + days);
    return d;
  }

  /** dd/MM/yyyy in the Gregorian calendar (SPEC §3). */
  function formatDate(v) {
    var d = toDate(v);
    if (!d) return '';
    return Utilities.formatDate(d, Config.getTimezone(), 'dd/MM/yyyy');
  }

  function formatDateTime(v) {
    var d = toDate(v);
    if (!d) return '';
    return Utilities.formatDate(d, Config.getTimezone(), 'dd/MM/yyyy HH:mm');
  }

  /** ISO strings survive the google.script.run boundary; Date objects do not, reliably. */
  function toIso(v) {
    var d = toDate(v);
    return d ? d.toISOString() : '';
  }

  /** yyyy-MM-dd from a Date, using local parts. Handy for tests and date inputs. */
  function formatDateForTest(v) {
    var d = toDate(v);
    if (!d) return '';
    return d.getFullYear() + '-' + pad(d.getMonth() + 1, 2) + '-' + pad(d.getDate(), 2);
  }

  function truncate(s, max) {
    var str = String(s === null || s === undefined ? '' : s);
    return str.length <= max ? str : str.substring(0, max);
  }

  function pad(n, width) {
    var s = String(n);
    while (s.length < width) s = '0' + s;
    return s;
  }

  /** Stable string form used when diffing old vs new values for the Change_Log. */
  function normalizeForCompare(v) {
    if (v === null || v === undefined || v === '') return '';
    if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString();
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return String(v);
    return String(v).trim();
  }

  function unique(arr) {
    var seen = {};
    return arr.filter(function (v) {
      var k = String(v);
      if (seen[k]) return false;
      seen[k] = true;
      return true;
    });
  }

  /**
   * Turns scattered row numbers into contiguous [start, end] runs, so a caller can
   * read them with a handful of getValues() calls instead of one per row.
   */
  function groupRuns(numbers) {
    var sorted = numbers.slice().sort(function (a, b) { return a - b; });
    var runs = [];
    var start = null;
    var previous = null;
    sorted.forEach(function (n) {
      if (start === null) { start = n; previous = n; return; }
      if (n === previous) return;                 // duplicates collapse
      if (n === previous + 1) { previous = n; return; }
      runs.push([start, previous]);
      start = n;
      previous = n;
    });
    if (start !== null) runs.push([start, previous]);
    return runs;
  }

  /** Groups rows by the value of `key`, preserving input order within each group. */
  function groupBy(rows, key) {
    var out = {};
    rows.forEach(function (r) {
      var k = String(r[key] === undefined ? '' : r[key]);
      if (!out[k]) out[k] = [];
      out[k].push(r);
    });
    return out;
  }

  return {
    LOCK_TIMEOUT_MS: LOCK_TIMEOUT_MS,
    withScriptLock: withScriptLock,
    isLockHeld: isLockHeld,
    now: now,
    isBlank: isBlank,
    toBool: toBool,
    toDate: toDate,
    startOfDay: startOfDay,
    today: today,
    addDays: addDays,
    formatDate: formatDate,
    formatDateTime: formatDateTime,
    formatDateForTest: formatDateForTest,
    toIso: toIso,
    truncate: truncate,
    pad: pad,
    normalizeForCompare: normalizeForCompare,
    unique: unique,
    groupRuns: groupRuns,
    groupBy: groupBy
  };
})();
