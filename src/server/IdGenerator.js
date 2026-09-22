/**
 * IdGenerator.js — running numbers from the Counters sheet, under the script lock.
 *
 * SPEC §5.4 / Acceptance Test 2: two buyers creating a Case at the same moment must
 * never receive the same Case_ID. Every allocation therefore happens inside
 * Utils.withScriptLock, and a batch of n IDs is reserved with a single increment.
 */
var IdGenerator = (function () {

  var COUNTER_SHEET = 'Counters';

  /** Mints one ID for the given table, e.g. next('Cases') -> 'SRC-2026-0001'. */
  function next(tableName) {
    return reserve(tableName, 1)[0];
  }

  /**
   * Reserves `count` consecutive IDs in one locked read-modify-write.
   * Used when a batch insert (e.g. several Change_Log rows) needs many IDs at once.
   */
  function reserve(tableName, count) {
    if (count <= 0) return [];
    var table = Schema.getTable(tableName);
    if (!table.id) {
      throw Err.internal('ตาราง ' + tableName + ' ไม่ได้กำหนดรูปแบบ ID ไว้ใน Schema');
    }
    return Utils.withScriptLock(function () {
      var year = table.id.yearly ? currentYear() : '';
      var last = bumpCounter(tableName, year, count);
      var ids = [];
      for (var i = 0; i < count; i++) {
        ids.push(format(table.id, year, last - count + 1 + i));
      }
      return ids;
    });
  }

  /** 'SRC-2026-0001' when the counter resets yearly, otherwise 'ITM-000001'. */
  function format(idSpec, year, number) {
    var body = Utils.pad(number, idSpec.pad);
    return year ? idSpec.prefix + '-' + year + '-' + body : idSpec.prefix + '-' + body;
  }

  function currentYear() {
    return Utilities.formatDate(new Date(), Config.getTimezone(), 'yyyy');
  }

  /**
   * Adds `count` to the counter row and returns the new Last_No.
   * Creates the row on first use so Setup does not have to pre-seed every table.
   * Caller must already hold the script lock.
   */
  function bumpCounter(counterName, year, count) {
    var sheet = Config.getSheet(COUNTER_SHEET);
    var lastRow = sheet.getLastRow();
    var values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];

    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === counterName && String(values[i][1]).trim() === String(year)) {
        var updated = (Number(values[i][2]) || 0) + count;
        sheet.getRange(i + 2, 3).setValue(updated);
        return updated;
      }
    }
    sheet.appendRow([counterName, year, count]);
    return count;
  }

  /** Current value without consuming it — used by tests and by the setup report. */
  function peek(counterName, year) {
    var sheet = Config.getSheet(COUNTER_SHEET);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;
    var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    var wanted = year === undefined ? '' : String(year);
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === counterName && String(values[i][1]).trim() === wanted) {
        return Number(values[i][2]) || 0;
      }
    }
    return 0;
  }

  return {
    COUNTER_SHEET: COUNTER_SHEET,
    next: next,
    reserve: reserve,
    format: format,
    currentYear: currentYear,
    peek: peek
  };
})();
