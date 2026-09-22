/**
 * Repository.js — the only code in the system that reads or writes a data sheet.
 *
 * Every service goes through here, which is what guarantees (SPEC §10.2) that
 * audit columns, Version, optimistic locking, soft delete and the Change_Log are
 * handled identically everywhere, and that a new module gets all of it for free.
 *
 * Two invariants worth stating outright:
 *   - Rows are addressed by ID, never by row number from the caller's side (SPEC §2.6).
 *   - Columns are located by header name, never by a fixed index, so a column
 *     appended by a later setup() run cannot shift anything.
 */
var Repository = (function () {

  /** Columns the caller may never set directly; Repository owns them. */
  var MANAGED_COLUMNS = ['Created_At', 'Created_By', 'Updated_At', 'Updated_By', 'Version', 'Is_Deleted'];

  /** Row number of a record inside its sheet. Internal; stripped before it reaches a client. */
  var ROW_KEY = '_row';

  var headerCache = {};                   // tableName -> { headers, index, sheet }
  var recordCache = {};                   // tableName -> records (full-table reads only)
  var cachedDbId = null;

  /* ------------------------------------------------------------------ cache */

  function resetCache(tableName) {
    if (tableName) {
      delete headerCache[tableName];
      delete recordCache[tableName];
      return;
    }
    headerCache = {};
    recordCache = {};
    cachedDbId = null;
  }

  /** Drops every cache when the underlying spreadsheet changed (tests, or a re-setup). */
  function assertSameDatabase() {
    var id = Config.getDb().getId();
    if (cachedDbId !== id) {
      headerCache = {};
      recordCache = {};
      cachedDbId = id;
    }
  }

  function getHeaders(tableName) {
    assertSameDatabase();
    if (headerCache[tableName]) return headerCache[tableName];

    var table = Schema.getTable(tableName);
    var sheet = Config.getSheet(table.sheet);
    var lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      throw Err.internal('ชีต ' + table.sheet + ' ยังไม่มีหัวคอลัมน์ — ผู้ดูแลระบบต้องรัน setup()');
    }
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
      .map(function (h) { return String(h).trim(); });
    var index = {};
    headers.forEach(function (h, i) { if (h) index[h] = i; });

    headerCache[tableName] = { headers: headers, index: index, sheet: sheet };
    return headerCache[tableName];
  }

  function columnPosition(tableName, columnName) {
    var meta = getHeaders(tableName);
    var i = meta.index[columnName];
    if (i === undefined) {
      throw Err.internal('ชีต ' + Schema.getTable(tableName).sheet + ' ไม่มีคอลัมน์ ' +
        columnName + ' — ผู้ดูแลระบบต้องรัน setup()');
    }
    return i + 1;                         // sheets are 1-based
  }

  /* -------------------------------------------------------------- coercion */

  /** Sheet value -> the type the rest of the code expects. */
  function coerceRead(column, raw) {
    if (!column) return raw === null || raw === undefined ? '' : raw;
    switch (column.type) {
      case 'bool':
        return Utils.toBool(raw);
      case 'number':
      case 'int': {
        if (raw === '' || raw === null || raw === undefined) return null;
        var n = Number(raw);
        return isNaN(n) ? null : n;
      }
      case 'date':
      case 'datetime':
        return Utils.toDate(raw);
      default:
        return raw === null || raw === undefined ? '' : String(raw);
    }
  }

  /** Caller value -> what actually lands in the cell. Rejects malformed input. */
  function coerceWrite(tableName, column, value) {
    if (!column) return Utils.isBlank(value) ? '' : value;
    switch (column.type) {
      case 'bool':
        return Utils.toBool(value);
      case 'number':
      case 'int': {
        if (Utils.isBlank(value)) return '';
        var n = Number(value);
        if (isNaN(n)) {
          throw Err.validation('ค่าของ ' + column.name + ' ต้องเป็นตัวเลข', { field: column.name });
        }
        return column.type === 'int' ? Math.round(n) : n;
      }
      case 'date':
      case 'datetime': {
        if (Utils.isBlank(value)) return '';
        var d = Utils.toDate(value);
        if (!d) {
          throw Err.validation('ค่าของ ' + column.name + ' ไม่ใช่วันที่ที่ถูกต้อง', { field: column.name });
        }
        return column.type === 'date' ? Utils.startOfDay(d) : d;
      }
      default: {
        if (Utils.isBlank(value)) return '';
        var s = String(value);
        return column.max ? Utils.truncate(s, column.max) : s;
      }
    }
  }

  function rowToRecord(tableName, headers, values, rowNumber) {
    var record = {};
    headers.forEach(function (name, i) {
      if (!name) return;
      var column = Schema.getColumn(tableName, name);
      if (!column) return;                // a column in the sheet that Schema does not know
      record[name] = coerceRead(column, values[i]);
    });
    record[ROW_KEY] = rowNumber;
    return record;
  }

  /* ------------------------------------------------------------------ reads */

  /** Data rows in a table, without reading them. Used to pick a read strategy. */
  function rowCount(tableName) {
    var meta = getHeaders(tableName);
    return Math.max(0, meta.sheet.getLastRow() - 1);
  }

  /** Every row of a table. Use only for small tables or when no index applies. */
  function readAll(tableName) {
    assertSameDatabase();
    if (recordCache[tableName]) return recordCache[tableName];

    var meta = getHeaders(tableName);
    var lastRow = meta.sheet.getLastRow();
    var records = [];
    if (lastRow >= 2) {
      var values = meta.sheet.getRange(2, 1, lastRow - 1, meta.headers.length).getValues();
      for (var i = 0; i < values.length; i++) {
        var pk = Schema.getTable(tableName).pk;
        if (pk && Utils.isBlank(values[i][meta.index[pk]])) continue;   // trailing blank row
        records.push(rowToRecord(tableName, meta.headers, values[i], i + 2));
      }
    }
    recordCache[tableName] = records;
    return records;
  }

  /**
   * Row numbers whose `columnName` equals `value`, found with TextFinder so a
   * sheet with years of Activities in it does not have to be read in full
   * to answer a question about one Case (SPEC §11).
   */
  function findRowNumbers(tableName, columnName, value) {
    var meta = getHeaders(tableName);
    var lastRow = meta.sheet.getLastRow();
    if (lastRow < 2) return [];
    var position = columnPosition(tableName, columnName);
    return meta.sheet.getRange(2, position, lastRow - 1, 1)
      .createTextFinder(String(value))
      .matchEntireCell(true)
      .findAll()
      .map(function (r) { return r.getRow(); });
  }

  /** Reads the given row numbers in as few getValues() calls as the layout allows. */
  function readRowNumbers(tableName, rowNumbers) {
    if (rowNumbers.length === 0) return [];
    var meta = getHeaders(tableName);
    var records = [];
    Utils.groupRuns(rowNumbers).forEach(function (run) {
      var values = meta.sheet.getRange(run[0], 1, run[1] - run[0] + 1, meta.headers.length).getValues();
      values.forEach(function (v, i) {
        records.push(rowToRecord(tableName, meta.headers, v, run[0] + i));
      });
    });
    return records;
  }

  function matchesFilter(record, where) {
    if (!where) return true;
    if (typeof where === 'function') return where(record);
    for (var key in where) {
      if (!Object.prototype.hasOwnProperty.call(where, key)) continue;
      if (Utils.normalizeForCompare(record[key]) !== Utils.normalizeForCompare(where[key])) return false;
    }
    return true;
  }

  /**
   * opts: { where, includeDeleted, caseId, indexColumn, indexValue }
   * Passing caseId (or an explicit index column) uses the TextFinder path.
   */
  function query(tableName, opts) {
    var options = opts || {};
    var table = Schema.getTable(tableName);

    var records;
    var indexColumn = options.indexColumn;
    var indexValue = options.indexValue;
    if (!indexColumn && options.caseId && table.caseIdCol) {
      indexColumn = table.caseIdCol;
      indexValue = options.caseId;
    }

    if (indexColumn) {
      records = readRowNumbers(tableName, findRowNumbers(tableName, indexColumn, indexValue));
      // TextFinder matches on displayed text; confirm against the typed value.
      records = records.filter(function (r) {
        return Utils.normalizeForCompare(r[indexColumn]) === Utils.normalizeForCompare(indexValue);
      });
    } else {
      records = readAll(tableName);
    }

    var includeDeleted = !!options.includeDeleted;
    return records.filter(function (r) {
      if (table.audit && !includeDeleted && r.Is_Deleted === true) return false;
      return matchesFilter(r, options.where);
    });
  }

  function queryByCase(tableName, caseId, opts) {
    var options = opts || {};
    options.caseId = caseId;
    return query(tableName, options);
  }

  /**
   * One record by primary key, or null.
   *
   * Refuses to guess when two live rows carry the same key. Taking the first of
   * them looks harmless on a read and is not: the next update writes to whichever
   * row happened to come first, and a caller comparing ids reaches a row that is
   * not the one on screen. That is how a vendor who was never invited to a Case
   * came back as "already invited" — two Vendors rows were sharing one Vendor_ID
   * after the counter handed the same id out twice. A duplicate key is a broken
   * sheet, and the error says which rows to look at.
   *
   * A deleted row alongside a live one is not ambiguous, so the live one wins.
   */
  function findById(tableName, id, opts) {
    var options = opts || {};
    var table = Schema.getTable(tableName);
    if (!table.pk) throw Err.internal('ตาราง ' + tableName + ' ไม่มี primary key');
    if (Utils.isBlank(id)) return null;

    var matches = readRowNumbers(tableName, findRowNumbers(tableName, table.pk, id))
      .filter(function (r) { return Utils.normalizeForCompare(r[table.pk]) === Utils.normalizeForCompare(id); });
    if (matches.length === 0) return null;

    var records = matches;
    if (table.audit && !options.includeDeleted) {
      records = matches.filter(function (r) { return r.Is_Deleted !== true; });
      if (records.length === 0) return null;
    }

    if (records.length > 1) {
      throw Err.internal(
        '\u0e02\u0e49\u0e2d\u0e21\u0e39\u0e25\u0e02\u0e31\u0e14\u0e41\u0e22\u0e49\u0e07: \u0e23\u0e2b\u0e31\u0e2a ' + id + ' \u0e21\u0e35\u0e2d\u0e22\u0e39\u0e48 ' + records.length +
        ' \u0e41\u0e16\u0e27\u0e43\u0e19\u0e0a\u0e35\u0e15 ' + table.sheet + ' (\u0e41\u0e16\u0e27\u0e17\u0e35\u0e48 ' +
        records.map(function (r) { return r[ROW_KEY]; }).join(', ') +
        ') \u2014 \u0e23\u0e2b\u0e31\u0e2a\u0e19\u0e35\u0e49\u0e15\u0e49\u0e2d\u0e07\u0e44\u0e21\u0e48\u0e0b\u0e49\u0e33 \u0e01\u0e23\u0e38\u0e13\u0e32\u0e40\u0e01\u0e47\u0e1a\u0e02\u0e49\u0e2d\u0e21\u0e39\u0e25\u0e41\u0e25\u0e49\u0e27\u0e41\u0e01\u0e49\u0e43\u0e2b\u0e49\u0e40\u0e2b\u0e25\u0e37\u0e2d\u0e41\u0e16\u0e27\u0e40\u0e14\u0e35\u0e22\u0e27',
        { table: tableName, id: id, rows: records.map(function (r) { return r[ROW_KEY]; }) });
    }
    return records[0];
  }

  function requireById(tableName, id, opts) {
    var record = findById(tableName, id, opts);
    if (!record) throw Err.notFound('ไม่พบข้อมูล ' + tableName + ' รหัส ' + id, { table: tableName, id: id });
    return record;
  }

  /* ----------------------------------------------------------------- writes */

  function buildRow(tableName, meta, values) {
    var row = new Array(meta.headers.length);
    for (var i = 0; i < row.length; i++) row[i] = '';
    Object.keys(values).forEach(function (name) {
      var position = meta.index[name];
      if (position === undefined) return;
      row[position] = values[name];
    });
    return row;
  }

  /** Business fields the caller supplied, coerced and stripped of managed columns. */
  function sanitizePayload(tableName, payload) {
    var table = Schema.getTable(tableName);
    var clean = {};
    Object.keys(payload || {}).forEach(function (name) {
      if (name === ROW_KEY) return;
      if (MANAGED_COLUMNS.indexOf(name) !== -1) return;
      var column = Schema.getColumn(tableName, name);
      if (!column) return;                // unknown key from the client is ignored, never written
      if (name === table.pk && table.id) return;   // the system issues primary keys
      clean[name] = coerceWrite(tableName, column, payload[name]);
    });
    return clean;
  }

  /** Short, readable snapshot of a new record for the CREATE log entry. */
  function summarize(tableName, values) {
    var parts = [];
    Schema.getColumns(tableName).forEach(function (column) {
      if (Schema.isAuditColumn(column.name)) return;
      var v = values[column.name];
      if (Utils.isBlank(v)) return;
      parts.push(column.name + '=' + Utils.normalizeForCompare(v));
    });
    return parts.join('; ');
  }

  function resolveCaseId(tableName, values, opts) {
    if (opts && opts.caseId) return opts.caseId;
    var table = Schema.getTable(tableName);
    return table.caseIdCol ? values[table.caseIdCol] || '' : '';
  }

  /**
   * Inserts one record. opts: { actor, reason, caseId, id }
   * Pass `id` when the caller reserved the primary key beforehand, which lets it
   * build dependent resources (such as a Drive folder named after the Case) and
   * still store the record in a single version.
   * Returns the stored record, including its audit columns.
   */
  function insert(tableName, payload, opts) {
    var options = opts || {};
    if (options.id) {
      options = Object.assign({}, options, { ids: [options.id] });
    }
    return insertMany(tableName, [payload], options)[0];
  }

  /** Inserts many records with one id reservation, one sheet write and one log write. */
  function insertMany(tableName, payloads, opts) {
    if (!payloads || payloads.length === 0) return [];
    var options = opts || {};
    var table = Schema.getTable(tableName);

    return Utils.withScriptLock(function () {
      var meta = getHeaders(tableName);
      var now = Utils.now();
      var actor = options.actor || ChangeLog.SYSTEM_USER;
      var ids = [];
      if (table.id) {
        ids = options.ids && options.ids.length === payloads.length
          ? options.ids
          : IdGenerator.reserve(tableName, payloads.length);
      }

      var stored = [];
      var rows = [];
      payloads.forEach(function (payload, i) {
        var values = sanitizePayload(tableName, payload);
        if (table.pk) {
          values[table.pk] = table.id ? ids[i] : coerceWrite(tableName, Schema.getColumn(tableName, table.pk), payload[table.pk]);
        }
        if (table.audit) {
          values.Created_At = now;
          values.Created_By = actor;
          values.Updated_At = now;
          values.Updated_By = actor;
          values.Version = 1;
          values.Is_Deleted = false;
        }
        rows.push(buildRow(tableName, meta, values));
        stored.push(values);
      });

      var startRow = meta.sheet.getLastRow() + 1;
      meta.sheet.getRange(startRow, 1, rows.length, meta.headers.length).setValues(rows);
      stored.forEach(function (values, i) { values[ROW_KEY] = startRow + i; });

      if (!table.appendOnly) {
        ChangeLog.write(stored.map(function (values) {
          return {
            user: actor,
            table: tableName,
            recordId: table.pk ? values[table.pk] : '',
            caseId: resolveCaseId(tableName, values, options),
            action: ChangeLog.ACTIONS.CREATE,
            field: '',
            oldValue: '',
            newValue: summarize(tableName, values),
            reason: options.reason || ''
          };
        }));
      }

      resetCache(tableName);
      return stored;
    });
  }

  /**
   * Applies `patch` to one record.
   *
   * opts: { actor, reason, caseId, fieldActions, action }
   *   fieldActions — per-field Change_Log action override, e.g. { Status: 'STATUS_CHANGE' },
   *                  so an auditor sees STATUS_CHANGE rather than a generic UPDATE.
   *
   * Throws CONFLICT when expectedVersion does not match the stored Version (SPEC §6.3).
   * Passing null for expectedVersion skips the check and is reserved for system writes
   * such as the auto-revert in Rules.recheck.
   */
  function update(tableName, id, patch, expectedVersion, opts) {
    var options = opts || {};
    var table = Schema.getTable(tableName);
    if (table.appendOnly) throw Err.internal('ตาราง ' + tableName + ' แก้ไขไม่ได้');

    return Utils.withScriptLock(function () {
      var record = requireById(tableName, id, { includeDeleted: !!options.includeDeleted });
      assertVersion(table, record, expectedVersion);

      var changes = diff(tableName, record, sanitizePayload(tableName, patch));
      if (changes.length === 0) return record;

      writeChanges(tableName, record, changes, options);
      return requireById(tableName, id, { includeDeleted: true });
    });
  }

  /**
   * Applies several patches in one locked pass: one table read, contiguous row writes
   * and a single Change_Log append. Backs api_saveQuoteLines (SPEC §9).
   * items: [{ id, patch, version }]
   */
  function updateMany(tableName, items, opts) {
    if (!items || items.length === 0) return [];
    var options = opts || {};
    var table = Schema.getTable(tableName);

    return Utils.withScriptLock(function () {
      var updated = [];
      var entries = [];
      var meta = getHeaders(tableName);
      var now = Utils.now();
      var actor = options.actor || ChangeLog.SYSTEM_USER;

      items.forEach(function (item) {
        var record = requireById(tableName, item.id);
        assertVersion(table, record, item.version === undefined ? null : item.version);
        var changes = diff(tableName, record, sanitizePayload(tableName, item.patch));
        if (changes.length === 0) { updated.push(record); return; }

        var row = applyChangesToRow(tableName, meta, record, changes, actor, now);
        meta.sheet.getRange(record[ROW_KEY], 1, 1, meta.headers.length).setValues([row]);
        entries = entries.concat(changeEntries(tableName, record, changes, options, actor));
        updated.push(rowToRecord(tableName, meta.headers, row, record[ROW_KEY]));
      });

      if (entries.length) ChangeLog.write(entries);
      resetCache(tableName);
      return updated;
    });
  }

  function assertVersion(table, record, expectedVersion) {
    if (!table.audit || expectedVersion === null || expectedVersion === undefined || expectedVersion === '') return;
    if (Number(record.Version) !== Number(expectedVersion)) {
      throw Err.conflict('ข้อมูลถูกแก้ไขโดยผู้อื่น กรุณาโหลดใหม่', {
        expected: Number(expectedVersion),
        actual: Number(record.Version)
      });
    }
  }

  /** Fields whose value actually changed. Audit columns and equal values are skipped (SPEC §5.4). */
  function diff(tableName, record, cleanPatch) {
    var changes = [];
    Object.keys(cleanPatch).forEach(function (name) {
      if (Schema.isAuditColumn(name)) return;
      var oldValue = record[name];
      var newValue = cleanPatch[name];
      if (Utils.normalizeForCompare(oldValue) === Utils.normalizeForCompare(newValue)) return;
      changes.push({ field: name, oldValue: oldValue, newValue: newValue });
    });
    return changes;
  }

  function applyChangesToRow(tableName, meta, record, changes, actor, now) {
    var values = {};
    meta.headers.forEach(function (name) {
      if (!name) return;
      if (Schema.getColumn(tableName, name)) values[name] = record[name];
    });
    changes.forEach(function (c) { values[c.field] = c.newValue; });
    if (Schema.getTable(tableName).audit) {
      values.Updated_At = now;
      values.Updated_By = actor;
      values.Version = Number(record.Version || 0) + 1;
    }
    return buildRow(tableName, meta, values);
  }

  function changeEntries(tableName, record, changes, options, actor) {
    var table = Schema.getTable(tableName);
    var fieldActions = options.fieldActions || {};
    return changes.map(function (c) {
      return {
        user: actor,
        table: tableName,
        recordId: table.pk ? record[table.pk] : '',
        caseId: resolveCaseId(tableName, record, options),
        action: options.action || fieldActions[c.field] || ChangeLog.ACTIONS.UPDATE,
        field: c.field,
        oldValue: Utils.normalizeForCompare(c.oldValue),
        newValue: Utils.normalizeForCompare(c.newValue),
        reason: options.reason || ''
      };
    });
  }

  function writeChanges(tableName, record, changes, options) {
    var meta = getHeaders(tableName);
    var actor = options.actor || ChangeLog.SYSTEM_USER;
    var row = applyChangesToRow(tableName, meta, record, changes, actor, Utils.now());
    meta.sheet.getRange(record[ROW_KEY], 1, 1, meta.headers.length).setValues([row]);
    ChangeLog.write(changeEntries(tableName, record, changes, options, actor));
    resetCache(tableName);
  }

  /**
   * Soft delete (SPEC §2.5) — no row is ever removed from any sheet.
   * Cascades are the caller's job; ItemService and VendorService declare theirs.
   */
  function softDelete(tableName, id, expectedVersion, opts) {
    return setDeleted(tableName, id, true, expectedVersion, opts);
  }

  function restore(tableName, id, expectedVersion, opts) {
    return setDeleted(tableName, id, false, expectedVersion, opts);
  }

  function setDeleted(tableName, id, deleted, expectedVersion, opts) {
    var options = opts || {};
    var table = Schema.getTable(tableName);
    if (!table.audit) throw Err.internal('ตาราง ' + tableName + ' ไม่รองรับการลบแบบ soft delete');

    return Utils.withScriptLock(function () {
      var record = requireById(tableName, id, { includeDeleted: true });
      assertVersion(table, record, expectedVersion);
      if (record.Is_Deleted === deleted) return record;

      var meta = getHeaders(tableName);
      var actor = options.actor || ChangeLog.SYSTEM_USER;
      var changes = [{ field: 'Is_Deleted', oldValue: record.Is_Deleted, newValue: deleted }];
      var row = applyChangesToRow(tableName, meta, record, changes, actor, Utils.now());
      meta.sheet.getRange(record[ROW_KEY], 1, 1, meta.headers.length).setValues([row]);

      ChangeLog.writeOne({
        user: actor,
        table: tableName,
        recordId: record[table.pk],
        caseId: resolveCaseId(tableName, record, options),
        action: deleted ? ChangeLog.ACTIONS.DELETE : ChangeLog.ACTIONS.RESTORE,
        field: '',
        oldValue: deleted ? 'ACTIVE' : 'DELETED',
        newValue: deleted ? 'DELETED' : 'ACTIVE',
        reason: options.reason || ''
      });

      resetCache(tableName);
      return requireById(tableName, id, { includeDeleted: true });
    });
  }

  /** Soft-deletes every row of `tableName` whose `column` equals `value`. */
  function softDeleteWhere(tableName, column, value, opts) {
    var victims = query(tableName, { indexColumn: column, indexValue: value });
    victims.forEach(function (record) {
      softDelete(tableName, record[Schema.getTable(tableName).pk], null, opts);
    });
    return victims.length;
  }

  /** Strips internal keys and turns Dates into ISO strings for the client. */
  function toClient(record) {
    if (!record) return null;
    var out = {};
    Object.keys(record).forEach(function (key) {
      if (key.charAt(0) === '_') return;
      var v = record[key];
      out[key] = v instanceof Date ? v.toISOString() : v;
    });
    return out;
  }

  return {
    ROW_KEY: ROW_KEY,
    MANAGED_COLUMNS: MANAGED_COLUMNS,
    resetCache: resetCache,
    getHeaders: getHeaders,
    rowCount: rowCount,
    readAll: readAll,
    query: query,
    queryByCase: queryByCase,
    findById: findById,
    requireById: requireById,
    insert: insert,
    insertMany: insertMany,
    update: update,
    updateMany: updateMany,
    softDelete: softDelete,
    softDeleteWhere: softDeleteWhere,
    restore: restore,
    toClient: toClient,
    coerceWrite: coerceWrite
  };
})();
