/**
 * GasMocks.js — in-memory stand-ins for the Google Apps Script services the
 * server code uses, so tests/Tests.js can run under Node.
 *
 * Only the surface the application actually calls is implemented. The mocks are
 * deliberately strict (out-of-range writes throw, values round-trip as written)
 * so a test failing here means the application code is wrong, not the mock.
 */
'use strict';

const DEFAULT_MAX_ROWS = 1000;
const DEFAULT_MAX_COLUMNS = 26;

/* ------------------------------------------------------------------- Range */

class MockRange {
  constructor(sheet, row, column, numRows, numColumns) {
    if (row < 1 || column < 1) throw new Error('Range coordinates are 1-based');
    this._sheet = sheet;
    this._row = row;
    this._column = column;
    this._numRows = numRows;
    this._numColumns = numColumns;
  }

  getRow() { return this._row; }
  getColumn() { return this._column; }
  getNumRows() { return this._numRows; }
  getNumColumns() { return this._numColumns; }
  getA1Notation() { return `R${this._row}C${this._column}:R${this._row + this._numRows - 1}C${this._column + this._numColumns - 1}`; }

  getValues() {
    const out = [];
    for (let r = 0; r < this._numRows; r++) {
      const row = [];
      for (let c = 0; c < this._numColumns; c++) {
        row.push(this._sheet._get(this._row + r, this._column + c));
      }
      out.push(row);
    }
    return out;
  }

  getValue() { return this.getValues()[0][0]; }

  getDisplayValues() {
    return this.getValues().map((row) => row.map((v) => (v === '' || v === null || v === undefined ? '' : String(v))));
  }

  setValues(values) {
    if (!Array.isArray(values) || values.length !== this._numRows) {
      throw new Error(`setValues: expected ${this._numRows} rows, got ${Array.isArray(values) ? values.length : typeof values}`);
    }
    values.forEach((row, r) => {
      if (!Array.isArray(row) || row.length !== this._numColumns) {
        throw new Error(`setValues: row ${r} expected ${this._numColumns} columns, got ${Array.isArray(row) ? row.length : typeof row}`);
      }
      row.forEach((v, c) => this._sheet._set(this._row + r, this._column + c, v));
    });
    return this;
  }

  setValue(value) {
    for (let r = 0; r < this._numRows; r++) {
      for (let c = 0; c < this._numColumns; c++) this._sheet._set(this._row + r, this._column + c, value);
    }
    return this;
  }

  clearContent() { return this.setValue(''); }

  // Formatting is irrelevant to the tests but the application calls it.
  setFontWeight() { return this; }
  setBackground() { return this; }
  setNumberFormat() { return this; }
  setHorizontalAlignment() { return this; }

  createTextFinder(text) { return new MockTextFinder(this, text); }
}

/* -------------------------------------------------------------- TextFinder */

class MockTextFinder {
  constructor(range, text) {
    this._range = range;
    this._text = String(text);
    this._entireCell = false;
  }
  matchEntireCell(flag) { this._entireCell = !!flag; return this; }
  matchCase() { return this; }
  ignoreDiacritics() { return this; }
  useRegularExpression() { return this; }

  findAll() {
    const found = [];
    const values = this._range.getValues();
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const cell = values[r][c];
        const asText = cell === null || cell === undefined ? '' : String(cell);
        const hit = this._entireCell ? asText === this._text : asText.indexOf(this._text) !== -1;
        if (hit) {
          found.push(new MockRange(
            this._range._sheet,
            this._range.getRow() + r,
            this._range.getColumn() + c,
            1, 1
          ));
        }
      }
    }
    return found;
  }

  findNext() { return this.findAll()[0] || null; }
}

/* ------------------------------------------------------------------- Sheet */

class MockSheet {
  constructor(spreadsheet, name) {
    this._spreadsheet = spreadsheet;
    this._name = name;
    this._data = [];                       // _data[rowIndex][colIndex], 0-based
    this._maxRows = DEFAULT_MAX_ROWS;
    this._maxColumns = DEFAULT_MAX_COLUMNS;
    this._frozenRows = 0;
  }

  getName() { return this._name; }
  setName(name) { this._name = name; return this; }
  getSheetId() { return this._name; }

  _get(row, column) {
    const r = this._data[row - 1];
    if (!r) return '';
    const v = r[column - 1];
    return v === undefined ? '' : v;
  }

  _set(row, column, value) {
    if (row > this._maxRows) this._maxRows = row;
    if (column > this._maxColumns) {
      throw new Error(`Sheet "${this._name}": write to column ${column} exceeds ${this._maxColumns} columns`);
    }
    while (this._data.length < row) this._data.push([]);
    const r = this._data[row - 1];
    while (r.length < column) r.push('');
    r[column - 1] = value === undefined || value === null ? '' : value;
  }

  getLastRow() {
    for (let r = this._data.length; r >= 1; r--) {
      const row = this._data[r - 1] || [];
      if (row.some((v) => v !== '' && v !== null && v !== undefined)) return r;
    }
    return 0;
  }

  getLastColumn() {
    let last = 0;
    this._data.forEach((row) => {
      for (let c = row.length; c >= 1; c--) {
        const v = row[c - 1];
        if (v !== '' && v !== null && v !== undefined) { if (c > last) last = c; break; }
      }
    });
    return last;
  }

  getMaxRows() { return this._maxRows; }
  getMaxColumns() { return this._maxColumns; }

  insertColumnsAfter(afterColumn, howMany) { this._maxColumns = Math.max(this._maxColumns, afterColumn + howMany); return this; }
  insertRowsAfter(afterRow, howMany) { this._maxRows = Math.max(this._maxRows, afterRow + howMany); return this; }
  setFrozenRows(n) { this._frozenRows = n; return this; }
  getFrozenRows() { return this._frozenRows; }
  autoResizeColumn() { return this; }
  setColumnWidth() { return this; }

  getRange(row, column, numRows, numColumns) {
    return new MockRange(this, row, column, numRows === undefined ? 1 : numRows, numColumns === undefined ? 1 : numColumns);
  }

  getDataRange() {
    return new MockRange(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1));
  }

  appendRow(values) {
    const row = this.getLastRow() + 1;
    values.forEach((v, i) => this._set(row, i + 1, v));
    return this;
  }

  createTextFinder(text) { return new MockTextFinder(this.getDataRange(), text); }

  clear() { this._data = []; return this; }
}

/* ------------------------------------------------------------- Spreadsheet */

class MockSpreadsheet {
  constructor(id, name) {
    this._id = id;
    this._name = name;
    this._sheets = [];
    this._timeZone = 'Asia/Bangkok';
    this.insertSheet('Sheet1');
  }
  getId() { return this._id; }
  getName() { return this._name; }
  getUrl() { return `https://docs.google.com/spreadsheets/d/${this._id}/edit`; }
  getSheets() { return this._sheets.slice(); }
  getSheetByName(name) { return this._sheets.find((s) => s.getName() === name) || null; }
  insertSheet(name) {
    if (this.getSheetByName(name)) throw new Error(`Sheet "${name}" already exists`);
    const sheet = new MockSheet(this, name);
    this._sheets.push(sheet);
    return sheet;
  }
  deleteSheet(sheet) {
    const i = this._sheets.indexOf(sheet);
    if (i >= 0) this._sheets.splice(i, 1);
    return this;
  }
  setSpreadsheetTimeZone(tz) { this._timeZone = tz; return this; }
  getSpreadsheetTimeZone() { return this._timeZone; }
}

/* ------------------------------------------------------------------ Drive */

let driveSeq = 0;

class MockFile {
  constructor(name, mimeType, bytes) {
    this._id = `file_${++driveSeq}`;
    this._name = name;
    this._mimeType = mimeType;
    this._bytes = bytes || [];
    this.sharing = null;
  }
  getId() { return this._id; }
  getName() { return this._name; }
  getUrl() { return `https://drive.google.com/file/d/${this._id}/view`; }
  getMimeType() { return this._mimeType; }
  setSharing(access, permission) { this.sharing = { access, permission }; return this; }
}

class MockFolder {
  constructor(name) {
    this._id = `folder_${++driveSeq}`;
    this._name = name;
    this._folders = [];
    this._files = [];
    this.sharing = null;
  }
  getId() { return this._id; }
  getName() { return this._name; }
  getUrl() { return `https://drive.google.com/drive/folders/${this._id}`; }
  createFolder(name) {
    const f = new MockFolder(name);
    this._folders.push(f);
    MockDriveRegistry.folders[f.getId()] = f;
    return f;
  }
  createFile(blob) {
    const f = new MockFile(blob.getName(), blob.getContentType(), blob.getBytes());
    this._files.push(f);
    MockDriveRegistry.files[f.getId()] = f;
    return f;
  }
  getFoldersByName(name) {
    const matches = this._folders.filter((f) => f.getName() === name);
    let i = 0;
    return { hasNext: () => i < matches.length, next: () => matches[i++] };
  }
  getFiles() {
    let i = 0;
    return { hasNext: () => i < this._files.length, next: () => this._files[i++] };
  }
  setSharing(access, permission) { this.sharing = { access, permission }; return this; }
}

const MockDriveRegistry = { folders: {}, files: {} };

/* ----------------------------------------------------------------- Blob */

class MockBlob {
  constructor(bytes, contentType, name) {
    this._bytes = bytes;
    this._contentType = contentType;
    this._name = name;
  }
  getBytes() { return this._bytes; }
  getContentType() { return this._contentType; }
  getName() { return this._name; }
  setName(n) { this._name = n; return this; }
}

/* ---------------------------------------------------------- date formatting */

const DATE_PART_FORMATTERS = {};

function formatDate(date, timeZone, pattern) {
  const key = timeZone;
  if (!DATE_PART_FORMATTERS[key]) {
    DATE_PART_FORMATTERS[key] = new Intl.DateTimeFormat('en-GB', {
      timeZone,
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit',
      hour12: false
    });
  }
  const parts = {};
  DATE_PART_FORMATTERS[key].formatToParts(date).forEach((p) => { parts[p.type] = p.value; });
  if (parts.hour === '24') parts.hour = '00';
  // Longest tokens first so 'yyyy' is not eaten by 'yy'.
  return pattern
    .replace(/yyyy/g, parts.year)
    .replace(/MM/g, parts.month)
    .replace(/dd/g, parts.day)
    .replace(/HH/g, parts.hour)
    .replace(/mm/g, parts.minute)
    .replace(/ss/g, parts.second);
}

/* ------------------------------------------------------------- the globals */

/**
 * Builds a fresh set of Apps Script globals. Each test run gets its own, so no
 * state leaks between tests.
 */
function createGasEnvironment() {
  let spreadsheetSeq = 0;
  const spreadsheets = {};
  const scriptProperties = {};
  const cache = {};
  const sentMail = [];
  const triggers = [];
  let activeUserEmail = 'system@example.com';
  let effectiveUserEmail = 'system@example.com';
  let lockHeld = false;
  // Ordered record of lock and flush calls. The mock writes synchronously, so it
  // cannot reproduce Apps Script's write buffering; what it CAN check is that the
  // application flushes before it lets go of the lock.
  const lockEvents = [];

  const SpreadsheetApp = {
    create(name) {
      const id = `ss_${++spreadsheetSeq}`;
      const ss = new MockSpreadsheet(id, name);
      spreadsheets[id] = ss;
      return ss;
    },
    __all() { return spreadsheets; },
    openById(id) {
      if (!spreadsheets[id]) throw new Error(`No spreadsheet with id ${id}`);
      return spreadsheets[id];
    },
    getActiveSpreadsheet() { return null; },
    flush() { lockEvents.push('flush'); }
  };

  const propertyStore = {
    getProperty: (k) => (Object.prototype.hasOwnProperty.call(scriptProperties, k) ? scriptProperties[k] : null),
    setProperty: (k, v) => { scriptProperties[k] = String(v); return propertyStore; },
    deleteProperty: (k) => { delete scriptProperties[k]; return propertyStore; },
    getProperties: () => Object.assign({}, scriptProperties)
  };

  const PropertiesService = {
    getScriptProperties: () => propertyStore,
    getUserProperties: () => propertyStore
  };

  const CacheService = {
    getScriptCache() {
      return {
        get: (k) => (Object.prototype.hasOwnProperty.call(cache, k) ? cache[k] : null),
        put: (k, v) => { cache[k] = String(v); },
        remove: (k) => { delete cache[k]; },
        removeAll: (keys) => { keys.forEach((k) => delete cache[k]); }
      };
    },
    getUserCache() { return CacheService.getScriptCache(); }
  };

  const LockService = {
    getScriptLock() {
      return {
        tryLock: () => {
          // A single Node test run is single-threaded; a second acquisition while
          // held means the application failed to release, which must surface.
          if (lockHeld) return false;
          lockHeld = true;
          lockEvents.push('tryLock');
          return true;
        },
        waitLock: () => { lockHeld = true; lockEvents.push('waitLock'); },
        releaseLock: () => { lockHeld = false; lockEvents.push('releaseLock'); },
        hasLock: () => lockHeld
      };
    },
    getDocumentLock() { return LockService.getScriptLock(); },
    getUserLock() { return LockService.getScriptLock(); }
  };

  const Session = {
    getActiveUser: () => ({ getEmail: () => activeUserEmail }),
    getEffectiveUser: () => ({ getEmail: () => effectiveUserEmail }),
    getScriptTimeZone: () => 'Asia/Bangkok'
  };

  const Utilities = {
    formatDate,
    sleep: () => {},
    getUuid: () => `uuid-${Math.random().toString(36).slice(2)}`,
    newBlob: (bytes, contentType, name) => new MockBlob(bytes, contentType, name),
    base64Decode: (s) => Array.from(Buffer.from(String(s), 'base64')),
    base64Encode: (b) => Buffer.from(b).toString('base64')
  };

  const rootFolder = new MockFolder('My Drive');
  MockDriveRegistry.folders[rootFolder.getId()] = rootFolder;

  const DriveApp = {
    getRootFolder: () => rootFolder,
    createFolder: (name) => rootFolder.createFolder(name),
    getFolderById: (id) => {
      const f = MockDriveRegistry.folders[id];
      if (!f) throw new Error(`No folder with id ${id}`);
      return f;
    },
    getFileById: (id) => {
      const f = MockDriveRegistry.files[id];
      if (f) return f;
      const ss = spreadsheets[id];
      if (ss) {
        return {
          getId: () => id,
          getName: () => ss.getName(),
          getUrl: () => ss.getUrl(),
          setTrashed: (flag) => { ss._trashed = !!flag; if (flag) delete spreadsheets[id]; },
          setSharing: () => {}
        };
      }
      throw new Error(`No file with id ${id}`);
    },
    Access: { DOMAIN: 'DOMAIN', DOMAIN_WITH_LINK: 'DOMAIN_WITH_LINK', PRIVATE: 'PRIVATE' },
    Permission: { VIEW: 'VIEW', EDIT: 'EDIT', NONE: 'NONE' }
  };

  const MailApp = {
    sendEmail(arg) { sentMail.push(arg); },
    getRemainingDailyQuota: () => 1000
  };

  const ScriptApp = {
    getService: () => ({ getUrl: () => 'https://script.google.com/a/macros/example.com/s/TEST/exec' }),
    getProjectTriggers: () => triggers.slice(),
    deleteTrigger: (t) => {
      const i = triggers.indexOf(t);
      if (i >= 0) triggers.splice(i, 1);
    },
    newTrigger(handler) {
      const spec = { handler, hour: null };
      const builder = {
        timeBased: () => builder,
        atHour: (h) => { spec.hour = h; return builder; },
        everyDays: () => builder,
        everyHours: () => builder,
        create: () => {
          const trigger = {
            getHandlerFunction: () => spec.handler,
            getUniqueId: () => `trigger_${triggers.length + 1}`,
            _spec: spec
          };
          triggers.push(trigger);
          return trigger;
        }
      };
      return builder;
    }
  };

  // HTML files as Apps Script sees them: a flat map of name -> content, where the
  // name carries slashes. Unknown names throw, exactly as the real service does,
  // which is what lets the tests exercise the name resolver.
  const htmlFiles = {};
  let globalScope = null;

  function htmlEscape(value) {
    return String(value === null || value === undefined ? '' : value)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function readHtmlFile(name) {
    if (!Object.prototype.hasOwnProperty.call(htmlFiles, name)) {
      throw new Error(`No HTML file named ${name}`);
    }
    return htmlFiles[name];
  }

  /**
   * Evaluates the scriptlet subset this project actually uses:
   *   <?!= include('name'); ?>   printed unescaped, via the app's own include()
   *   <?= property ?>            printed escaped, from the template's properties
   * Anything else is left alone; this is a stand-in, not a template engine.
   */
  function evaluateTemplate(content, properties) {
    return String(content)
      .replace(/<\?!=\s*include\(\s*['"]([^'"]+)['"]\s*\)\s*;?\s*\?>/g, (match, name) => {
        if (!globalScope || typeof globalScope.include !== 'function') {
          throw new Error('include() is not defined in the loaded script');
        }
        return globalScope.include(name);
      })
      .replace(/<\?=\s*([A-Za-z_$][\w$]*)\s*\?>/g, (match, property) =>
        htmlEscape(properties[property]));
  }

  const HtmlService = {
    createTemplateFromFile(name) {
      const content = readHtmlFile(name);
      const template = {
        evaluate() { return HtmlService.createHtmlOutput(evaluateTemplate(content, template)); }
      };
      return template;
    },
    createHtmlOutputFromFile(name) {
      return HtmlService.createHtmlOutput(readHtmlFile(name));
    },
    createHtmlOutput: (content) => ({
      _content: content || '',
      getContent() { return this._content; },
      setTitle() { return this; },
      setXFrameOptionsMode() { return this; },
      addMetaTag() { return this; }
    }),
    XFrameOptionsMode: { ALLOWALL: 'ALLOWALL', DEFAULT: 'DEFAULT' }
  };

  return {
    SpreadsheetApp,
    PropertiesService,
    CacheService,
    LockService,
    Session,
    Utilities,
    DriveApp,
    MailApp,
    ScriptApp,
    HtmlService,
    console,
    // Test-side handles the suite uses to inspect or steer the environment.
    __test: {
      setActiveUser(email) { activeUserEmail = email; effectiveUserEmail = email; },
      /** The runner hands the mock the loaded script's globals so include() can be called back. */
      setGlobalScope(scope) { globalScope = scope; },
      /** Replaces the HTML file table, to simulate a different clasp naming scheme. */
      setHtmlFiles(files) {
        Object.keys(htmlFiles).forEach((k) => delete htmlFiles[k]);
        Object.keys(files).forEach((k) => { htmlFiles[k] = files[k]; });
      },
      htmlFileNames() { return Object.keys(htmlFiles); },
      lockEvents,
      clearLockEvents() { lockEvents.length = 0; },
      getActiveUser() { return activeUserEmail; },
      sentMail,
      clearMail() { sentMail.length = 0; },
      triggers,
      scriptProperties,
      cache,
      clearCache() { Object.keys(cache).forEach((k) => delete cache[k]); },
      isLockHeld() { return lockHeld; },
      drive: MockDriveRegistry,
      rootFolder
    }
  };
}

module.exports = { createGasEnvironment, MockSpreadsheet, MockSheet, MockRange };
