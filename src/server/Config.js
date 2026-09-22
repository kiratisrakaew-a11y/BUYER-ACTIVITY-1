/**
 * Config.js — access to the DB spreadsheet and to the three config sheets.
 *
 * Nothing in the system hardcodes a value that lives in Config_Settings,
 * Config_Lists or Status_Master (SPEC §2.7). Reads are cached in CacheService
 * for CACHE_TTL_SECONDS and memoised for the duration of one execution;
 * an admin clears both through api_clearCache.
 */
var Config = (function () {

  var CACHE_TTL_SECONDS = 600;            // SPEC §11 — 10 minutes
  var CACHE_PREFIX = 'bpa_v1_';
  var DB_PROPERTY_KEY = 'DB_SPREADSHEET_ID';

  /** Used when a key is missing from Config_Settings (SPEC §5.3). */
  var DEFAULT_SETTINGS = {
    BUYER_CAN_VIEW_ALL: 'TRUE',
    DRIVE_ROOT_FOLDER_ID: '',
    REMINDER_HOUR: '8',
    REMINDER_DAYS_AHEAD: '1',
    APP_TIMEZONE: 'Asia/Bangkok'
  };

  var dbOverride = null;                  // set by Tests.js to point at a scratch spreadsheet
  var dbMemo = null;                      // openById() is an API call; do it once per execution
  var memo = {};                          // per-execution memo, never outlives the request

  /* ------------------------------------------------------------ spreadsheet */

  /** Tests call this to run against a throwaway spreadsheet instead of the real DB. */
  function __setDbOverride(ss) {
    dbOverride = ss;
    dbMemo = null;
    resetMemo();
  }

  function __clearDbOverride() {
    dbOverride = null;
    dbMemo = null;
    resetMemo();
  }

  function getDb() {
    if (dbOverride) return dbOverride;
    if (dbMemo) return dbMemo;
    var id = PropertiesService.getScriptProperties().getProperty(DB_PROPERTY_KEY);
    if (!id) {
      throw Err.internal('ยังไม่ได้ตั้งค่าไฟล์ฐานข้อมูล — ผู้ดูแลระบบต้องรัน setup() ก่อน');
    }
    dbMemo = SpreadsheetApp.openById(id);
    return dbMemo;
  }

  function getSheet(sheetName) {
    var sheet = getDb().getSheetByName(sheetName);
    if (!sheet) {
      throw Err.internal('ไม่พบชีต "' + sheetName + '" — ผู้ดูแลระบบต้องรัน setup() ก่อน');
    }
    return sheet;
  }

  /** Whole sheet including the header row, as a 2D array. Empty sheet → []. */
  function readSheetValues(sheetName) {
    var sheet = getSheet(sheetName);
    if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) return [];
    return sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  }

  /* ------------------------------------------------------------------ cache */

  function resetMemo() {
    memo = {};
    dbMemo = null;
  }

  function cacheGet(key) {
    if (Object.prototype.hasOwnProperty.call(memo, key)) return memo[key];
    var raw = null;
    try {
      raw = CacheService.getScriptCache().get(CACHE_PREFIX + key);
    } catch (e) {
      raw = null;                         // cache is an optimisation, never a hard dependency
    }
    if (raw === null || raw === undefined) return undefined;
    var parsed;
    try {
      parsed = JSON.parse(raw);
    } catch (e) {
      return undefined;
    }
    memo[key] = parsed;
    return parsed;
  }

  function cachePut(key, value) {
    memo[key] = value;
    try {
      CacheService.getScriptCache().put(CACHE_PREFIX + key, JSON.stringify(value), CACHE_TTL_SECONDS);
    } catch (e) {
      // A value too large for the cache is not an error; the next read re-reads the sheet.
    }
  }

  function clearCache() {
    resetMemo();
    try {
      CacheService.getScriptCache().removeAll(
        ['settings', 'lists', 'status'].map(function (k) { return CACHE_PREFIX + k; })
      );
    } catch (e) {
      // ignore
    }
  }

  /* --------------------------------------------------------------- settings */

  function getSettings() {
    var cached = cacheGet('settings');
    if (cached !== undefined) return cached;

    var values = readSheetValues('Config_Settings');
    var settings = {};
    for (var k in DEFAULT_SETTINGS) {
      if (Object.prototype.hasOwnProperty.call(DEFAULT_SETTINGS, k)) settings[k] = DEFAULT_SETTINGS[k];
    }
    for (var i = 1; i < values.length; i++) {
      var key = String(values[i][0] || '').trim();
      if (!key) continue;
      settings[key] = values[i][1] === null || values[i][1] === undefined ? '' : String(values[i][1]).trim();
    }
    cachePut('settings', settings);
    return settings;
  }

  function get(key, fallback) {
    var settings = getSettings();
    var value = settings[key];
    if (value === undefined || value === '') {
      return fallback !== undefined ? fallback : (DEFAULT_SETTINGS[key] || '');
    }
    return value;
  }

  function getNumber(key, fallback) {
    var n = Number(get(key, fallback));
    return isNaN(n) ? Number(fallback) : n;
  }

  function getBool(key, fallback) {
    var v = String(get(key, fallback)).trim().toUpperCase();
    if (v === 'TRUE' || v === 'YES' || v === '1') return true;
    if (v === 'FALSE' || v === 'NO' || v === '0') return false;
    return !!fallback;
  }

  function getTimezone() {
    return get('APP_TIMEZONE', 'Asia/Bangkok');
  }

  /**
   * Writes a setting back to Config_Settings. Used only for values the system
   * discovers for itself, such as the Drive root folder it had to create.
   */
  function setSetting(key, value, description) {
    var sheet = getSheet('Config_Settings');
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0]).trim() === key) {
          sheet.getRange(i + 2, 2).setValue(value);
          clearCache();
          return value;
        }
      }
    }
    sheet.appendRow([key, value, description || '']);
    clearCache();
    return value;
  }

  /* ------------------------------------------------------------------ lists */

  /** All active Config_Lists rows grouped by List_Name and sorted by Sort_Order. */
  function getAllLists() {
    var cached = cacheGet('lists');
    if (cached !== undefined) return cached;

    var values = readSheetValues('Config_Lists');
    var lists = {};
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var listName = String(row[0] || '').trim();
      var code = String(row[1] || '').trim();
      if (!listName || !code) continue;
      var isActive = row[5];
      if (isActive === false || String(isActive).trim().toUpperCase() === 'FALSE') continue;
      if (!lists[listName]) lists[listName] = [];
      lists[listName].push({
        code: code,
        label: String(row[2] || code),
        parent: String(row[3] || '').trim(),
        sort: Number(row[4]) || 0
      });
    }
    Object.keys(lists).forEach(function (name) {
      lists[name].sort(function (a, b) { return a.sort - b.sort; });
    });
    cachePut('lists', lists);
    return lists;
  }

  /** One list, optionally filtered by Parent_Code (e.g. SUB_TYPE under a Budget_Type). */
  function getList(listName, parentCode) {
    var items = getAllLists()[listName] || [];
    if (parentCode === undefined || parentCode === null || parentCode === '') return items;
    return items.filter(function (it) { return !it.parent || it.parent === parentCode; });
  }

  function isValidCode(listName, code, parentCode) {
    return getList(listName, parentCode).some(function (it) { return it.code === code; });
  }

  /** Thai label for a code, falling back to the code itself, as the client does. */
  function labelOf(listName, code) {
    if (!code) return '';
    var items = getList(listName);
    for (var i = 0; i < items.length; i++) {
      if (items[i].code === code) return items[i].label;
    }
    return String(code);
  }

  /* ---------------------------------------------------------- status master */

  function getStatusMaster() {
    var cached = cacheGet('status');
    if (cached !== undefined) return cached;

    var values = readSheetValues('Status_Master');
    var statuses = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var code = String(row[0] || '').trim();
      if (!code) continue;
      statuses.push({
        code: code,
        label: String(row[1] || code),
        sequence: Number(row[2]) || 0,
        module: String(row[3] || '').trim(),
        isTerminal: row[4] === true || String(row[4]).trim().toUpperCase() === 'TRUE',
        allowedNext: String(row[5] || '')
          .split(',')
          .map(function (s) { return s.trim(); })
          .filter(function (s) { return !!s; })
      });
    }
    statuses.sort(function (a, b) { return a.sequence - b.sequence; });
    cachePut('status', statuses);
    return statuses;
  }

  function getStatus(code) {
    var all = getStatusMaster();
    for (var i = 0; i < all.length; i++) {
      if (all[i].code === code) return all[i];
    }
    return null;
  }

  return {
    CACHE_TTL_SECONDS: CACHE_TTL_SECONDS,
    DB_PROPERTY_KEY: DB_PROPERTY_KEY,
    DEFAULT_SETTINGS: DEFAULT_SETTINGS,
    __setDbOverride: __setDbOverride,
    __clearDbOverride: __clearDbOverride,
    getDb: getDb,
    getSheet: getSheet,
    readSheetValues: readSheetValues,
    resetMemo: resetMemo,
    clearCache: clearCache,
    getSettings: getSettings,
    get: get,
    getNumber: getNumber,
    getBool: getBool,
    getTimezone: getTimezone,
    setSetting: setSetting,
    getAllLists: getAllLists,
    getList: getList,
    isValidCode: isValidCode,
    labelOf: labelOf,
    getStatusMaster: getStatusMaster,
    getStatus: getStatus
  };
})();
