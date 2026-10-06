/**
 * DraftService.js — what a buyer has typed but has not saved yet.
 *
 * A dialog's contents live in the page and nowhere else, so anything that takes
 * the page away takes the work with it: Chrome reclaiming a background tab, the
 * Apps Script frame re-establishing its session, a mis-aimed click on the
 * backdrop, a closed laptop. For a form with a dozen fields that means typing it
 * all again, which is what buyers were reporting.
 *
 * The draft is kept on the server rather than in the browser for two reasons.
 * The page itself is what disappears, so anything held in it disappears too; and
 * the Apps Script user-code frame is served from an origin whose subdomain can
 * change between loads, which makes localStorage and sessionStorage unreliable
 * for exactly the case this is meant to survive.
 *
 * A draft is not a save. Nothing here touches the database: it is a scratch note
 * in CacheService that expires on its own, and the Cases and Activities sheets
 * see nothing until the buyer presses the button.
 */
var DraftService = (function () {

  var KEY_PREFIX = 'draft_v1_';
  var TTL_SECONDS = 21600;                // six hours — the longest CacheService allows
  var MAX_JSON_LENGTH = 90 * 1024;        // CacheService refuses a value over 100 KB

  /** The forms allowed to keep a draft. A caller naming anything else is a bug. */
  var KINDS = ['newCase', 'activity'];

  function assertKind(kind) {
    if (KINDS.indexOf(kind) === -1) {
      throw Err.validation('ชนิดของร่างไม่ถูกต้อง', { field: 'kind', kind: String(kind) });
    }
  }

  /**
   * The cache is shared by everyone, so the owner's address is part of the key —
   * and it is the address the server resolved, never one the browser sent.
   *
   * CacheService.getUserCache() looks like the obvious answer and is the wrong
   * one here: this web app runs as the deploying account, so "the user" that
   * cache belongs to is that one account for every visitor, and one buyer's
   * half-written Case would be handed to the next.
   */
  function keyFor(user, kind, scope) {
    var suffix = Utils.isBlank(scope)
      ? ''
      : '_' + String(scope).replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40);
    return KEY_PREFIX + kind + suffix + '_' + String(user.email).trim().toLowerCase();
  }

  /**
   * Stores the draft. Returns false rather than throwing when it cannot be kept:
   * losing a draft is a missed convenience, and it must never be the thing that
   * stops somebody from filling in a form.
   */
  function save(user, kind, scope, values) {
    assertKind(kind);
    var json = JSON.stringify(values || {});
    if (json.length > MAX_JSON_LENGTH) return false;
    try {
      CacheService.getScriptCache().put(keyFor(user, kind, scope), json, TTL_SECONDS);
      return true;
    } catch (e) {
      console.error('saveDraft failed: ' + ((e && e.message) || e));
      return false;
    }
  }

  /** The stored draft, or null — including when the stored text is no longer JSON. */
  function load(user, kind, scope) {
    assertKind(kind);
    var raw;
    try {
      raw = CacheService.getScriptCache().get(keyFor(user, kind, scope));
    } catch (e) {
      return null;
    }
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch (e) {
      return null;
    }
  }

  function clear(user, kind, scope) {
    assertKind(kind);
    try {
      CacheService.getScriptCache().remove(keyFor(user, kind, scope));
    } catch (e) {
      // A draft that cannot be removed expires by itself within hours.
    }
    return true;
  }

  return {
    KINDS: KINDS,
    TTL_SECONDS: TTL_SECONDS,
    save: save,
    load: load,
    clear: clear
  };
})();
