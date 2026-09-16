/*
 * Company-wide holidays, published once and applied to everyone.
 *
 * The list lives at <host>/holidays.json, which the app owner edits directly on
 * GitHub. It is fetched same-origin, so no CORS and no credentials are involved,
 * and it holds nothing sensitive - just dates.
 *
 * Every consumer caches the parsed result in roaming settings, so the list keeps
 * working offline and inside the compose handler's tight time budget.
 */
var OooShared = (function () {
  'use strict';

  var CACHE_KEY = 'oooSharedHolidays';
  var CACHE_AT_KEY = 'oooSharedHolidaysAt';

  function url() {
    var base = (typeof OooConfig !== 'undefined' && OooConfig.hostUrl) ? OooConfig.hostUrl : '';
    // Bucketed by the hour: fresh enough for a list that changes a few times a
    // year, without defeating caching on every single compose.
    var bucket = Math.floor(Date.now() / 3600000);
    return base + '/holidays.json?h=' + bucket;
  }

  /**
   * @returns {Promise<{holidays:Array, updated:string}|null>} null on any failure
   */
  function fetchList() {
    if (typeof fetch !== 'function') { return Promise.resolve(null); }
    return fetch(url())
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (j) {
        if (!j || !Object.prototype.hasOwnProperty.call(j, 'holidays')) { return null; }
        if (!(j.holidays instanceof Array)) { return null; }
        return { holidays: j.holidays, updated: j.updated || '' };
      })
      .catch(function () { return null; });
  }

  function readCache(rs) {
    try {
      var raw = rs.get(CACHE_KEY);
      if (!raw) { return null; }
      var parsed = JSON.parse(raw);
      return (parsed && parsed.holidays instanceof Array) ? parsed : null;
    } catch (e) { return null; }
  }

  function writeCache(rs, list) {
    try {
      rs.set(CACHE_KEY, JSON.stringify(list));
      rs.set(CACHE_AT_KEY, new Date().toISOString());
      return true;
    } catch (e) { return false; }
  }

  // Turns whichever list we have into a day map for the current window.
  function toDays(list, opts) {
    if (!list || !(list.holidays instanceof Array)) { return {}; }
    return OooCore.daysFromHolidays(list.holidays, opts);
  }

  return {
    CACHE_KEY: CACHE_KEY,
    CACHE_AT_KEY: CACHE_AT_KEY,
    fetchList: fetchList,
    readCache: readCache,
    writeCache: writeCache,
    toDays: toDays
  };
}());
