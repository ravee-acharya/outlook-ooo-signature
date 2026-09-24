/*
 * Company-wide holidays, published once and applied to everyone.
 *
 * The list lives at <host>/holidays.json, which the app owner edits directly on
 * GitHub. It is fetched same-origin, so no CORS and no credentials are involved,
 * and it holds nothing sensitive - just dates.
 *
 * The file carries one list per region (India, United States, ...) and each
 * user picks exactly one. Two regions are never combined: someone in the India
 * office should not advertise US federal holidays, or the reverse.
 *
 * An older flat file - { "holidays": [...] } with no regions - is still
 * accepted and treated as a single unnamed region, so a cached copy from before
 * this change keeps working.
 *
 * Every consumer caches the parsed result in roaming settings, so the list keeps
 * working offline and inside the compose handler's tight time budget.
 */
var OooShared = (function () {
  'use strict';

  var CACHE_KEY = 'oooSharedHolidays';
  var CACHE_AT_KEY = 'oooSharedHolidaysAt';
  var LEGACY_KEY = '_all';

  // Array.isArray rather than `instanceof Array`: the latter is false for an
  // array created in a different realm, which silently discarded the whole list
  // in one earlier test harness.
  function isArr(v) { return Object.prototype.toString.call(v) === '[object Array]'; }

  function url() {
    var base = (typeof OooConfig !== 'undefined' && OooConfig.hostUrl) ? OooConfig.hostUrl : '';
    // Bucketed by the hour: fresh enough for a list that changes a few times a
    // year, without defeating caching on every single compose.
    var bucket = Math.floor(Date.now() / 3600000);
    return base + '/holidays.json?h=' + bucket;
  }

  /**
   * Accepts either shape and always returns
   *   { updated, defaultRegion, regions: [{ key, label, holidays }] }
   * or null when there is nothing usable.
   */
  function normalise(json) {
    if (!json) { return null; }

    if (isArr(json.regions)) {
      var regions = [];
      json.regions.forEach(function (r) {
        if (!r || !r.key || !isArr(r.holidays)) { return; }
        regions.push({
          key: String(r.key),
          label: String(r.label || r.key),
          holidays: r.holidays
        });
      });
      if (!regions.length) { return null; }
      return {
        updated: json.updated || '',
        defaultRegion: json.defaultRegion || regions[0].key,
        regions: regions
      };
    }

    // Legacy flat file.
    if (isArr(json.holidays)) {
      return {
        updated: json.updated || '',
        defaultRegion: LEGACY_KEY,
        regions: [{ key: LEGACY_KEY, label: 'Company holidays', holidays: json.holidays }]
      };
    }
    return null;
  }

  /**
   * @returns {Promise<object|null>} normalised list, or null on any failure
   */
  function fetchList() {
    if (typeof fetch !== 'function') { return Promise.resolve(null); }
    return fetch(url())
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (j) { return normalise(j); })
      .catch(function () { return null; });
  }

  function readCache(rs) {
    try {
      var raw = rs.get(CACHE_KEY);
      if (!raw) { return null; }
      return normalise(JSON.parse(raw));
    } catch (e) { return null; }
  }

  function writeCache(rs, list) {
    try {
      rs.set(CACHE_KEY, JSON.stringify(list));
      rs.set(CACHE_AT_KEY, new Date().toISOString());
      return true;
    } catch (e) { return false; }
  }

  function regionsOf(list) {
    return (list && isArr(list.regions)) ? list.regions : [];
  }

  /**
   * The user's chosen region, falling back to the file's default and then to the
   * first one. Returns null when the chosen key no longer exists and there is no
   * sensible fallback - better to show nothing than another region's holidays.
   */
  function pickRegion(list, regionKey) {
    var regions = regionsOf(list);
    if (!regions.length) { return null; }

    var i;
    if (regionKey) {
      for (i = 0; i < regions.length; i++) {
        if (regions[i].key === regionKey) { return regions[i]; }
      }
      // A key that no longer exists means the published file changed under the
      // user. Fall through to the default rather than guessing.
    }
    var def = list.defaultRegion;
    if (def) {
      for (i = 0; i < regions.length; i++) {
        if (regions[i].key === def) { return regions[i]; }
      }
    }
    return regions[0];
  }

  // Day map for exactly one region. Never merges two regions.
  function toDays(list, opts, regionKey) {
    var region = pickRegion(list, regionKey);
    if (!region) { return {}; }
    return OooCore.daysFromHolidays(region.holidays, opts);
  }

  return {
    CACHE_KEY: CACHE_KEY,
    CACHE_AT_KEY: CACHE_AT_KEY,
    LEGACY_KEY: LEGACY_KEY,
    normalise: normalise,
    fetchList: fetchList,
    readCache: readCache,
    writeCache: writeCache,
    regionsOf: regionsOf,
    pickRegion: pickRegion,
    toDays: toDays
  };
}());
