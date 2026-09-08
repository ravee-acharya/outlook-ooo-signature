/*
 * Task pane: sign in, read the calendar, and store the result in roaming
 * settings for the compose-time handler to use.
 *
 * Roaming settings live in the mailbox, so what is saved here follows the user
 * to every device and every Outlook client without any local state.
 */
(function () {
  'use strict';

  var STORE = {
    days: 'oooDays',
    signature: 'signatureHtml',
    options: 'oooOptions',
    enabled: 'oooEnabled',
    refreshed: 'oooRefreshed',
    lastEvent: 'oooLastEvent',
    account: 'oooAccount'
  };

  // Outlook caps roaming settings at 32 KB for the whole bag.
  var SIGNATURE_WARN = 20000;

  var $ = function (id) { return document.getElementById(id); };
  var rs = null;

  function setStatus(msg, kind) {
    var el = $('status');
    el.textContent = msg || '';
    el.className = 'status' + (kind ? ' ' + kind : '');
  }

  function currentOptions() {
    return {
      months: Math.max(1, Math.min(12, parseInt($('months').value, 10) || 3)),
      timeZoneLabel: $('tz').value.trim() || 'IST',
      heading: $('heading').value.trim() || OooConfig.defaults.heading,
      color: $('color').value.trim() || OooConfig.defaults.color
    };
  }

  function loadDays() {
    try { return JSON.parse(rs.get(STORE.days) || '{}'); } catch (e) { return {}; }
  }

  function renderPreview() {
    var days = loadDays();
    var opts = currentOptions();
    var html = '';
    try { html = OooCore.blockFromDays(days, opts); } catch (e) { html = ''; }

    $('preview').innerHTML = html ||
      '<span class="muted">No upcoming days off &#8212; nothing will be added.</span>';

    var kept = Object.keys(OooCore.clipDays(days, opts)).length;
    var when = rs.get(STORE.refreshed);
    var bits = [kept + ' day(s) in the next ' + opts.months + ' month(s)'];
    if (when) { bits.push('last refreshed ' + new Date(when).toLocaleString()); }
    $('meta').textContent = bits.join(' · ');

    // The compose runtime has no console, so it records each run here.
    // This is the only way to see why a new message came up without the block.
    var ev = null;
    try { ev = JSON.parse(rs.get(STORE.lastEvent) || 'null'); } catch (e) { ev = null; }
    var el = $('lastevent');
    // Guard: during a redeploy Outlook can serve cached HTML that predates this
    // element. Throwing here would abort restore() and leave the pane blank -
    // a diagnostic must never be able to break the UI it reports into.
    if (!el) { return; }
    if (ev && ev.at) {
      el.textContent = 'Last new-message event: ' + ev.outcome +
                       (ev.detail ? ' (' + ev.detail + ')' : '') +
                       ' at ' + new Date(ev.at).toLocaleString();
      el.className = 'muted small' + (ev.outcome === 'inserted' ? '' : ' err');
    } else {
      el.textContent = 'Last new-message event: none recorded yet.';
      el.className = 'muted small';
    }
  }

  function updateSigSize() {
    var n = $('signature').value.length;
    var el = $('sigsize');
    el.textContent = n ? (n + ' characters') : '';
    el.className = 'muted small' + (n > SIGNATURE_WARN ? ' err' : '');
    if (n > SIGNATURE_WARN) {
      el.textContent = n + ' characters — too large, Outlook may refuse to save. ' +
                       'Use image URLs instead of embedded images.';
    }
  }

  function saveSettings(cb) {
    rs.set(STORE.options, JSON.stringify(currentOptions()));
    rs.set(STORE.signature, $('signature').value);
    rs.set(STORE.enabled, $('enabled').checked);
    rs.saveAsync(function (res) {
      // Compare loosely: the local shim reports a plain 'succeeded' string and
      // Office.AsyncResultStatus is unavailable when running outside Outlook.
      var ok = res && (res.status === 'succeeded' ||
                       (typeof Office !== 'undefined' && Office.AsyncResultStatus &&
                        res.status === Office.AsyncResultStatus.Succeeded));
      if (ok) { cb(null); }
      else { cb(new Error((res && res.error && res.error.message) || 'Could not save settings.')); }
    });
  }

  async function refreshFromCalendar(interactive) {
    setStatus('Signing in…');
    var token = await OooAuth.getToken(interactive);
    if (!token) {
      setStatus('Sign in to read your calendar.', null);
      return false;
    }

    setStatus('Reading your calendar…');
    var opts = currentOptions();
    var today = OooCore.dateOnly(new Date());
    var windowEnd = OooCore.addMonths(today, opts.months);

    var events = await OooGraph.getCalendarEvents(
      token, today, windowEnd, OooConfig.defaults.timeZoneId);

    var days = OooCore.daysFromEvents(events, {
      today: today, months: opts.months, timeZoneLabel: opts.timeZoneLabel
    });

    var me = null;
    try { me = await OooGraph.getMe(token); } catch (e) { /* non-fatal */ }

    rs.set(STORE.days, JSON.stringify(days));
    rs.set(STORE.refreshed, new Date().toISOString());
    if (me) { rs.set(STORE.account, me.label); }

    await new Promise(function (resolve, reject) {
      saveSettings(function (err) { err ? reject(err) : resolve(); });
    });

    if (me) { $('account').textContent = 'Linked account: ' + me.label; }
    renderPreview();
    setStatus('Updated from your calendar.', 'ok');
    return true;
  }

  function wire() {
    $('refresh').addEventListener('click', function () {
      $('refresh').disabled = true;
      refreshFromCalendar(true)
        .catch(function (e) { setStatus(e.message || String(e), 'err'); })
        .then(function () { $('refresh').disabled = false; });
    });

    $('signin').addEventListener('click', function () {
      refreshFromCalendar(true).catch(function (e) {
        setStatus(e.message || String(e), 'err');
      });
    });

    $('save').addEventListener('click', function () {
      saveSettings(function (err) {
        if (err) { setStatus(err.message, 'err'); return; }
        renderPreview();
        setStatus('Saved.', 'ok');
      });
    });

    $('signout').addEventListener('click', function () {
      OooAuth.signOut().then(function () {
        $('account').textContent = 'Not signed in';
        setStatus('Signed out. Your stored dates are unchanged.', null);
      });
    });

    ['months', 'tz', 'heading', 'color'].forEach(function (id) {
      $(id).addEventListener('input', renderPreview);
    });
    $('signature').addEventListener('input', updateSigSize);
  }

  function restore() {
    var opts = {};
    try { opts = JSON.parse(rs.get(STORE.options) || '{}'); } catch (e) { opts = {}; }

    $('months').value = opts.months || OooConfig.defaults.months;
    $('tz').value = opts.timeZoneLabel || OooConfig.defaults.timeZoneLabel;
    $('heading').value = opts.heading || OooConfig.defaults.heading;
    $('color').value = opts.color || OooConfig.defaults.color;
    $('signature').value = rs.get(STORE.signature) || '';
    $('enabled').checked = rs.get(STORE.enabled) !== false;

    var acct = rs.get(STORE.account);
    if (acct) { $('account').textContent = 'Linked account: ' + acct; }

    updateSigSize();
    renderPreview();
  }

  // Opened directly in a browser rather than inside Outlook there is no mailbox,
  // so roaming settings do not exist. Rather than sit on "Loading..." forever,
  // fall back to a localStorage-backed stand-in. Sign-in, the calendar read and
  // the preview can then all be exercised outside Outlook, which is the only way
  // to test the Graph side before the add-in is deployed.
  function makeLocalShim() {
    var KEY = 'oooLocalSettings';
    var bag = {};
    try { bag = JSON.parse(localStorage.getItem(KEY) || '{}'); } catch (e) { bag = {}; }
    return {
      get: function (k) { return Object.prototype.hasOwnProperty.call(bag, k) ? bag[k] : undefined; },
      set: function (k, v) { bag[k] = v; },
      remove: function (k) { delete bag[k]; },
      saveAsync: function (cb) {
        try {
          localStorage.setItem(KEY, JSON.stringify(bag));
          cb({ status: 'succeeded' });
        } catch (e) {
          cb({ status: 'failed', error: { message: e.message } });
        }
      }
    };
  }

  function showBrowserBanner() {
    var b = document.createElement('p');
    b.className = 'status err';
    b.style.margin = '0 0 10px';
    b.textContent = 'Running outside Outlook — signing in and the calendar preview work, ' +
                    'but settings are saved to this browser only, not to your mailbox.';
    $('app').insertBefore(b, $('app').firstChild);
  }

  var started = false;

  function boot(isOutlook) {
    if (started) { return; }
    started = true;

    rs = isOutlook ? Office.context.roamingSettings : makeLocalShim();

    $('boot').hidden = true;
    $('app').hidden = false;
    if (!isOutlook) { showBrowserBanner(); }

    // Never let a rendering fault leave the user staring at an empty pane.
    try { wire(); } catch (e) { setStatus('UI error: ' + (e.message || e), 'err'); }
    try { restore(); } catch (e) { setStatus('Could not load saved settings: ' + (e.message || e), 'err'); }

    if (String(OooConfig.clientId).indexOf('__') === 0) {
      setStatus('Not configured yet: config.js still has placeholder values.', 'err');
      return;
    }

    // Quietly top up if we already have a usable token; never prompt on open.
    refreshFromCalendar(false).catch(function () { /* stay silent on load */ });
  }

  Office.onReady(function (info) {
    boot(!!(info && info.host === Office.HostType.Outlook));
  });

  // Office.js can fail to resolve at all when the page is loaded outside an
  // Office host; without this the pane would hang on "Loading..." indefinitely.
  setTimeout(function () { boot(false); }, 4000);
}());
