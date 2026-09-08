/*
 * Microsoft Graph calls. Read-only: calendarView and /me, nothing else.
 */
var OooGraph = (function () {
  'use strict';

  var BASE = 'https://graph.microsoft.com/v1.0';

  function pad(n) { return ('0' + n).slice(-2); }

  // Local wall-clock, no zone suffix - paired with the Prefer header below so
  // Graph interprets and returns everything in the user's own time zone.
  function graphLocal(d) {
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) +
           'T' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
  }

  async function call(url, token, timeZoneId) {
    var headers = { Authorization: 'Bearer ' + token };
    if (timeZoneId) { headers.Prefer = 'outlook.timezone="' + timeZoneId + '"'; }

    var res = await fetch(url, { headers: headers });
    if (!res.ok) {
      var detail = '';
      try {
        var body = await res.json();
        detail = (body.error && body.error.message) ? ': ' + body.error.message : '';
      } catch (e) { /* non-JSON error body */ }
      throw new Error('Graph ' + res.status + detail);
    }
    return res.json();
  }

  async function getMe(token) {
    var me = await call(BASE + '/me?$select=displayName,userPrincipalName,mail', token, null);
    var addr = me.mail || me.userPrincipalName;
    return { name: me.displayName || '', address: addr || '', label: me.displayName ? (me.displayName + ' <' + addr + '>') : addr };
  }

  /**
   * Events overlapping the window. Starts 90 days early so multi-day leave that
   * began before today is still caught and clipped, matching the desktop tool.
   */
  async function getCalendarEvents(token, from, to, timeZoneId) {
    var url = BASE + '/me/calendarView' +
      '?startDateTime=' + encodeURIComponent(graphLocal(OooCore.addDays(from, -90))) +
      '&endDateTime=' + encodeURIComponent(graphLocal(OooCore.addDays(to, 1))) +
      '&$select=subject,start,end,isAllDay,showAs,categories,isCancelled' +
      '&$top=200';

    var events = [];
    var guard = 0;
    while (url && guard++ < 50) {
      var page = await call(url, token, timeZoneId);
      events = events.concat(page.value || []);
      url = page['@odata.nextLink'];
    }
    return events;
  }

  return { getMe: getMe, getCalendarEvents: getCalendarEvents };
}());
