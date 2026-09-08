/*
 * Fill these in AFTER IT completes the Entra app registration and tells you
 * where the add-in will be hosted. Everything else in the add-in reads from here.
 *
 * manifest.xml contains the same placeholders - keep the two in step.
 */
var OooConfig = {
  // Application (client) ID from the Entra app registration.
  clientId: '5173e29d-0a45-45bd-a734-474006421d42',

  // Directory (tenant) ID. 'common' works but pinning the tenant is tighter.
  tenantId: '17f2fc0e-75aa-464d-9428-9811d5baa84e',

  // HTTPS origin the add-in is served from, no trailing slash,
  // e.g. https://dashtech.github.io/signature-ooo
  hostUrl: 'https://ravee-acharya.github.io/outlook-ooo-signature',

  // Must exactly match a redirect URI registered on the app, character for
  // character. auth-end.html sits in src/ alongside this file - omitting that
  // segment yields a 404 that only surfaces mid sign-in.
  get redirectUri() { return this.hostUrl + '/src/auth-end.html'; },

  // Read-only, calendar-only. User.Read is just for showing whose account is linked.
  scopes: ['User.Read', 'Calendars.Read'],

  // Defaults for a new user; changeable in the task pane.
  defaults: {
    months: 3,
    timeZoneLabel: 'IST',
    timeZoneId: 'India Standard Time',
    heading: 'Upcoming Out of Office Days',
    color: '#FF0000'
  }
};

if (typeof module === 'object' && module.exports) { module.exports = OooConfig; }
