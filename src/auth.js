/*
 * Graph token acquisition for the task pane.
 *
 * Prefers Nested App Authentication (NAA), which lets an Office add-in get a
 * Graph token directly from the host with no middle-tier service and no popup.
 * Falls back to a normal MSAL popup where NAA is unavailable (older Outlook
 * builds), which is why the registration needs both redirect URIs.
 *
 * Deliberately only used by the task pane. The OnNewMessageCompose handler runs
 * under a short timeout and must never block on the network or a sign-in prompt,
 * so it reads pre-computed dates from roaming settings instead.
 */
var OooAuth = (function () {
  'use strict';

  var app = null;
  var usingNaa = false;

  // A token obtained through the Office dialog arrives as a bare string and is
  // not in this window's MSAL cache, so hold it here for the life of the pane.
  // Expiry is not tracked precisely - it is only used to avoid re-prompting for
  // every action in one sitting, and a stale one simply falls back to signing in.
  var dialogToken = null;
  var dialogTokenAt = 0;
  var dialogAccount = null;
  var DIALOG_TOKEN_TTL_MS = 45 * 60 * 1000;

  function cachedDialogToken() {
    if (!dialogToken) { return null; }
    if (Date.now() - dialogTokenAt > DIALOG_TOKEN_TTL_MS) { dialogToken = null; return null; }
    return dialogToken;
  }

  function msalConfig() {
    return {
      auth: {
        clientId: OooConfig.clientId,
        authority: 'https://login.microsoftonline.com/' + (OooConfig.tenantId || 'common'),
        redirectUri: OooConfig.redirectUri
      },
      cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false },
      system: { loggerOptions: { loggerCallback: function () {} } }
    };
  }

  async function getApp() {
    if (app) { return app; }
    if (typeof msal === 'undefined') {
      throw new Error('MSAL failed to load. Check that the CDN is reachable.');
    }
    var cfg = msalConfig();

    if (typeof msal.createNestablePublicClientApplication === 'function') {
      try {
        app = await msal.createNestablePublicClientApplication(cfg);
        usingNaa = true;
        return app;
      } catch (e) {
        // host does not support NAA - fall through to the standard flow
      }
    }
    app = new msal.PublicClientApplication(cfg);
    if (typeof app.initialize === 'function') { await app.initialize(); }
    usingNaa = false;
    return app;
  }

  function pickAccount(a) {
    return a.getActiveAccount() || (a.getAllAccounts() || [])[0] || null;
  }

  /**
   * @param {boolean} interactive allow a sign-in prompt if silent fails
   * @returns {Promise<string|null>} access token, or null when silent-only and not signed in
   */
  async function getToken(interactive) {
    var cached = cachedDialogToken();
    if (cached) { return cached; }

    var a = await getApp();
    var account = pickAccount(a);

    if (account) {
      try {
        var silent = await a.acquireTokenSilent({ scopes: OooConfig.scopes, account: account });
        return silent.accessToken;
      } catch (e) {
        if (!interactive) { return null; }
      }
    }
    if (!interactive) { return null; }

    // Inside Outlook, go through the Office Dialog API rather than a popup.
    // Task panes routinely block window.open, which surfaces as MSAL's
    // popup_window_error; an Office-managed dialog cannot be blocked.
    if (isInOutlook()) {
      var viaDialog = await getTokenViaDialog();
      if (viaDialog) {
        dialogToken = viaDialog.accessToken;
        dialogTokenAt = Date.now();
        dialogAccount = viaDialog.username || null;
        return dialogToken;
      }
    }

    var res = await a.acquireTokenPopup({ scopes: OooConfig.scopes });
    if (res.account) { a.setActiveAccount(res.account); }
    return res.accessToken;
  }

  function isInOutlook() {
    try {
      return typeof Office !== 'undefined' && Office.context && Office.context.ui &&
             typeof Office.context.ui.displayDialogAsync === 'function' &&
             !!Office.context.mailbox;
    } catch (e) { return false; }
  }

  // Opens auth-dialog.html in an Office dialog; that page redirects through
  // Entra and auth-end.html posts the token back with messageParent.
  function getTokenViaDialog() {
    return new Promise(function (resolve) {
      var url = OooConfig.hostUrl + '/src/auth-dialog.html';
      var dialog = null;
      var settled = false;

      function finish(payload) {
        if (settled) { return; }
        settled = true;
        try { if (dialog) { dialog.close(); } } catch (e) { /* already gone */ }
        resolve(payload || null);
      }

      Office.context.ui.displayDialogAsync(
        url,
        { height: 60, width: 30, promptBeforeOpen: false },
        function (result) {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            finish(null);   // fall back to the popup path
            return;
          }
          dialog = result.value;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
            var payload = null;
            try { payload = JSON.parse(arg.message); } catch (e) { payload = null; }
            finish(payload && payload.ok ? payload : null);
          });

          // Covers the user closing the dialog themselves.
          dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
            finish(null);
          });
        }
      );
    });
  }

  async function getSignedInAccount() {
    if (dialogAccount) { return dialogAccount; }
    var a = await getApp();
    var acct = pickAccount(a);
    return acct ? (acct.username || acct.name || null) : null;
  }

  async function signOut() {
    // Clear the dialog-held token too, otherwise signing out would appear to do
    // nothing for the rest of the session.
    dialogToken = null;
    dialogTokenAt = 0;
    dialogAccount = null;

    var a = await getApp();
    var acct = pickAccount(a);
    if (acct && typeof a.clearCache === 'function') {
      await a.clearCache({ account: acct });
    }
    app = null;
  }

  return {
    getToken: getToken,
    getSignedInAccount: getSignedInAccount,
    signOut: signOut,
    isUsingNaa: function () { return usingNaa; }
  };
}());
