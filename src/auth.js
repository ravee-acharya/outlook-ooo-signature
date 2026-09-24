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

  // MSAL records an "interaction.status" lock in storage as soon as an
  // interactive flow starts, and clears it when that flow finishes. Closing the
  // sign-in dialog, or a redirect that never lands, leaves the lock behind - and
  // from then on EVERY sign-in fails with interaction_in_progress until it is
  // removed. There is no supported API to clear it, so remove it directly.
  function clearInteractionLocks() {
    var removed = 0;
    [typeof localStorage !== 'undefined' ? localStorage : null,
     typeof sessionStorage !== 'undefined' ? sessionStorage : null].forEach(function (store) {
      if (!store) { return; }
      try {
        var doomed = [];
        for (var i = 0; i < store.length; i++) {
          var k = store.key(i);
          if (k && k.indexOf('interaction.status') !== -1) { doomed.push(k); }
        }
        doomed.forEach(function (k) { try { store.removeItem(k); removed++; } catch (e) { } });
      } catch (e) { /* storage blocked */ }
    });
    return removed;
  }

  // A second sign-in started while the first is still running is itself a cause
  // of interaction_in_progress - two clicks on Sign in is enough. Share the one
  // in-flight attempt instead of starting another.
  var signInInFlight = null;

  function msalConfig() {
    return {
      auth: {
        clientId: OooConfig.clientId,
        authority: 'https://login.microsoftonline.com/' + (OooConfig.tenantId || 'common'),
        redirectUri: OooConfig.redirectUri,
        navigateToLoginRequestUrl: false
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

    if (signInInFlight) { return signInInFlight; }
    signInInFlight = doInteractiveSignIn(a)
      .then(function (t) { signInInFlight = null; return t; },
            function (e) { signInInFlight = null; throw e; });
    return signInInFlight;
  }

  async function doInteractiveSignIn(a) {
    // Always start from a clean slate: a lock left over from an abandoned
    // attempt would otherwise fail this one too.
    clearInteractionLocks();

    // Inside Outlook, go through the Office Dialog API rather than a popup.
    // Task panes block window.open, which surfaces as popup_window_error.
    if (isInOutlook()) {
      var viaDialog = await getTokenViaDialog();
      if (viaDialog && viaDialog.accessToken) {
        dialogToken = viaDialog.accessToken;
        dialogTokenAt = Date.now();
        dialogAccount = viaDialog.username || null;
        return dialogToken;
      }
      // Do NOT fall back to a popup here. Outlook blocks it, so the user would
      // get popup_window_error on top of whatever actually went wrong, with no
      // idea what to do. Report the real problem instead.
      clearInteractionLocks();
      var why = (viaDialog && viaDialog.error) ? (' (' + viaDialog.error + ')') : '';
      throw new Error('Sign-in did not complete' + why +
                      '. Close any sign-in window that is still open and try again.');
    }

    try {
      var res = await a.acquireTokenPopup({ scopes: OooConfig.scopes });
      if (res.account) { a.setActiveAccount(res.account); }
      return res.accessToken;
    } catch (e) {
      clearInteractionLocks();
      throw e;
    }
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
            // Pass failures through rather than collapsing them to null, so the
            // caller can tell the user what actually went wrong.
            finish(payload || { ok: false, error: 'no response from sign-in' });
          });

          // The user closed the dialog, or the host dismissed it. MSAL's lock is
          // still set in that case, so clear it before anyone tries again.
          dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
            clearInteractionLocks();
            finish({ ok: false, error: 'sign-in window was closed' +
                     (arg && arg.error ? ' [' + arg.error + ']' : '') });
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

    try {
      var a = await getApp();
      var acct = pickAccount(a);
      if (acct) {
        if (typeof a.clearCache === 'function') { await a.clearCache({ account: acct }); }
        else if (typeof a.removeAccount === 'function') { await a.removeAccount(acct); }
      } else if (typeof a.clearCache === 'function') {
        await a.clearCache();
      }
    } catch (e) { /* fall through to the explicit purge below */ }

    // A token acquired through the Office dialog may never have reached this
    // window's MSAL cache, so the calls above can find no account and clear
    // nothing. Whatever MSAL left behind is removed explicitly here - otherwise
    // the next silent refresh quietly signs the user back in and Sign out looks
    // like it did nothing.
    try {
      var doomed = [];
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (!k) { continue; }
        if (k.indexOf('msal') === 0 || k.indexOf('msal.') !== -1 ||
            (OooConfig && OooConfig.clientId && k.indexOf(OooConfig.clientId) !== -1)) {
          doomed.push(k);
        }
      }
      doomed.forEach(function (k) { try { localStorage.removeItem(k); } catch (e) { } });
    } catch (e) { /* private mode or storage blocked - nothing more we can do */ }

    app = null;
    usingNaa = false;
  }

  return {
    getToken: getToken,
    clearInteractionLocks: clearInteractionLocks,
    getSignedInAccount: getSignedInAccount,
    signOut: signOut,
    isUsingNaa: function () { return usingNaa; }
  };
}());
