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

    var res = await a.acquireTokenPopup({ scopes: OooConfig.scopes });
    if (res.account) { a.setActiveAccount(res.account); }
    return res.accessToken;
  }

  async function getSignedInAccount() {
    var a = await getApp();
    var acct = pickAccount(a);
    return acct ? (acct.username || acct.name || null) : null;
  }

  async function signOut() {
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
