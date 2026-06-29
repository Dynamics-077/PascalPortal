/* pp-auth.js — shared token refresh helper
   Include AFTER msal-browser.min.js on every portal page.
   Provides: getValidToken(), authHeaders(), scheduleTokenRefresh()
*/
(function () {
  const APP_CLIENT_ID = '6bc856af-36d8-424d-a089-1860d402627b';
  const TENANT_ID     = '132fee41-6bf5-4f91-be3e-5c3b2a2fb1b8';
  const REFRESH_BEFORE_MS = 5 * 60 * 1000; // refresh 5 min before expiry

  const msalConfig = {
    auth: {
      clientId:    APP_CLIENT_ID,
      authority:   'https://login.microsoftonline.com/' + TENANT_ID,
      redirectUri: window.location.origin + '/index.html',
    },
    cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
  };

  let _msalInstance = null;
  function getMsal() {
    if (!_msalInstance && typeof msal !== 'undefined') {
      _msalInstance = new msal.PublicClientApplication(msalConfig);
    }
    return _msalInstance;
  }

  async function refreshToken() {
    try {
      const instance = getMsal();
      if (!instance) return null;
      const accounts = instance.getAllAccounts();
      if (!accounts.length) return null;
      const res = await instance.acquireTokenSilent({
        scopes:  ['api://' + APP_CLIENT_ID + '/.default'],
        account: accounts[0],
      });
      const newToken = res.idToken || res.accessToken;
      localStorage.setItem('pp_token', newToken);
      try {
        const payload = JSON.parse(atob(newToken.split('.')[1]));
        localStorage.setItem('pp_token_expiry', String(payload.exp * 1000));
      } catch(e) {}
      console.info('[Auth] Token refreshed silently');
      return newToken;
    } catch (e) {
      console.warn('[Auth] Silent refresh failed:', e.message);
      return null;
    }
  }

  async function getValidToken() {
    const token  = localStorage.getItem('pp_token');
    const expiry = parseInt(localStorage.getItem('pp_token_expiry') || '0', 10);

    if (!token) return null;

    // Refresh proactively if expiring within 5 minutes
    if (expiry && Date.now() > expiry - REFRESH_BEFORE_MS) {
      const fresh = await refreshToken();
      return fresh || token;
    }
    return token;
  }

  // Sync version for backwards compatibility — use existing token as-is
  function getTokenSync() {
    return localStorage.getItem('pp_token');
  }

  // authHeaders() — async-safe: awaitable if needed, sync fallback otherwise
  window.authHeaders = function authHeaders(extra) {
    const token = getTokenSync();
    return Object.assign(token ? { Authorization: 'Bearer ' + token } : {}, extra || {});
  };

  // Async version for critical calls
  window.authHeadersAsync = async function authHeadersAsync(extra) {
    const token = await getValidToken();
    return Object.assign(token ? { Authorization: 'Bearer ' + token } : {}, extra || {});
  };

  // Schedule proactive refresh based on stored expiry
  function scheduleTokenRefresh() {
    const expiry = parseInt(localStorage.getItem('pp_token_expiry') || '0', 10);
    if (!expiry) return;
    const msUntilRefresh = expiry - Date.now() - REFRESH_BEFORE_MS;
    if (msUntilRefresh > 0) {
      setTimeout(async function () {
        await refreshToken();
        scheduleTokenRefresh(); // reschedule after refresh
      }, msUntilRefresh);
      console.info('[Auth] Token refresh scheduled in', Math.round(msUntilRefresh / 60000), 'min');
    } else {
      // Already near expiry — refresh now
      refreshToken().then(scheduleTokenRefresh);
    }
  }

  window.scheduleTokenRefresh = scheduleTokenRefresh;

  // Auto-start refresh scheduler on load
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', scheduleTokenRefresh);
  } else {
    scheduleTokenRefresh();
  }
})();
