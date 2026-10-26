/**
 * MSAL Authentication module — Portal (public client, browser-only)
 * Copyright © 2026 Sebastian Flæng Markdanner — MIT License
 *
 * Wraps MSAL's PublicClientApplication with:
 *   - Redirect-flow handling on page load
 *   - Silent token acquisition with interactive popup fallback
 *   - ARM token acquisition for Azure Resource PIM
 *   - Several signed-in accounts per tab (one active), with an optional
 *     deploy-time tenant allow-list and chained sign-out of every account
 */

/* global msal, msalConfig, GRAPH_SCOPES, ARM_SCOPE, accountUtils */

const msalInstance = new msal.PublicClientApplication(window.msalConfig);

/** Resolves when MSAL v5 initialization completes. Awaited once in initAuth(). */
const _msalReady = msalInstance.initialize();

/** Currently active account */
let _account = null;

/** sessionStorage key that persists the active account entry (user + tenant) across F5 refreshes */
const ACTIVE_ACCOUNT_KEY = 'pim-portal-active-account';

/** sessionStorage key holding the homeAccountIds still to be signed out by signOutAll() */
const SIGNOUT_QUEUE_KEY = 'pim-portal-signout-queue';

/** Tenant allow-list injected at deploy time; empty allows any tenant. */
const _allowedTenants = accountUtils.parseAllowedTenants(window.PORTAL_ALLOWED_TENANTS);

/** Tenant of an account that was rejected by the allow-list during initAuth(). */
let _blockedTenantId = null;

/** Claims string to include in every token acquisition while an auth-context activation is in progress. */
let _activeAuthContextClaims = null;

function _tenantOf(account) {
  return account?.tenantId || account?.idTokenClaims?.tid || '';
}

/**
 * Must be called once on app startup.
 * Handles the redirect callback if present and activates an allowed account:
 * the one just signed in, else the one stored for this tab, else the first.
 * Accounts from tenants outside the allow-list are never activated; see
 * getBlockedTenantId().
 * @returns {Promise<msal.AccountInfo|null>}
 */
async function initAuth() {
  try {
    await _msalReady;
    const response = await msalInstance.handleRedirectPromise();
    if (response && response.account) {
      if (isTenantAllowed(_tenantOf(response.account))) {
        sessionStorage.setItem(ACTIVE_ACCOUNT_KEY, accountUtils.entryKey(response.account));
      } else {
        _blockedTenantId = _tenantOf(response.account);
      }
    }
  } catch (err) {
    console.error('[Auth] handleRedirectPromise error:', err);
  }

  const accounts = msalInstance.getAllAccounts();
  const chosen   = accountUtils.pickActiveAccount(accounts, sessionStorage.getItem(ACTIVE_ACCOUNT_KEY), _allowedTenants);
  if (!chosen) {
    if (!_blockedTenantId) {
      const disallowed = accounts.find(a => !isTenantAllowed(_tenantOf(a)));
      if (disallowed) _blockedTenantId = _tenantOf(disallowed);
    }
    return null;
  }

  msalInstance.setActiveAccount(chosen);
  _account = chosen;
  sessionStorage.setItem(ACTIVE_ACCOUNT_KEY, accountUtils.entryKey(chosen));
  return _account;
}

/**
 * Tenant ID of an account the allow-list rejected during initAuth(), or null.
 * When initAuth() returned an account, the rejected one was a newly added
 * account and the previous account stayed active.
 * @returns {string|null}
 */
function getBlockedTenantId() {
  return _blockedTenantId;
}

/**
 * @param {string} tenantId
 * @returns {boolean} true when the deploy-time allow-list permits this tenant
 */
function isTenantAllowed(tenantId) {
  return accountUtils.isTenantAllowed(tenantId, _allowedTenants);
}

/**
 * @returns {string} entry key (homeAccountId|tenantId) of the active account, or ''
 */
function getActiveAccountKey() {
  return accountUtils.entryKey(_account);
}

/**
 * Signed-in account entries the portal may use, active entry first, then by name.
 * @returns {msal.AccountInfo[]}
 */
function listAccounts() {
  const activeKey = getActiveAccountKey();
  return msalInstance.getAllAccounts()
    .filter(a => isTenantAllowed(_tenantOf(a)))
    .sort((a, b) => {
      if (accountUtils.entryKey(a) === activeKey) return -1;
      if (accountUtils.entryKey(b) === activeKey) return 1;
      return String(a.name || a.username).localeCompare(String(b.name || b.username));
    });
}

/**
 * Make another signed-in account entry active. Reloads the page so no
 * in-memory state (roles, policies, selections, timers) carries over;
 * tokens come from the MSAL cache without a Microsoft round trip.
 * @param {string} key  entry key from accountUtils.entryKey()
 */
function switchAccount(key) {
  const target = listAccounts().find(a => accountUtils.entryKey(a) === key);
  if (!target) throw new Error('Account not available');
  msalInstance.setActiveAccount(target);
  sessionStorage.setItem(ACTIVE_ACCOUNT_KEY, key);
  window.location.reload();
}

/**
 * Sign one user out of the portal (all of their tenant entries — MSAL removes
 * accounts by homeAccountId). The user's Microsoft browser session is left
 * alone; signOutAll() ends sessions too. Signing out the last allowed user
 * falls through to signOutAll().
 * @param {string} homeAccountId
 */
async function signOutUser(homeAccountId) {
  const accounts = msalInstance.getAllAccounts();
  const target   = accounts.find(a => a.homeAccountId === homeAccountId);
  if (!target) return;

  const remaining = accounts.filter(a => a.homeAccountId !== homeAccountId && isTenantAllowed(_tenantOf(a)));
  if (!remaining.length) return signOutAll();

  await msalInstance.clearCache({ account: target });
  if (_account?.homeAccountId === homeAccountId) {
    sessionStorage.setItem(ACTIVE_ACCOUNT_KEY, accountUtils.entryKey(remaining[0]));
  }
  window.location.reload();
}

/**
 * Sign every account out of the portal and end each one's Microsoft session.
 * Entra signs out one account per logout redirect, so the remaining users are
 * queued in sessionStorage and bootstrap resumes via continueSignOutAll().
 */
async function signOutAll() {
  const targets = accountUtils.logoutTargets(msalInstance.getAllAccounts());
  sessionStorage.setItem(SIGNOUT_QUEUE_KEY, JSON.stringify(targets));
  sessionStorage.removeItem(ACTIVE_ACCOUNT_KEY);
  if (!(await continueSignOutAll())) window.location.reload();
}

/**
 * @returns {boolean} true while a signOutAll() chain is in progress
 */
function isSigningOutAll() {
  return sessionStorage.getItem(SIGNOUT_QUEUE_KEY) !== null;
}

/**
 * Sign out the next queued user. Each entry is removed from the queue before
 * its redirect, so a failed or abandoned logout can never loop.
 * @returns {Promise<boolean>} true when the page is navigating to Microsoft;
 *   false when the queue is finished (and cleared)
 */
async function continueSignOutAll() {
  await _msalReady;
  try { await msalInstance.handleRedirectPromise(); } catch { /* returning from logout */ }

  let queue;
  try { queue = JSON.parse(sessionStorage.getItem(SIGNOUT_QUEUE_KEY) || '[]'); } catch { queue = []; }
  if (!Array.isArray(queue)) queue = [];

  while (queue.length) {
    const homeAccountId = queue.shift();
    sessionStorage.setItem(SIGNOUT_QUEUE_KEY, JSON.stringify(queue));
    const account = msalInstance.getAllAccounts().find(a => a.homeAccountId === homeAccountId);
    if (!account) continue;
    try {
      await msalInstance.logoutRedirect({ account, postLogoutRedirectUri: window.location.origin });
      return true;
      // Page navigates away — execution does not continue here.
    } catch (err) {
      // A failed logoutRedirect leaves MSAL's interaction flag set for this page,
      // which would make every further logout throw. Reload: the next load's
      // handleRedirectPromise clears the flag and continues with the next entry.
      console.warn('[Auth] logoutRedirect failed, continuing after reload:', err);
      await msalInstance.clearCache({ account }).catch(() => {});
      window.location.reload();
      return true;
    }
  }

  sessionStorage.removeItem(SIGNOUT_QUEUE_KEY);
  return false;
}

/**
 * Trigger interactive sign-in (redirect).
 * The page navigates away to Microsoft and returns to /auth-callback,
 * which is rewritten to index.html. handleRedirectPromise() in initAuth()
 * picks up the response on the way back.
 */
async function signIn() {
  // Include ARM scope alongside Graph scopes so the admin/user only sees
  // a single consent prompt covering all required permissions.
  await msalInstance.loginRedirect({
    scopes: [...window.GRAPH_SCOPES, window.ARM_SCOPE],
    prompt: 'select_account'
  });
  // Page navigates away — execution does not continue here.
}

/**
 * Acquire a Graph access token (silent first, popup fallback).
 * @returns {Promise<string>} access token
 */
async function getGraphToken() {
  return _acquireToken(window.GRAPH_SCOPES, _activeAuthContextClaims || undefined);
}

/**
 * Acquire an ARM access token (silent first, popup fallback).
 * @returns {Promise<string>} access token
 */
async function getArmToken() {
  return _acquireToken([window.ARM_SCOPE], _activeAuthContextClaims || undefined);
}

/**
 * Set an auth context ID whose claims should be included in every token
 * acquisition until cleared. Call before bulkActivate, clear in a finally block.
 * @param {string|null} authContextId
 */
function setAuthContextClaims(authContextId) {
  _activeAuthContextClaims = authContextId
    ? JSON.stringify({ access_token: { acrs: { essential: true, value: authContextId } } })
    : null;
}

/**
 * Set a raw claims JSON string (as MSAL expects) to include in every subsequent
 * token acquisition. Used to resume after a reactive claims-challenge step-up
 * where the server-provided claims blob is replayed verbatim.
 * @param {string|null} claims
 */
function setRawClaims(claims) {
  _activeAuthContextClaims = claims || null;
}

/**
 * Error thrown by API clients when the server returns a 401 with an
 * "insufficient_claims" challenge. Carries the raw MSAL-ready claims blob and
 * the scopes the next token acquisition should request, so the caller can
 * trigger acquireTokenRedirect({ claims, scopes }).
 */
class ClaimsChallengeError extends Error {
  constructor({ claims, scopes, status, message }) {
    super(message || 'Claims challenge required');
    this.name    = 'ClaimsChallengeError';
    this.claims  = claims;
    this.scopes  = scopes;
    this.status  = status;
  }
}

/**
 * Parse an "insufficient_claims" challenge from a 401 response.
 * Looks for a `claims="<base64url>"` parameter in the WWW-Authenticate header
 * (the standard CAE / Conditional Access claims-challenge format), falling
 * back to scanning the response body for a `claims` field. Returns the decoded
 * JSON string MSAL expects, or null if no challenge is present.
 * @param {string|null} wwwAuthenticate
 * @param {string|null} bodyText
 * @returns {string|null}
 */
function parseClaimsChallenge(wwwAuthenticate, bodyText) {
  // 1) WWW-Authenticate: Bearer realm="...", error="insufficient_claims", claims="<base64url>"
  if (wwwAuthenticate) {
    const match = wwwAuthenticate.match(/claims\s*=\s*"([^"]+)"/i);
    if (match) {
      const decoded = _decodeBase64UrlJson(match[1]);
      if (decoded) return decoded;
    }
  }
  // 2) JSON body fallback: { "error": "insufficient_claims", "claims": "<base64url|json>" }
  if (bodyText) {
    try {
      const parsed = JSON.parse(bodyText);
      // ARM/Graph sometimes nest under .error
      const claimsField =
        parsed?.claims ??
        parsed?.error?.claims ??
        parsed?.error?.additionalInfo?.find(i => i?.type === 'claims')?.info?.claims;
      if (typeof claimsField === 'string' && claimsField.length > 0) {
        const decoded = _decodeBase64UrlJson(claimsField);
        if (decoded) return decoded;
        // Some servers already return a JSON string
        try { JSON.parse(claimsField); return claimsField; } catch { /* not JSON */ }
      }
    } catch { /* not JSON */ }
  }
  return null;
}

function _decodeBase64UrlJson(value) {
  try {
    let s = String(value).replace(/-/g, '+').replace(/_/g, '/');
    while (s.length % 4) s += '=';
    const json = atob(s);
    // Validate it parses as JSON before returning
    JSON.parse(json);
    return json;
  } catch {
    return null;
  }
}

/**
 * Trigger an interactive redirect to acquire a new token with the supplied
 * claims challenge. The page navigates away — the caller must persist any
 * state they want to resume in sessionStorage *before* awaiting this.
 * @param {{ scopes: string[], claims: string }} opts
 */
async function stepUpWithClaims({ scopes, claims }) {
  if (!_account) throw new Error('Not signed in');
  await msalInstance.acquireTokenRedirect({
    scopes,
    claims,
    account:   _account,
    authority: _authority()
  });
  // Page navigates away — execution does not continue here.
}

/**
 * @returns {msal.AccountInfo|null}
 */
function getAccount() {
  return _account;
}

/**
 * Returns the signed-in user's object ID (oid), required for Graph/ARM
 * $filter expressions. localAccountId is the oid claim from the id token.
 * @returns {string|null}
 */
function getUserId() {
  return _account?.localAccountId || null;
}

// ── Private helpers ──────────────────────────────────────────────────────────

/**
 * Returns the authority URL for the currently active account's tenant.
 * Always specifying this explicitly prevents MSAL from falling back to the
 * msalConfig default authority (the app's home tenant) when acquiring tokens
 * for a guest/switched tenant — which would cause ARM to return home-tenant
 * resources instead of the target tenant's resources.
 */
function _authority() {
  const tid = _account?.tenantId || _account?.idTokenClaims?.tid;
  return tid
    ? `https://login.microsoftonline.com/${tid}`
    : window.msalConfig.auth.authority;
}

async function _acquireToken(scopes, claims) {
  if (!_account) throw new Error('Not signed in');

  const authority = _authority();
  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes,
      account:   _account,
      authority,
      ...(claims ? { claims } : {})
    });
    return result.accessToken;
  } catch (err) {
    if (err instanceof msal.InteractionRequiredAuthError) {
      // Silent acquisition failed — redirect to Microsoft for fresh tokens.
      // The page will reload on return and initAuth() will process the response.
      await msalInstance.acquireTokenRedirect({ scopes, account: _account, authority, ...(claims ? { claims } : {}) });
      // Page navigates away — execution does not continue here.
    }
    throw err;
  }
}

/**
 * Acquire a Graph token with a specific authentication context (auth context ID).
 * Used when a PIM role requires Conditional Access step-up.
 * @param {string} authContextId
 * @returns {Promise<string>}
 */
async function getGraphTokenWithAuthContext(authContextId) {
  const claims = JSON.stringify({
    access_token: { acrs: { essential: true, value: authContextId } }
  });
  return _acquireToken(window.GRAPH_SCOPES, claims);
}

/**
 * Proactively step up authentication for Conditional Access auth context IDs.
 * Uses acquireTokenRedirect (same flow as sign-in and MFA challenges) so the
 * user never sees a popup window. The page navigates away to Microsoft and
 * returns; bootstrap detects the saved PENDING_ACTIVATION_KEY and resumes.
 *
 * @param {string[]} authContextIds  - unique auth context IDs (e.g. 'c2', 'c3')
 * @param {object[]} roles           - the roles being activated
 */
async function stepUpForAuthContexts(authContextIds, roles) {
  const hasEntraGroup = roles.some(r => r.type === 'User' || r.type === 'Group');

  // Prefer Graph scopes; ARM follows automatically because CA is session-level.
  const scopes = hasEntraGroup ? window.GRAPH_SCOPES : [window.ARM_SCOPE];

  // Use the first auth context ID — session-level satisfaction covers all.
  const claims = JSON.stringify({
    access_token: { acrs: { essential: true, value: authContextIds[0] } }
  });

  await msalInstance.acquireTokenRedirect({ scopes, account: _account, authority: _authority(), claims });
  // Page navigates away — execution does not continue here.
}

/**
 * Request ARM consent interactively via a popup (used by the consent banner).
 * Unlike getArmToken(), this always uses a popup so the user doesn't navigate away.
 * @returns {Promise<void>}
 */
async function grantArmConsent() {
  if (!_account) throw new Error('Not signed in');
  await msalInstance.acquireTokenPopup({ scopes: [window.ARM_SCOPE], account: _account, authority: _authority() });
}

/**
 * Switch to a different Entra tenant using the currently signed-in account.
 * Uses loginRedirect with a tenant-specific authority and the current account's
 * UPN as loginHint so Microsoft can SSO the user silently when they are already
 * a guest/member of that tenant. The page navigates away; on return, initAuth()
 * activates and stores the new tenant's account entry via handleRedirectPromise().
 * If the switch is cancelled, the previous entry stays active.
 * @param {string} tenantId  Target tenant GUID
 */
async function switchTenant(tenantId) {
  // Include ARM scope so both Graph and ARM are consented in one single prompt.
  await msalInstance.loginRedirect({
    authority: `https://login.microsoftonline.com/${tenantId}`,
    scopes:     [...window.GRAPH_SCOPES, window.ARM_SCOPE],
    loginHint:  _account?.username,
  });
  // Page navigates away — execution does not continue here.
}

// Expose globally for other modules
window.portalAuth = { initAuth, signIn, signOutUser, signOutAll, isSigningOutAll, continueSignOutAll, getBlockedTenantId, isTenantAllowed, getActiveAccountKey, listAccounts, switchAccount, getGraphToken, getArmToken, getGraphTokenWithAuthContext, stepUpForAuthContexts, stepUpWithClaims, parseClaimsChallenge, ClaimsChallengeError, grantArmConsent, setAuthContextClaims, setRawClaims, getAccount, getUserId, switchTenant };
