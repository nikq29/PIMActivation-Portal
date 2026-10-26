/**
 * Account helpers — Portal (pure functions, no MSAL or DOM access)
 * Copyright © 2026 Sebastian Flæng Markdanner — MIT License
 *
 * Shared by auth.js (browser, via window.accountUtils) and the unit tests
 * (Node, via module.exports). An "account entry" is a user in a tenant:
 * MSAL returns one AccountInfo per tenant profile, so the same user can
 * appear once for their home tenant and once per guest tenant.
 */

(function () {
  const PLACEHOLDER = '__PORTAL_ALLOWED_TENANTS__';

  function _tenantOf(account) {
    return account?.tenantId || account?.idTokenClaims?.tid || '';
  }

  /**
   * Stable identifier for an account entry (user + tenant).
   * @param {object|null} account  MSAL AccountInfo
   * @returns {string}
   */
  function entryKey(account) {
    if (!account) return '';
    return account.homeAccountId + '|' + _tenantOf(account);
  }

  /**
   * Parse the deploy-time allow-list. An empty value or the unreplaced
   * placeholder means "allow any tenant". Unrecognised tokens are kept so a
   * malformed list fails closed (they simply never match).
   * @param {string|undefined} raw  comma/semicolon/whitespace separated tenant IDs
   * @returns {string[]} lowercased, de-duplicated tenant IDs
   */
  function parseAllowedTenants(raw) {
    if (typeof raw !== 'string' || raw.includes(PLACEHOLDER)) return [];
    const tokens = raw.split(/[\s,;]+/).map(t => t.trim().toLowerCase()).filter(Boolean);
    return [...new Set(tokens)];
  }

  /**
   * @param {string} tenantId
   * @param {string[]} allowed  output of parseAllowedTenants
   * @returns {boolean}
   */
  function isTenantAllowed(tenantId, allowed) {
    if (!allowed.length) return true;
    return allowed.includes(String(tenantId || '').toLowerCase());
  }

  /**
   * Choose which cached account entry to activate.
   * @param {object[]} accounts      MSAL AccountInfo list
   * @param {string|null} preferredKey  entryKey stored for this tab
   * @param {string[]} allowed
   * @returns {object|null}
   */
  function pickActiveAccount(accounts, preferredKey, allowed) {
    const candidates = accounts.filter(a => isTenantAllowed(_tenantOf(a), allowed));
    if (!candidates.length) return null;
    return candidates.find(a => entryKey(a) === preferredKey) ?? candidates[0];
  }

  /**
   * All tenant entries belonging to one user.
   * @param {object[]} accounts
   * @param {string} homeAccountId
   * @returns {object[]}
   */
  function entriesForUser(accounts, homeAccountId) {
    return accounts.filter(a => a.homeAccountId === homeAccountId);
  }

  /**
   * One sign-out target per user, in first-seen order.
   * @param {object[]} accounts
   * @returns {string[]} homeAccountIds
   */
  function logoutTargets(accounts) {
    return [...new Set(accounts.map(a => a.homeAccountId))];
  }

  const accountUtils = { entryKey, parseAllowedTenants, isTenantAllowed, pickActiveAccount, entriesForUser, logoutTargets };

  if (typeof window !== 'undefined') window.accountUtils = accountUtils;
  if (typeof module !== 'undefined' && module.exports) module.exports = accountUtils;
})();
