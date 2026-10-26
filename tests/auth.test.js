/**
 * Unit tests for the multi-account parts of Portal/js/auth.js.
 * auth.js is a browser script, so it is loaded into a node:vm context with a
 * minimal fake MSAL PublicClientApplication and sessionStorage.
 * Run with: node --test 'tests/*.test.js'
 */

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const path   = require('node:path');
const vm     = require('node:vm');

const ROOT = path.join(__dirname, '..', 'Portal', 'js');

const T1 = 'aaaaaaaa-0000-0000-0000-000000000001';
const T2 = 'bbbbbbbb-0000-0000-0000-000000000002';
const A  = { homeAccountId: 'uidA.tidA', tenantId: T1, username: 'alice@contoso.com', name: 'Alice' };
const B  = { homeAccountId: 'uidB.tidB', tenantId: T2, username: 'bob@fabrikam.com',  name: 'Bob' };
const key = a => a.homeAccountId + '|' + a.tenantId;

const ACTIVE = 'pim-portal-active-account';
const QUEUE  = 'pim-portal-signout-queue';

function loadAuth({ accounts = [], redirectResponse = null, allowed = '', storage = {}, logoutThrows = false } = {}) {
  let cache = [...accounts];
  const calls = { logoutRedirect: [], clearCache: [], loginRedirect: [], reload: 0 };

  class FakePCA {
    initialize()            { return Promise.resolve(); }
    handleRedirectPromise() { return Promise.resolve(redirectResponse); }
    getAllAccounts()        { return [...cache]; }
    setActiveAccount()      {}
    async clearCache(req) {
      calls.clearCache.push(req?.account?.homeAccountId ?? '*');
      cache = req?.account ? cache.filter(a => a.homeAccountId !== req.account.homeAccountId) : [];
    }
    async logoutRedirect(req) {
      calls.logoutRedirect.push(req.account.homeAccountId);
      if (logoutThrows) throw new Error('logout failed');
      cache = cache.filter(a => a.homeAccountId !== req.account.homeAccountId);
    }
    async loginRedirect(req) { calls.loginRedirect.push(req); }
  }

  const data = { ...storage };
  const sessionStorage = {
    getItem: k => (k in data ? data[k] : null),
    setItem: (k, v) => { data[k] = String(v); },
    removeItem: k => { delete data[k]; }
  };
  // As in a browser, `window` is the global object itself.
  const window = vm.createContext({
    msalConfig: { auth: { authority: 'https://login.microsoftonline.com/organizations' } },
    PORTAL_ALLOWED_TENANTS: allowed,
    GRAPH_SCOPES: ['User.Read'],
    ARM_SCOPE: 'https://management.azure.com/user_impersonation',
    location: { origin: 'https://portal.test', reload() { calls.reload++; } },
    sessionStorage,
    msal: { PublicClientApplication: FakePCA, InteractionRequiredAuthError: class extends Error {} },
    console: { error() {}, warn() {} },
    atob
  });
  window.window = window;
  const ctx = window;

  vm.runInContext(fs.readFileSync(path.join(ROOT, 'accounts.js'), 'utf8'), ctx);
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'auth.js'), 'utf8'), ctx);

  return { auth: window.portalAuth, calls, data, cache: () => cache };
}

test('initAuth restores the stored account entry', async () => {
  const { auth } = loadAuth({ accounts: [A, B], storage: { [ACTIVE]: key(B) } });
  assert.equal(await auth.initAuth(), B);
  assert.equal(auth.getActiveAccountKey(), key(B));
  assert.equal(auth.getBlockedTenantId(), null);
});

test('initAuth stores an allowed account returned by a redirect', async () => {
  const { auth, data } = loadAuth({ accounts: [A, B], redirectResponse: { account: B }, storage: { [ACTIVE]: key(A) } });
  assert.equal(await auth.initAuth(), B);
  assert.equal(data[ACTIVE], key(B));
});

test('initAuth keeps the previous account when a disallowed one is added', async () => {
  const { auth, data } = loadAuth({ accounts: [A, B], allowed: T1, redirectResponse: { account: B }, storage: { [ACTIVE]: key(A) } });
  assert.equal(await auth.initAuth(), A);
  assert.equal(auth.getBlockedTenantId(), T2);
  assert.equal(data[ACTIVE], key(A));
});

test('initAuth returns null and reports the tenant when only disallowed accounts exist', async () => {
  const { auth } = loadAuth({ accounts: [B], allowed: T1 });
  assert.equal(await auth.initAuth(), null);
  assert.equal(auth.getBlockedTenantId(), T2);
});

test('initAuth never falls back to a disallowed account after reload', async () => {
  const { auth } = loadAuth({ accounts: [B, A], allowed: T1, storage: { [ACTIVE]: key(B) } });
  assert.equal(await auth.initAuth(), A);
});

test('listAccounts hides disallowed entries and lists the active entry first', async () => {
  const C = { homeAccountId: 'uidC.tidC', tenantId: T1, username: 'carol@contoso.com', name: 'Carol' };
  const { auth } = loadAuth({ accounts: [A, B, C], allowed: T1, storage: { [ACTIVE]: key(C) } });
  await auth.initAuth();
  assert.deepEqual(auth.listAccounts().map(a => a.name), ['Carol', 'Alice']);
});

test('switchAccount stores the key and reloads; unknown keys throw', async () => {
  const { auth, calls, data } = loadAuth({ accounts: [A, B] });
  await auth.initAuth();
  auth.switchAccount(key(B));
  assert.equal(data[ACTIVE], key(B));
  assert.equal(calls.reload, 1);
  assert.throws(() => auth.switchAccount('nope|nope'), /Account not available/);
});

test('signOutUser removes that user locally and switches to a remaining account', async () => {
  const { auth, calls, data } = loadAuth({ accounts: [A, B], storage: { [ACTIVE]: key(A) } });
  await auth.initAuth();
  await auth.signOutUser(A.homeAccountId);
  assert.deepEqual(calls.clearCache, [A.homeAccountId]);
  assert.deepEqual(calls.logoutRedirect, []);
  assert.equal(data[ACTIVE], key(B));
  assert.equal(calls.reload, 1);
});

test('signOutUser on the last allowed account signs out of everything', async () => {
  const { auth, calls } = loadAuth({ accounts: [A] });
  await auth.initAuth();
  await auth.signOutUser(A.homeAccountId);
  assert.deepEqual(calls.logoutRedirect, [A.homeAccountId]);
});

test('signOutAll queues every user and logs out the first', async () => {
  const { auth, calls, data } = loadAuth({ accounts: [A, B], allowed: T1 });
  await auth.initAuth();
  await auth.signOutAll();
  assert.deepEqual(calls.logoutRedirect, [A.homeAccountId]);
  assert.equal(data[QUEUE], JSON.stringify([B.homeAccountId]));
  assert.equal(data[ACTIVE], undefined);
  assert.equal(auth.isSigningOutAll(), true);
});

test('continueSignOutAll logs out the next user, then finishes', async () => {
  const { auth, calls, data } = loadAuth({ accounts: [B], storage: { [QUEUE]: JSON.stringify([B.homeAccountId]) } });
  assert.equal(await auth.continueSignOutAll(), true);
  assert.deepEqual(calls.logoutRedirect, [B.homeAccountId]);
  assert.equal(await auth.continueSignOutAll(), false);
  assert.equal(data[QUEUE], undefined);
  assert.equal(auth.isSigningOutAll(), false);
});

test('continueSignOutAll skips missing users and never loops', async () => {
  const { auth, calls, data } = loadAuth({ accounts: [], storage: { [QUEUE]: JSON.stringify(['gone']) } });
  assert.equal(await auth.continueSignOutAll(), false);
  assert.deepEqual(calls.logoutRedirect, []);
  assert.equal(data[QUEUE], undefined);
});

test('continueSignOutAll continues past a failed logout', async () => {
  const { auth, calls, data } = loadAuth({ accounts: [A, B], logoutThrows: true, storage: { [QUEUE]: JSON.stringify([A.homeAccountId, B.homeAccountId]) } });
  assert.equal(await auth.continueSignOutAll(), false);
  assert.deepEqual(calls.logoutRedirect, [A.homeAccountId, B.homeAccountId]);
  assert.deepEqual(calls.clearCache, [A.homeAccountId, B.homeAccountId]);
  assert.equal(data[QUEUE], undefined);
});

test('a corrupt sign-out queue is discarded', async () => {
  const { auth, data } = loadAuth({ accounts: [A], storage: { [QUEUE]: '{not json' } });
  assert.equal(await auth.continueSignOutAll(), false);
  assert.equal(data[QUEUE], undefined);
});
