/**
 * Unit tests for Portal/js/accounts.js
 * Run with: node --test 'tests/*.test.js'
 */

const test   = require('node:test');
const assert = require('node:assert/strict');

const {
  entryKey,
  parseAllowedTenants,
  isTenantAllowed,
  pickActiveAccount,
  entriesForUser,
  logoutTargets,
  pendingBelongsToAccount,
  mergeTenantCache
} = require('../Portal/js/accounts.js');

const A      = { homeAccountId: 'uidA.tidA', tenantId: 'aaaaaaaa-0000-0000-0000-000000000001', name: 'Alice' };
const B      = { homeAccountId: 'uidB.tidB', tenantId: 'bbbbbbbb-0000-0000-0000-000000000002', name: 'Bob' };
const Aguest = { homeAccountId: 'uidA.tidA', tenantId: 'cccccccc-0000-0000-0000-000000000003', name: 'Alice' };

test('entryKey combines homeAccountId and tenant', () => {
  assert.equal(entryKey(A), 'uidA.tidA|aaaaaaaa-0000-0000-0000-000000000001');
  assert.equal(entryKey({ homeAccountId: 'x', idTokenClaims: { tid: 't1' } }), 'x|t1');
  assert.equal(entryKey(null), '');
});

test('parseAllowedTenants: empty and placeholder allow any', () => {
  for (const raw of [undefined, '', '__PORTAL_ALLOWED_TENANTS__']) {
    assert.deepEqual(parseAllowedTenants(raw), []);
  }
});

test('parseAllowedTenants normalises case, separators, duplicates', () => {
  assert.deepEqual(
    parseAllowedTenants(' AAAAAAAA-0000-0000-0000-000000000001 ; aaaaaaaa-0000-0000-0000-000000000001,,\nbbbbbbbb-0000-0000-0000-000000000002, '),
    ['aaaaaaaa-0000-0000-0000-000000000001', 'bbbbbbbb-0000-0000-0000-000000000002']
  );
});

test('parseAllowedTenants fails closed on malformed entries', () => {
  assert.deepEqual(parseAllowedTenants('not-a-guid'), ['not-a-guid']);
  assert.equal(isTenantAllowed('aaaaaaaa-0000-0000-0000-000000000001', parseAllowedTenants('not-a-guid')), false);
});

test('isTenantAllowed: empty list allows all, match is case-insensitive', () => {
  assert.equal(isTenantAllowed('anything', []), true);
  assert.equal(isTenantAllowed('AAAAAAAA-0000-0000-0000-000000000001', ['aaaaaaaa-0000-0000-0000-000000000001']), true);
  assert.equal(isTenantAllowed('', ['aaaaaaaa-0000-0000-0000-000000000001']), false);
});

test('pickActiveAccount prefers stored key, skips disallowed, falls back', () => {
  const allowB = [B.tenantId];
  assert.equal(pickActiveAccount([A, B], entryKey(B), []), B);
  assert.equal(pickActiveAccount([A, B], entryKey(A), allowB), B);   // stored key disallowed
  assert.equal(pickActiveAccount([A, B], null, []), A);              // legacy session, no key
  assert.equal(pickActiveAccount([A], null, allowB), null);          // nothing allowed
  assert.equal(pickActiveAccount([], null, []), null);
});

test('entriesForUser and logoutTargets group by homeAccountId', () => {
  assert.deepEqual(entriesForUser([A, B, Aguest], A.homeAccountId), [A, Aguest]);
  assert.deepEqual(logoutTargets([A, B, Aguest]), [A.homeAccountId, B.homeAccountId]);
});

test('pendingBelongsToAccount only resumes for the account that saved it', () => {
  assert.equal(pendingBelongsToAccount({ accountKey: entryKey(A) }, entryKey(A)), true);
  assert.equal(pendingBelongsToAccount({ accountKey: entryKey(A) }, entryKey(B)), false);
  assert.equal(pendingBelongsToAccount({ accountKey: entryKey(A) }, ''), false);
  assert.equal(pendingBelongsToAccount({}, entryKey(A)), true);   // saved by a pre-upgrade tab
});

test('mergeTenantCache keeps tenants reported by other accounts', () => {
  const cache = { [A.tenantId]: { tenantId: A.tenantId, displayName: 'Contoso' } };
  const merged = mergeTenantCache(cache, [{ tenantId: B.tenantId, displayName: 'Fabrikam' }]);
  assert.deepEqual(Object.keys(merged).sort(), [A.tenantId, B.tenantId]);
  assert.equal(mergeTenantCache(merged, [{ tenantId: A.tenantId, displayName: 'Contoso Ltd' }])[A.tenantId].displayName, 'Contoso Ltd');
  assert.deepEqual(cache, { [A.tenantId]: { tenantId: A.tenantId, displayName: 'Contoso' } }); // input not mutated
});

test('parseAllowedTenants fails closed when a non-empty list has no usable entries', () => {
  for (const raw of [',', ' , ', '   ', ';']) {
    assert.equal(isTenantAllowed(A.tenantId, parseAllowedTenants(raw)), false, JSON.stringify(raw));
  }
});
