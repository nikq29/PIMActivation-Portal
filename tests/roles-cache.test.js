/**
 * Unit tests for the RoleManager session cache in Portal/js/roles.js:
 * cached roles belong to one account entry (user + tenant), never a tenant.
 * Run with: node --test 'tests/*.test.js'
 */

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const path   = require('node:path');
const vm     = require('node:vm');

const SRC = fs.readFileSync(path.join(__dirname, '..', 'Portal', 'js', 'roles.js'), 'utf8');

function loadRoleManager(activeKey) {
  const data = {};
  const window = vm.createContext({
    sessionStorage: {
      getItem: k => (k in data ? data[k] : null),
      setItem: (k, v) => { data[k] = String(v); },
      removeItem: k => { delete data[k]; }
    },
    portalAuth: {
      getActiveAccountKey: () => activeKey.current,
      getAccount: () => ({ tenantId: 'shared-tenant' })
    },
    console
  });
  window.window = window;
  vm.runInContext(SRC, window);
  return { rm: window.roleManager, data };
}

test('role cache is stored under the active account entry', () => {
  const active = { current: 'alice|t1' };
  const { rm, data } = loadRoleManager(active);
  rm._saveRoleCache();
  const stored = JSON.parse(data['pim-portal-roles-cache']);
  assert.equal(stored.v, 2);
  assert.equal(stored.accountKey, 'alice|t1');
});

test('role cache is reused for the same account entry', () => {
  const active = { current: 'alice|t1' };
  const { rm } = loadRoleManager(active);
  rm.eligibleRoles = [{ id: 'r1' }];
  rm._saveRoleCache();
  // JSON round-trip: the cached array comes from the vm realm
  assert.deepEqual(JSON.parse(JSON.stringify(rm._loadRoleCache().eligibleRoles)), [{ id: 'r1' }]);
});

test('role cache is not shared between two users in the same tenant', () => {
  const active = { current: 'alice|t1' };
  const { rm } = loadRoleManager(active);
  rm.eligibleRoles = [{ id: 'r1' }];
  rm._saveRoleCache();
  active.current = 'bob|t1';
  assert.equal(rm._loadRoleCache(), null);
});

test('tenant-keyed v1 cache entries are ignored', () => {
  const active = { current: 'alice|t1' };
  const { rm, data } = loadRoleManager(active);
  data['pim-portal-roles-cache'] = JSON.stringify({ v: 1, tenantId: 'shared-tenant', eligibleRoles: [], activeRoles: [] });
  assert.equal(rm._loadRoleCache(), null);
});
