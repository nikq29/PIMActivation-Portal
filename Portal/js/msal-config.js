/**
 * MSAL Public Client Configuration
 * Copyright © 2026 Sebastian Flæng Markdanner — MIT License
 *
 * CLIENT_ID and TENANT_ID are injected at deploy time via sed replacement.
 * __PORTAL_CLIENT_ID__  → GitHub secret PORTAL_CLIENT_ID  (hosted CI)
 *                        → Bicep-created Entra app ID        (self-hosted ARM)
 * __PORTAL_TENANT_ID__  → 'organizations'                   (hosted CI, multi-tenant)
 *                        → Bicep parameter tenantId          (self-hosted ARM, single-tenant)
 *                        → 'organizations'                   (self-hosted ARM, multiTenant = true)
 * __PORTAL_ALLOWED_TENANTS__ → ''                            (hosted CI, any tenant)
 *                        → Bicep parameter allowedTenantIds  (self-hosted ARM, comma-separated)
 *
 * Scopes requested on first sign-in (delegated, user-consented).
 * We request the minimal set; additional scopes are acquired silently on demand.
 */

/* global msal */

window.msalConfig = {
  auth: {
    clientId:    '__PORTAL_CLIENT_ID__',
    authority:   'https://login.microsoftonline.com/__PORTAL_TENANT_ID__',
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin
  },
  cache: {
    cacheLocation: 'sessionStorage',
  },
  system: {
    loggerOptions: {
      loggerCallback(level, message) {
        if (level === msal.LogLevel.Error || level === msal.LogLevel.Warning) {
          console.warn('[MSAL]', message);
        }
      },
      piiLoggingEnabled: false,
      logLevel: msal.LogLevel.Warning
    }
  }
};

/**
 * Comma-separated tenant IDs allowed to sign in. Empty (or the unreplaced
 * placeholder during local development) allows any tenant. This only shapes
 * the UI — access control is per-tenant admin consent / assignment.
 */
window.PORTAL_ALLOWED_TENANTS = '__PORTAL_ALLOWED_TENANTS__';

/** Delegated scopes needed for Entra PIM and user profile */
window.GRAPH_SCOPES = [
  'User.Read',
  'RoleManagement.ReadWrite.Directory',
  'PrivilegedAccess.ReadWrite.AzureADGroup',
  'RoleManagementPolicy.Read.AzureADGroup',
  'PrivilegedAccess-CustomExt.Read.All',
  'Policy.Read.All',
  'AdministrativeUnit.Read.All',
  'AuditLog.Read.All',
  'offline_access',
  'openid',
  'profile',
  'email'
];

/** Delegated scope for Azure ARM (Resource PIM) */
window.ARM_SCOPE = 'https://management.azure.com/user_impersonation';
