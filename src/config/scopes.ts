/**
 * Microsoft Graph API scopes configuration for OneDrive/SharePoint/Excel integration
 */

export const GRAPH_SCOPES = {
  // File and drive access
  FILES_READ: "Files.Read",
  FILES_READ_ALL: "Files.Read.All",
  FILES_READWRITE: "Files.ReadWrite",
  FILES_READWRITE_ALL: "Files.ReadWrite.All",

  // SharePoint access
  SITES_READ_ALL: "Sites.Read.All",
  SITES_READWRITE_ALL: "Sites.ReadWrite.All",
  SITES_MANAGE_ALL: "Sites.Manage.All",
  // Required to create new SharePoint sites via `POST /beta/sites`.
  // Not implied by Sites.ReadWrite.All — it's a separate, narrower scope.
  SITES_CREATE_ALL: "Sites.Create.All",

  // User profile
  USER_READ: "User.Read",
  // Required to resolve owner email → user id when creating M365-group-backed
  // team sites (so that `owners@odata.bind` can be populated). Without it,
  // `POST /groups` fails on owner resolution in app-only mode.
  USER_READ_ALL: "User.Read.All",

  // Group management — required for `POST /groups` (creating the M365 group
  // that backs a Teams-connected SharePoint site).
  GROUP_READWRITE_ALL: "Group.ReadWrite.All",

  // Directory access (for business accounts)
  DIRECTORY_READ_ALL: "Directory.Read.All",

  // Application permissions (for unattended scenarios)
  APP_FILES_READ_ALL: "Files.Read.All",
  APP_FILES_READWRITE_ALL: "Files.ReadWrite.All",
  APP_SITES_READ_ALL: "Sites.Read.All",
  APP_SITES_READWRITE_ALL: "Sites.ReadWrite.All",
} as const;

// Scope configurations for different use cases
export const SCOPE_CONFIGURATIONS = {
  // Personal OneDrive access
  PERSONAL: [GRAPH_SCOPES.USER_READ, GRAPH_SCOPES.FILES_READWRITE],

  // Business OneDrive + SharePoint + site/group provisioning.
  // Sites.Create.All, Group.ReadWrite.All and User.Read.All are required by
  // the site-creation tools; they're harmless additions for users who don't
  // use those tools, and including them by default avoids per-tool re-consent.
  BUSINESS: [
    GRAPH_SCOPES.USER_READ,
    GRAPH_SCOPES.USER_READ_ALL,
    GRAPH_SCOPES.FILES_READWRITE_ALL,
    GRAPH_SCOPES.SITES_READWRITE_ALL,
    GRAPH_SCOPES.SITES_CREATE_ALL,
    GRAPH_SCOPES.GROUP_READWRITE_ALL,
  ],

  // Full enterprise access
  ENTERPRISE: [
    GRAPH_SCOPES.USER_READ,
    GRAPH_SCOPES.USER_READ_ALL,
    GRAPH_SCOPES.FILES_READWRITE_ALL,
    GRAPH_SCOPES.SITES_READWRITE_ALL,
    GRAPH_SCOPES.SITES_CREATE_ALL,
    GRAPH_SCOPES.GROUP_READWRITE_ALL,
    GRAPH_SCOPES.DIRECTORY_READ_ALL,
  ],

  // Application-only (service principal). App-only callers use static
  // application permissions granted in Azure (admin consent); the scope list
  // here is informational. The site-creation tools additionally require
  // `Group.Create` (or `Group.ReadWrite.All`) and `User.Read.All` granted as
  // application permissions on the app registration.
  APPLICATION: [
    GRAPH_SCOPES.APP_FILES_READWRITE_ALL,
    GRAPH_SCOPES.APP_SITES_READWRITE_ALL,
  ],
} as const;

// Default scope for device code flow
export const DEFAULT_SCOPES = SCOPE_CONFIGURATIONS.BUSINESS;

export type ScopeConfiguration = keyof typeof SCOPE_CONFIGURATIONS;
