# Usage Guide

This document is the long-form companion to [README.md](../README.md). It walks through everything you need to go from a fresh clone to running tools against a real Microsoft 365 tenant, then documents every one of the 33 MCP tools with example inputs and expected outputs.

If you just want to wire the server into an MCP client and call something, jump to [Quick start](#quick-start). For tool-by-tool reference, jump to [Tool reference](#tool-reference).

---

## Contents

- [How it fits together](#how-it-fits-together)
- [Prerequisites](#prerequisites)
- [Azure app registration](#azure-app-registration)
  - [Device-code (delegated) registration](#device-code-delegated-registration)
  - [Client-credentials (app-only) registration](#client-credentials-app-only-registration)
- [Quick start](#quick-start)
- [Configuration](#configuration)
  - [Environment variables](#environment-variables)
  - [Site registry (`config/sites.local.json`)](#site-registry-configsiteslocaljson)
- [Wiring it into MCP clients](#wiring-it-into-mcp-clients)
  - [Claude Code](#claude-code)
  - [Claude Desktop](#claude-desktop)
  - [Cursor / VS Code MCP-compatible clients](#cursor--vs-code-mcp-compatible-clients)
  - [`mcporter` and other CLI clients](#mcporter-and-other-cli-clients)
- [Using it from the shell (`ods` CLI / `spcall`)](#using-it-from-the-shell-ods-cli--spcall)
- [Drive / site targeting model](#drive--site-targeting-model)
- [Pagination](#pagination)
- [Error envelope](#error-envelope)
- [Tool reference](#tool-reference)
  - [Files](#files-10-tools)
  - [SharePoint](#sharepoint-9-tools)
  - [Utilities](#utilities-5-tools)
  - [Advanced](#advanced-9-tools)
- [Common workflows](#common-workflows)
- [Operational notes](#operational-notes)
- [Troubleshooting](#troubleshooting)

---

## How it fits together

```
┌────────────────────────┐        stdio        ┌──────────────────────────┐
│ MCP client             │ ◀─────────────────▶ │ This server (`mcp-       │
│ (Claude Code / Desktop,│   JSON-RPC frames   │  onedrive-sharepoint`)   │
│  Cursor, mcporter, …)  │                     │                          │
└────────────────────────┘                     │  • 33 tools across       │
                                               │    files / SP / utils /  │
                                               │    advanced              │
                                               │  • Auth (MSAL): device-  │
                                               │    code OR client-       │
                                               │    credentials           │
                                               │  • Graph client (axios   │
                                               │    + retries + 320 KiB-  │
                                               │    aligned chunked       │
                                               │    upload)               │
                                               └────────────┬─────────────┘
                                                            │ HTTPS
                                                            ▼
                                              ┌─────────────────────────┐
                                              │ Microsoft Graph v1.0    │
                                              │ /me, /drives, /sites,   │
                                              │ /sites/{id}/lists,      │
                                              │ /drives/{id}/items, …   │
                                              └─────────────────────────┘
```

All tools speak the same MCP envelope: a JSON-encoded text content block returned by `jsonTextResponse(payload)` on success, or an error envelope produced by `toolErrorResponse(toolName, err)` on failure. See [Error envelope](#error-envelope).

---

## Prerequisites

- **Node.js 20 or later.** Required by `@azure/msal-node` v5; older Node versions are unsupported upstream.
- **Microsoft Entra ID (Azure AD) app registration.** Cookbook for both flows is below.
- **macOS, Linux, or Windows.** On macOS the auth cache prefers the system Keychain; on other platforms (or when Keychain access fails) it falls back to an encrypted-mode file in `$XDG_CACHE_HOME/mcp-onedrive-sharepoint/` (configurable via `MCP_ONEDRIVE_SHAREPOINT_CACHE_DIR`).
- **`mcporter` (optional)** if you want to use the `spcall` wrapper to invoke tools from a shell without binding the server to an MCP client.

---

## Azure app registration

You only need **one** of these two flows. Pick based on whether the server runs interactively (a human can sign in periodically) or unattended (CI, scheduled scripts, multi-user automation).

### Device-code (delegated) registration

Use this when the server acts on behalf of a specific user — files appear as if that user created them, permissions inherit from that user, OneDrive personal drive is accessible.

1. Azure portal → **Entra ID → App registrations → New registration**.
2. Name it (e.g. `mcp-onedrive-sharepoint-local`), choose your supported account type (single tenant is fine; multi-tenant works too).
3. Under **Authentication**:
   - **Allow public client flows: Yes.**
   - No redirect URI required for device-code; leave the platform list empty.
4. Under **API permissions → Add a permission → Microsoft Graph → Delegated permissions**, add:
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
   - `User.Read`
   - `offline_access` (required for silent refresh)
5. **Grant admin consent** for the tenant if your IT policy requires it (recommended).
6. Copy the **Application (client) ID** and **Directory (tenant) ID** from the Overview page.

Then in your `.env`:
```bash
MICROSOFT_GRAPH_CLIENT_ID=<application_client_id>
MICROSOFT_GRAPH_TENANT_ID=<tenant_uuid>   # or "common" for multi-tenant
```
Run `npm run setup-auth` to complete the device-code flow once. The MSAL refresh token is good for 90 days of inactivity.

### Client-credentials (app-only) registration

Use this when the server runs unattended (automation, schedulers, multi-user shared services). The app itself has identity; there is no signed-in user.

1. App registration as above, **single-tenant only** (multi-tenant + client-credentials is rejected by Azure AD).
2. Under **Certificates & secrets → Client secrets → New client secret**, generate a secret and copy the **Value** (you won't see it again).
3. Under **API permissions → Add a permission → Microsoft Graph → Application permissions**, add:
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
4. **Grant admin consent.** Application permissions don't work until consented.

Then in your `.env`:
```bash
MICROSOFT_GRAPH_CLIENT_ID=<application_client_id>
MICROSOFT_GRAPH_TENANT_ID=<tenant_uuid>            # MUST be a UUID, not "common"
MICROSOFT_GRAPH_CLIENT_SECRET=<the_secret_value>   # or use SP_CLIENT_SECRET
```
No `npm run setup-auth` needed in this mode — the server acquires its token on the first tool call. See [`health_check`](#health_check) to verify connectivity.

> **Note on `/me`:** app-only auth has no user identity, so any tool that hits `/me`-style endpoints (`get_user_profile`, the user-info section of `list_drives` without a site, the legacy global search fallback) will not work. Use `siteId` / `driveId` / `site` aliases instead. `health_check` automatically skips `/me` in this mode.

---

## Quick start

```bash
git clone https://github.com/AlexHamn/mcp-onedrive-sharepoint.git
cd mcp-onedrive-sharepoint
npm install
cp .env.example .env
# edit .env with your client id / tenant id (and optionally client secret)

npm run build

# device-code mode only:
npm run setup-auth

# smoke test:
npm run spcall -- health_check
```

If `health_check` returns `"status": "healthy"`, you're done. Wire the server into your MCP client (next section) or call tools from the CLI directly.

---

## Configuration

### Environment variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `MICROSOFT_GRAPH_CLIENT_ID` | yes | — | Azure app registration's Application (client) ID. |
| `MICROSOFT_GRAPH_TENANT_ID` | yes | `common` | Tenant UUID. `common` only works in device-code mode. |
| `MICROSOFT_GRAPH_CLIENT_SECRET` | CC mode only | — | Activates client-credentials. Alias: `SP_CLIENT_SECRET`. |
| `MICROSOFT_GRAPH_SCOPES` | no | `Files.ReadWrite.All,Sites.ReadWrite.All,Directory.Read.All,User.Read,offline_access` | Comma-separated scope list. Applies to device-code only; CC mode requests `https://graph.microsoft.com/.default`. |
| `MICROSOFT_GRAPH_BASE_URL` | no | `https://graph.microsoft.com/v1.0` | Override for sovereign clouds (US Gov, 21Vianet, etc.). |
| `MICROSOFT_GRAPH_TIMEOUT` | no | `30000` | HTTP timeout in ms. |
| `MICROSOFT_GRAPH_MAX_RETRIES` | no | `3` | Retries for transient (429/5xx/network) failures. |
| `MICROSOFT_GRAPH_CACHE_ENABLED` | no | `true` | In-process metadata/search caches. |
| `MICROSOFT_GRAPH_CACHE_TTL` | no | `3600` | Cache TTL in seconds. |
| `MCP_SITES_CONFIG_PATH` | no | `config/sites.local.json` | Path to the site-registry file (see below). |
| `MCP_ONEDRIVE_SHAREPOINT_CACHE_DIR` | no | `$XDG_CACHE_HOME/mcp-onedrive-sharepoint` | File-fallback location when the system keychain is unavailable. |

### Site registry (`config/sites.local.json`)

Tools accept three site-reference styles: raw `siteId`, canonical `siteUrl`, or short `site` alias. Aliases are loaded from a local JSON file so tenant-specific IDs never get committed.

```bash
cp config/sites.example.json config/sites.local.json
# edit with your sites, then they become reachable as site="primary", etc.
```

Schema (an array under `sites`, or the array directly):
```json
{
  "sites": [
    {
      "key": "primary",
      "name": "Primary Workspace",
      "siteId": "yourtenant.sharepoint.com,<site-guid>,<web-guid>",
      "siteUrl": "https://yourtenant.sharepoint.com/sites/Primary",
      "driveId": "b!<long-base64-id>",
      "aliases": ["primary", "main", "/sites/Primary"]
    }
  ]
}
```
- `siteId` is the only required identifier field; everything else is optional but recommended.
- Including `driveId` is what lets the tools skip a `/sites/{id}/drive` lookup on every call.
- Lookups are case-insensitive and ignore protocol/trailing-slash differences.

To find the values, run `discover_sites` or look at the `webUrl` of the site in SharePoint admin. To get a site's `driveId`, run `list_drives --site=<alias>` once and copy the `id` field.

---

## Wiring it into MCP clients

The shipped wrappers (`scripts/run-stdio.sh`, `scripts/spcall.sh`) load `.env` and forward to the built JS, so most clients only need to point at the wrapper.

### Claude Code

Add to `~/.config/claude-code/mcp.json` (or your project's `.claude/mcp.json`):
```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "/absolute/path/to/mcp-onedrive-sharepoint/scripts/run-stdio.sh"
    }
  }
}
```
Then restart Claude Code; the 33 tools appear under the `sharepoint` namespace.

### Claude Desktop

`~/Library/Application Support/Claude/claude_desktop_config.json` on macOS (or the platform equivalent):
```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "/absolute/path/to/mcp-onedrive-sharepoint/scripts/run-stdio.sh"
    }
  }
}
```

### Cursor / VS Code MCP-compatible clients

Cursor accepts the same shape under **Settings → MCP**. If the client wants `command` + `args` split explicitly:
```json
{
  "command": "node",
  "args": ["/absolute/path/to/mcp-onedrive-sharepoint/build/index.js"],
  "env": {
    "MICROSOFT_GRAPH_CLIENT_ID": "…",
    "MICROSOFT_GRAPH_TENANT_ID": "…",
    "MICROSOFT_GRAPH_CLIENT_SECRET": "…"
  }
}
```

### `mcporter` and other CLI clients

```bash
npx -y mcporter call \
  --stdio ./scripts/run-stdio.sh \
  --cwd "$(pwd)" \
  --name sharepoint \
  health_check
```
The `spcall` wrapper does this for you with cleanup:
```bash
./scripts/spcall.sh health_check
./scripts/spcall.sh list_files driveId=b!abc path=/Shared%20Documents
```

---

## Using it from the shell (`ods` CLI / `spcall`)

There are two shell entrypoints, with different lifetimes:

| Entry point | Spawns a long-running process? | When to use |
|---|---|---|
| `ods <tool> ...` | No — direct one-shot call to the handler, in-process. | Development, scripting, anywhere you don't need the MCP wire protocol. |
| `./scripts/spcall.sh <tool> ...` | Yes briefly — spawns the MCP stdio server, calls the tool via `mcporter`, then kills the child. | Exercising the *exact* MCP wire path the LLM client will use. |

`ods` after `npm run build`:
```bash
ods list                                         # list every tool + description
ods schema list_files                            # print JSON schema for a tool
ods auth                                         # interactive device-code setup
ods list_files --site=primary --path=/           # invoke with flags
ods upload_file --json '{"driveId":"b!abc","remotePath":"/x.txt","localPath":"./x.txt"}'
```
Rules for `ods` flags:
- `--key=value` and `--key value` both work.
- `true` / `false` / `null` and numeric strings are coerced; everything else stays a string.
- Bare flags become `true`.
- `--json '<payload>'` is the escape hatch for nested objects/arrays. Layered `--key=value` overrides win.
- Exit code is `0` on success, `2` if the tool returned an error envelope.

During development, `npm run cli -- <tool> ...` skips the build step.

---

## Drive / site targeting model

Almost every tool accepts the same set of optional fields to pick a target:

| Field | Meaning | Resolution |
|---|---|---|
| `driveId` | Explicit Graph drive id (`b!…`). | Used verbatim. Highest precedence. |
| `siteId` | Graph site id (`host.sharepoint.com,<guid>,<guid>`). | Used verbatim; the site's default drive is targeted unless a `driveId` is also supplied. |
| `siteUrl` | Canonical site URL. | If it matches a `sites.local.json` entry, pulls in that entry's `siteId`/`driveId`. Otherwise resolved via Graph (`/sites/{host}:{path}`). |
| `site` | Alias or fuzzy reference. | Matched against `sites.local.json` aliases/names/keys; falls back to URL resolution if it looks URL-shaped. |
| _(none)_ | Default. | Personal OneDrive (`/me/drive`). In CC mode there is no `/me`, so this defaults to nothing — you **must** supply a target. |

If you pass a SharePoint reference that can't be resolved, the tools refuse to silently fall back to personal OneDrive (this was a deliberate fix — see the [resolver test](../src/tests/sharepoint-site-resolver.test.ts#L122) for the "Refuses to fall back" guarantee).

For `copy_item` specifically there's a parallel destination-side set: `destinationDriveId`, `destinationSite`, `destinationSiteId`, `destinationSiteUrl`. Cross-drive copies must use one of these.

---

## Pagination

Five listing tools return a `pagination` object alongside their items:

- `list_files`
- `search_files`
- `discover_sites`
- `list_site_lists`
- `list_items`

When Graph returns `@odata.nextLink`, the payload looks like:
```json
{
  "items": [ ... ],
  "pagination": {
    "returned": 50,
    "limit": 50,
    "totalCount": 1342,
    "nextPageToken": "https://graph.microsoft.com/v1.0/drives/b!.../root/children?$skiptoken=...",
    "hasMore": true
  }
}
```
Pass the `nextPageToken` back as the tool's `pageToken` argument to fetch the next page. The token is opaque — don't try to parse it.

The underlying `getAllPages` helper caps unbounded crawls at 10 000 items / 1 000 pages by default; if a crawl hits a cap, the response includes `truncated: true` and a `truncationReason` so the caller can resume.

---

## Error envelope

Every tool routes failures through `toolErrorResponse(name, err)`. The response is **two** content blocks:

1. A short human-readable summary line (`Error in <tool>: <message>`).
2. A JSON payload with full structure:
```json
{
  "success": false,
  "tool": "list_files",
  "error": "Access denied",
  "category": "Permission",
  "code": "accessDenied",
  "statusCode": 403,
  "isRetryable": false,
  "suggestedAction": "Check that your app has the required permissions and admin consent"
}
```

`category` is one of `Authentication`, `Permission`, `NotFound`, `Throttling`, `Quota`, `Validation`, `Conflict`, `Server`, `Network`, `Unknown`. `isRetryable: true` means the inner retry helper has already exhausted its budget — surface the failure to the user.

When the Graph response body itself contains an `error` payload (some endpoints return `200 OK` with an embedded error), `assertGraphPayloadHasNoError` re-throws it through the same envelope rather than masking it as an empty list.

---

## Tool reference

All input field names are camelCase. Required fields are flagged. Examples assume `site="primary"` is registered in `sites.local.json` with a `driveId`.

### Files (10 tools)

#### `list_files`
List files and folders in a drive.

Inputs:
- `path` *(string, default `""`)* — folder path; empty string means root.
- one of `driveId` | `siteId` | `site` | `siteUrl` *(optional)* — target. Defaults to `/me/drive`.
- `filter` *(string, optional)* — OData `$filter` (e.g. `"file ne null"` to drop folders).
- `orderBy` *(string, default `"name"`)*.
- `limit` *(number, default `100`)*.
- `pageToken` *(string, optional)*.

```json
{ "site": "primary", "path": "/Shared Documents", "limit": 25, "orderBy": "lastModifiedDateTime desc" }
```

#### `download_file`
Download a file by id or path.

Inputs:
- one of `fileId` | `filePath` *(required-ish — at least one)*.
- target fields.
- `outputPath` *(string, optional)* — if set, the file is written here and the response includes its size + path; otherwise the response includes a `<n bytes>` placeholder (raw binary doesn't fit MCP text content well, so use `outputPath` for real downloads).

#### `upload_file`
Upload a local file. Auto-switches to chunked resumable upload above 4 MB.

Inputs:
- `localPath` *(required)*.
- `remotePath` *(required)* — destination including filename.
- target fields.
- `conflictBehavior` *(`fail` | `replace` | `rename`, default `rename`)*.

```json
{ "site": "primary", "localPath": "/Users/me/q4.xlsx", "remotePath": "/Shared Documents/Reports/Q4.xlsx", "conflictBehavior": "replace" }
```

The upload path is sanitized first (see `prepareUploadPath` in [`src/tools/utils/path-helper.ts`](../src/tools/utils/path-helper.ts)). Path changes are reported back in `pathChanges`.

#### `create_folder`
```json
{ "site": "primary", "name": "2026 Q4 Reports", "parentPath": "/Shared Documents" }
```

#### `move_item`
Rename or move an item. Requires at least one of `newName`, `parentFolderId`, `parentFolderPath` — calls with none are rejected (fixed in v1.1).

```json
{ "site": "primary", "itemPath": "/Drafts/foo.xlsx", "parentFolderPath": "/Archive/2026", "newName": "foo-final.xlsx" }
```

If `parentFolderPath` can't be resolved, the PATCH is skipped and the tool returns an error — no more silent no-ops.

#### `delete_item`
```json
{ "site": "primary", "itemPath": "/Trash/old.txt" }
```
`permanent: true` is supported but is honoured by Graph only on certain drive types; verify in your tenant.

#### `search_files`
```json
{ "site": "primary", "query": "quarterly report", "fileTypes": ["xlsx", "pdf"], "limit": 25 }
```

#### `get_file_metadata`
Detail view: hashes, owner, child count (folders), version list (if `includeVersions: true`).

#### `share_item`
Create a sharing link.
```json
{
  "site": "primary",
  "itemPath": "/Shared Documents/Reports/Q4.xlsx",
  "type": "view",
  "scope": "organization",
  "expirationDateTime": "2026-12-31T00:00:00Z"
}
```

#### `copy_item`
Copy a file/folder. Asynchronous on the Graph side — the tool returns once Graph accepts the request.

Same-drive copy:
```json
{ "site": "primary", "itemPath": "/Templates/Report.xlsx", "newName": "Report-2026.xlsx" }
```

Cross-drive / cross-site copy (the v1.1 fix):
```json
{
  "site": "primary",
  "itemId": "01ABC...",
  "destinationSite": "archive",
  "destinationFolderPath": "/Archive/2026",
  "newName": "Q4.xlsx"
}
```
`destinationDriveId`, `destinationSite`, `destinationSiteId`, `destinationSiteUrl` are all accepted; the destination folder path is resolved against the destination drive.

---

### SharePoint (9 tools)

#### `discover_sites`
```json
{ "search": "finance", "limit": 20 }
```
With no `search`, returns `*`-matched sites. `includePersonalSite: true` adds `/sites/root` when accessible.

#### `resolve_site`
Resolve a reference (alias, URL, id) **without** hitting Graph if the registry can satisfy it. Useful for debugging the registry.

```json
{ "site": "primary" }
```
Output includes the full `knownSites` registry snapshot for visibility.

#### `list_site_lists`
```json
{ "site": "primary", "includeHidden": false, "limit": 50 }
```
Hidden / `_`-prefixed system lists are filtered client-side by default.

#### `get_list_schema`
```json
{ "site": "primary", "listId": "9c6b8b70-...", "includeContentTypes": true }
```

#### `list_items`
```json
{
  "site": "primary",
  "listId": "9c6b8b70-...",
  "filter": "fields/Status eq 'Open'",
  "orderBy": "Created desc",
  "select": "Title,Status,AssignedTo",
  "limit": 100
}
```
`select` is automatically rewritten into the correct `$select=id,webUrl,...,fields(<your fields>)` shape.

#### `get_list_item`
```json
{ "site": "primary", "listId": "9c6b8b70-...", "itemId": "42", "expand": "AssignedTo" }
```

#### `create_list_item`
```json
{
  "site": "primary",
  "listId": "9c6b8b70-...",
  "fields": { "Title": "Onboarding", "Status": "Open", "Priority": "High" }
}
```

#### `update_list_item`
```json
{ "site": "primary", "listId": "9c...", "itemId": "42", "fields": { "Status": "Closed" } }
```

#### `delete_list_item`
```json
{ "site": "primary", "listId": "9c...", "itemId": "42" }
```

---

### Utilities (5 tools)

#### `health_check`
Reports server status, auth mode, API connectivity, and (in delegated mode only) user / default-drive info.

```json
{ "includeUserInfo": true, "includeDriveInfo": true }
```

In client-credentials mode the `/me`-based probes are skipped automatically — connectivity is verified by checking that a token can be acquired. The response includes `authentication.authMode` (`"device_code"` or `"client_credentials"`) and `apiConnectivity.probe` (`"me"` vs `"client_credentials_token"`).

#### `get_user_profile`
Returns the authenticated user (device-code mode only — has no meaning in CC mode). Optional `includeManager` and `includePhoto` flags.

#### `list_drives`
List drives the caller can see. With `site=…` it returns only that site's default drive. Without a site reference it joins `/me/drive` with `/me/drives` and dedups.

```json
{ "site": "primary", "includeQuota": true }
```

#### `global_search`
Searches via the Microsoft Search API (`/search/query`) across one or more entity types, with a delegated-mode fallback to `/me/drive/search` when Microsoft Search is unavailable.

```json
{ "query": "Q4 budget", "entityTypes": ["driveItem", "listItem"], "limit": 25, "includeSummary": true }
```

#### `batch_operations`
Forwards a list of up to 20 raw Graph requests to `/$batch` (the Graph-imposed cap). Use this when you need to fan out heterogeneous reads.

```json
{
  "requests": [
    { "id": "1", "method": "GET", "url": "/me/drive/root/children" },
    { "id": "2", "method": "GET", "url": "/sites/contoso.sharepoint.com,abc,def/lists" }
  ],
  "continueOnError": true
}
```

---

### Advanced (9 tools)

#### `advanced_share`
Send permission invitations via email.
```json
{
  "site": "primary",
  "itemPath": "/Shared Documents/Reports/Q4.xlsx",
  "recipients": ["alice@contoso.com", "bob@contoso.com"],
  "permission": "write",
  "requireSignIn": true,
  "sendInvitation": true,
  "message": "Please review by EOW.",
  "expirationDateTime": "2026-12-31T00:00:00Z"
}
```

#### `manage_permissions`
List, update, or revoke an item's permissions.
```json
{ "site": "primary", "itemPath": "/x.xlsx", "action": "list" }
```
For `update`, supply `permissionId` and `newRoles: ["read"]`. For `revoke`, just `permissionId`.

#### `check_user_access`
Inspect what level of access a specific user has on an item — merges direct, shared-link, and inherited roles into a single answer.
```json
{ "site": "primary", "itemPath": "/x.xlsx", "userEmail": "alice@contoso.com", "includeInherited": true }
```

#### `sync_folder`
Bidirectional folder sync between a local path and a remote drive folder.
```json
{
  "site": "primary",
  "localPath": "/Users/me/work/Q4",
  "remotePath": "/Shared Documents/Reports/Q4",
  "direction": "bidirectional",
  "conflictResolution": "newer",
  "includePatterns": ["*.xlsx", "*.docx"],
  "excludePatterns": ["~*", "*.tmp"],
  "deleteOrphans": false
}
```
Conflict policies: `local`, `remote`, `newer`, `rename`. `deleteOrphans: true` removes items that exist only on one side — irreversible, use with care.

#### `batch_file_operations`
Up to 50 file ops in one call (uploads, downloads, moves, copies, deletes, renames). Each op runs through the existing file handlers, so the same site/drive resolution applies. With `parallel: true` Graph throttling is the concern; with `stopOnError: true` the run halts at the first failure.

```json
{
  "site": "primary",
  "operations": [
    { "operation": "upload", "source": "/tmp/a.txt", "destination": "/Inbox/a.txt" },
    { "operation": "delete", "itemId": "01ABC..." },
    { "operation": "rename", "itemId": "01DEF...", "newName": "renamed.txt" }
  ],
  "stopOnError": false,
  "parallel": false
}
```

#### `storage_analytics`
Analyze storage usage on a drive or subfolder.
```json
{
  "site": "primary",
  "analysisType": "large_files",
  "thresholds": { "largeFileSize": 250, "oldFileDays": 730 }
}
```
`analysisType` ∈ `summary`, `detailed`, `duplicates`, `large_files`, `old_files`, `file_types`.

#### `version_management`
List, restore, delete, clean up, or compare versions of a single file.
```json
{ "site": "primary", "itemPath": "/x.xlsx", "action": "list" }
{ "site": "primary", "itemId": "01ABC...", "action": "cleanup", "keepVersions": 10 }
```

#### `excel_operations`
Read/write workbook content. Eight operations: `read_range`, `write_range`, `add_worksheet`, `list_worksheets`, `get_formulas`, `set_formulas`, `create_table`, `create_chart`. Pass `useSession: true` for multi-op flows so all reads/writes share a workbook session (cheaper and more consistent).

```json
{
  "site": "primary",
  "itemPath": "/Shared Documents/Reports/Q4.xlsx",
  "operation": "write_range",
  "worksheet": "Summary",
  "range": "B2:C4",
  "values": [["Revenue", 12500], ["Cost", 7800], ["Margin", 4700]],
  "useSession": true
}
```

#### `excel_analysis`
Read-mostly analytics on a workbook.
```json
{ "site": "primary", "itemPath": "/x.xlsx", "analysisType": "statistics", "worksheet": "Data", "range": "A1:E1000" }
```
`analysisType` ∈ `statistics`, `pivot_summary`, `data_validation`, `named_ranges`, `used_range`.

---

## Common workflows

### Find a site, then list its top-level files
```bash
ods discover_sites --search="finance" --limit=5
# copy the id of the one you want into sites.local.json with an alias, then:
ods list_files --site=finance --path=/
```

### Bulk-upload a directory tree
```bash
ods sync_folder --json '{
  "site": "primary",
  "localPath": "/Users/me/exports",
  "remotePath": "/Shared Documents/Exports",
  "direction": "upload",
  "conflictResolution": "rename"
}'
```

### Atomically read-then-write an Excel range
```bash
ods excel_operations --json '{
  "site": "primary",
  "itemPath": "/budget.xlsx",
  "operation": "read_range",
  "worksheet": "Inputs",
  "range": "A1:C10",
  "useSession": true
}'

ods excel_operations --json '{
  "site": "primary",
  "itemPath": "/budget.xlsx",
  "operation": "write_range",
  "worksheet": "Outputs",
  "range": "A1:A1",
  "values": [["=SUM(Inputs!A:A)"]],
  "useSession": true
}'
```
Sessions are cached per `(itemId, persistChanges)` for the lifetime of the process.

### Migrate items between two SharePoint document libraries
```bash
ods copy_item --json '{
  "site": "old-archive",
  "itemId": "01XYZ...",
  "destinationSite": "new-archive",
  "destinationFolderPath": "/Imported/2026"
}'
```
Cross-site copies are asynchronous on Graph's side — the tool reports "Item copy initiated successfully" once the request is accepted.

### Create a list item from a script
```bash
ods create_list_item --json '{
  "site": "primary",
  "listId": "9c6b8b70-0000-0000-0000-111111111111",
  "fields": { "Title": "Renew SSL cert", "Status": "Open", "DueDate": "2026-09-01" }
}'
```

---

## Operational notes

- **One-shot vs long-lived**: prefer `spcall` / `ods` for ad-hoc use so the process exits when the call returns. Keeping the stdio server bound to Claude Code permanently can accumulate idle child processes; the `spcall` wrapper does `pkill` cleanup on EXIT.
- **Token lifetime**: in device-code mode, MSAL refresh tokens last 90 days of inactivity. If `silent token refresh failed` shows up in logs, re-run `npm run setup-auth`. In CC mode the in-memory token is refreshed automatically when it expires.
- **Caching**: the metadata/search/drive caches are in-process only, so each new process starts cold. Persistent caching is out of scope.
- **Throttling**: Graph 429 responses are honoured via `retry-after`. The retry budget defaults to 3 (configurable per call via `RequestOptions.retries`). Upload chunks share the same retry policy.
- **Chunk size**: the resumable-upload chunk is `10 * 320 * 1024 ≈ 3.13 MiB`. This is a multiple of the mandatory 320 KiB stride, but it sits below Microsoft's recommended 5–10 MiB band. If you regularly upload very large files on a fast link, raise it in `src/graph/client.ts` (look for `chunkSize`).

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `Authentication required. Run npm run setup-auth before calling tools` | Device-code cache is empty / stale. | `npm run setup-auth` again. Or, if you meant to use CC mode, ensure `MICROSOFT_GRAPH_CLIENT_SECRET` is set in the env the server sees. |
| `AADSTS7000215` / `invalid client secret` | Secret expired or was rotated. | Mint a new secret in Azure AD and update env. |
| `AADSTS700016` / `invalid client` with CC | Tenant id is `common` or a typo. | Must be a specific tenant UUID. |
| `403 Forbidden` from a SharePoint site | App lacks site permissions / admin consent not granted. | Check Entra → Enterprise Applications → your app → Permissions. |
| `404 itemNotFound` on a known site | `siteId` or `driveId` is stale / from a different tenant. | Re-run `discover_sites` / `list_drives` and update `sites.local.json`. |
| Tools that hit `/me` return errors in CC mode | App-only auth has no user. | Switch to `siteId` / `driveId` / `site` targeting. |
| `Refusing to fall back to personal OneDrive` | A `site` reference didn't resolve. | Check it against `resolve_site`; if it's correct, ensure the registry has an entry or that the URL resolves via Graph. |
| `silent token refresh failed` in logs | Keychain entry is corrupt or 90-day window expired. | `npm run setup-auth`. Or `security delete-generic-password -s mcp-onedrive-sharepoint -a msal_token_cache` on macOS then re-run setup-auth. |
| Build fails with engine warnings about Node 23/24 | ESLint 9 marks Node 23/24 as unsupported but the runtime is fine. | Safe to ignore. Use Node 20 or 22 LTS if you want clean installs. |
| `npm audit` flags ReDoS in `minimatch` | You're on the pre-v1.1 dependency tree. | `npm install` again after pulling the v1.1 deps refresh — that chain was upgraded. |

If something here is wrong or missing, file an issue against the fork or open a PR.
