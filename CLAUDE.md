# CLAUDE.md

Context for Claude Code sessions working in this repo. Read this first; for the long-form tool reference fall back to `docs/USAGE.md`, and for the wrapper scripts to `docs/operations.md`.

## What this is

A fork of [`ftaricano/mcp-onedrive-sharepoint`](https://github.com/ftaricano/mcp-onedrive-sharepoint), patched and dependency-refreshed (v1.1). Exposes 33 MCP tools that drive Microsoft Graph for OneDrive / SharePoint operations. This fork is configured for an internal Lanpro Microsoft 365 tenant — operator identity, tenant ID, and registered site aliases live in machine-local files (see [Local-only state](#local-only-state) below); none of that is committed.

The operator owns the tenant and has full admin rights (can grant admin consent, create app registrations, modify tenant-wide settings). Don't gate suggestions on "ask your IT admin"; if something needs an Azure portal click, walk them through it directly.

---

## ⚠️ Spanish-only content rule

**Lanpro is a Spanish-speaking company.** Any artifact this MCP *persists into SharePoint* must be in Spanish:

- New folder names and file names you create or upload
- SharePoint list item field values (Title, Description, Status, etc.)
- Sharing-invitation messages (`advanced_share.message`)
- Comments, descriptions, and any other user-facing text written to Graph

This rule applies to **outputs**, not to the working language of the session. Conversational replies to the user, code comments, commit messages, and log output stay in English. The rule is: anything that ends up in someone's SharePoint UI must be Spanish.

Edge cases:
- **File contents** (uploading a `.docx`/`.xlsx`/`.pdf` the user provides): use the file as-is, don't translate its content. Only the *filename* and destination folder name need to be Spanish if you're naming them.
- **User pastes English text** and asks you to put it in SharePoint: ask whether to translate first, or use as-is. Default to asking when the content is more than a few words.
- **Renaming an existing English-named file:** make the new name Spanish unless the user says otherwise.

When you genuinely don't know the right Spanish term for a domain concept, ask. Don't guess on technical vocabulary the org might already have a standard term for.

---

## Local-only state

These three sources hold tenant- and operator-specific configuration. They are deliberately gitignored / outside the repo so a public fork can't leak them. **Read them at session start instead of asking the user to re-state the context.**

| What | Where | Holds |
|---|---|---|
| Auth config | `.env` (gitignored) | `MICROSOFT_GRAPH_CLIENT_ID`, `MICROSOFT_GRAPH_TENANT_ID`, optional `MICROSOFT_GRAPH_CLIENT_SECRET` |
| MSAL token cache | macOS Keychain service `mcp-onedrive-sharepoint` | Device-code refresh token (90-day window) |
| Site registry | `config/sites.local.json` (gitignored) | Short aliases for the org's SharePoint sites; pass them as `site=<alias>` |
| Persistent context | `~/.claude/projects/-Users-alex-local-files-lanpro-sharepoint-mcp-onedrive-sharepoint/memory/` | `user_role`, `project-tenant`, `project-site-registry` |

If `.env` is missing, the auth flow won't start. If `sites.local.json` is missing, the registry is empty but the tools still accept raw `siteId`/`siteUrl`/`driveId`. If memory is empty, read `.env` and `sites.local.json` to rehydrate the context.

---

## Wired into Claude Code

This repo is registered as the `sharepoint` MCP server at user scope in `~/.claude.json`. Tools appear under that namespace in every Claude Code session. If the entry is missing on a fresh machine, add:
```json
{ "mcpServers": { "sharepoint": { "type": "stdio", "command": "<absolute path>/scripts/run-stdio.sh", "args": [], "env": {} } } }
```

---

## How to run things

```bash
npm run build       # tsc, no emit on errors
npm run test        # node --test, 106 tests pass on main
npm run lint        # eslint 9 flat config, --max-warnings 410
npm run ci          # build + lint + test
npm run setup-auth  # only when the MSAL refresh window has expired
./scripts/spcall.sh <tool> [key=value ...]   # one-shot tool call from shell
```

`spcall` launches the MCP through `mcporter`, runs one tool, kills the child. Use it for debugging or scripting; the registered Claude Code MCP is the right path for interactive work.

---

## Where things live

| Area | Path |
|---|---|
| Auth (device-code + client-credentials, keychain + file fallback) | `src/auth/microsoft-graph-auth.ts` |
| Singleton init (auth-mode-aware user check) | `src/core/bootstrap.ts` |
| Graph HTTP client (retries, chunked upload, `$batch`, caching) | `src/graph/client.ts` |
| Endpoint builders + OData escaping | `src/graph/resource-resolver.ts` |
| Site-registry + alias resolution | `src/sharepoint/site-resolver.ts` |
| Tools (33 total) | `src/tools/{files,sharepoint,utils,advanced}/*.ts` |
| Tests (helpers, fixtures) | `src/tests/`, `src/tests/helpers/` |
| Long-form tool reference | `docs/USAGE.md` |
| Wrapper / spcall docs | `docs/operations.md` |

---

## Conventions to keep

- **No new `any`** unless interfacing with an existing `any` boundary. Lint allows it as a warning to track existing debt — don't add to the pile.
- **Dependency injection over module-level singletons in tests.** `__set*ForTests` hooks exist for the auth instance, the graph client, the bootstrap dependencies, the utility deps, and the site registry. Use those instead of stubbing module imports.
- **Tool handlers always return through `jsonTextResponse` / `toolErrorResponse`.** Don't hand-roll the MCP content envelope.
- **OData string parameters must go through `escapeODataString`.** `encodeURIComponent` alone leaves `'` unescaped and that's a query-corruption vector.
- **Don't commit tenant identifiers, site IDs, drive IDs, folder paths, or operator email/UPN.** `.env`, `config/sites.local.json`, and the macOS Keychain entries are gitignored for this reason. New examples in committed docs/tests use anonymized placeholders (`yourtenant.sharepoint.com`, `b!YOUR_DRIVE_ID_HERE`, `alice@example.com`).

---

## v1.1 fork notes

This fork ships fixes the upstream doesn't yet:
- Client-credentials bootstrap regression fixed (`bootstrap.ts` gates the cached-user check on `!clientSecret`).
- `copy_item` cross-site bug fixed (`destinationSiteId` is properly resolved to its drive id rather than mis-assigned to `parentReference.driveId`).
- `move_item` errors loudly on a bad `parentFolderPath` instead of silently no-op'ing.
- `health_check` reports real auth mode and skips `/me` probes when running app-only.
- `keytar` (archived March 2026) → `@github/keytar`. `@azure/msal-node` 2.x → 5.x (Node ≥ 20).
- ESLint 8 → 9 flat config; clears 6 high-severity `minimatch` ReDoS advisories from the `@typescript-eslint` 6.x chain.

See `git log --oneline` for the four-commit landing on `main`.

---

## When in doubt

- Tool inputs / outputs: `docs/USAGE.md`.
- Wiring the MCP into other clients or running from shell: `docs/operations.md`.
- The operator owns the tenant. If a Graph call returns `Permission` / `AccessDenied`, the fix is almost always in Azure (consent / app permissions) rather than in this code. Walk through the Azure portal steps instead of suggesting workarounds.
- Don't permanently bind this MCP in clients that load it lazily. Claude Code's user-scope registration is fine because Claude Code starts it on demand and tears it down. Prefer `spcall` for any other shell context.
