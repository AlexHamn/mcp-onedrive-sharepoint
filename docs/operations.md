# Operational SharePoint usage

This document is the practical, "I just want to run a tool" companion to [USAGE.md](USAGE.md). It covers the shell wrappers shipped in `scripts/`, the recommended one-shot lifecycle, and how to register your own SharePoint sites with short aliases instead of memorizing GUIDs.

> The previous version of this file was bound to a specific tenant (CPZ Seguros) with real site IDs and folder paths embedded in the examples. That tenant-specific content has been removed and replaced with anonymized examples. If you forked the upstream repo and need the original values, recover them from git history before this commit.

## What the wrappers do

This repository includes lightweight wrappers for day-to-day SharePoint operations:

- `./scripts/run-stdio.sh` — launches the MCP stdio server from this repo with the repo-local `.env` parsed safely.
- `./scripts/spcall.sh` — runs ad-hoc `mcporter call` requests against the local MCP server.
- `npm run stdio` — same as `./scripts/run-stdio.sh`.
- `npm run spcall -- <tool> ...` — same as `./scripts/spcall.sh <tool> ...`.

The wrappers do **not** `source` `.env` in the shell. They parse the file with `dotenv`, merge it into the process environment, and keep explicit environment variables higher priority.

## Process lifecycle rule

This MCP is operationally on-demand only.

- Do not keep it registered as a permanently loaded MCP in your MCP client.
- Prefer `spcall` or ad-hoc `mcporter call --stdio ...`.
- Each call should start the MCP process, use it for that request, and let it exit immediately afterward.
- `spcall` runs a cleanup trap after each invocation to kill leftover repo-local MCP processes if any child fails to exit cleanly.
- The goal is zero long-lived idle MCP processes and no zombie leftovers after routine usage.

In practice, `spcall` is the canonical path because it launches an ephemeral stdio MCP process through `mcporter`, returns the result, and exits.

## Wrapper usage

### Start the MCP server over stdio

```bash
npm run build
./scripts/run-stdio.sh
```

Equivalent:

```bash
npm run stdio
```

This is the wrapper to use in MCP client configs because it always launches from the repo root and loads the repo-local `.env` safely.

### Ad-hoc tool calls with `spcall`

```bash
npm run build
./scripts/spcall.sh health_check
./scripts/spcall.sh list_drives
./scripts/spcall.sh list_files \
  driveId=b!YOUR_DRIVE_ID_HERE \
  path=/Shared%20Documents/Reports
```

Equivalent npm form:

```bash
npm run spcall -- health_check
npm run spcall -- list_files \
  driveId=b!YOUR_DRIVE_ID_HERE \
  path=/Shared%20Documents/Reports
```

Behavior notes:

- defaults to `--output json` unless you pass your own `--output ...`
- uses `mcporter call --stdio` so the MCP process is ephemeral
- passes `--cwd` as the repo root so local build artifacts and `.env` resolution stay consistent
- URL-encode spaces in path arguments (`%20`) so the shell doesn't split them

## Setting up canonical site aliases

Tools accept site references as raw `siteId`, canonical `siteUrl`, or short `site` alias. Aliases come from a local registry file so tenant-specific IDs never get committed.

1. Copy the example file:
   ```bash
   cp config/sites.example.json config/sites.local.json
   ```
2. Discover your sites via the MCP:
   ```bash
   ./scripts/spcall.sh discover_sites search=YourTeam
   ```
3. Look up each site's default drive id:
   ```bash
   ./scripts/spcall.sh list_drives \
     siteId=yourtenant.sharepoint.com,<site-guid>,<web-guid>
   ```
4. Add entries to `sites.local.json`:
   ```json
   {
     "sites": [
       {
         "key": "team-docs",
         "name": "Team Docs",
         "siteId": "yourtenant.sharepoint.com,<site-guid>,<web-guid>",
         "siteUrl": "https://yourtenant.sharepoint.com/sites/TeamDocs",
         "driveId": "b!<long-drive-id>",
         "aliases": ["team-docs", "team", "/sites/TeamDocs"]
       }
     ]
   }
   ```
5. After that, every tool accepts `site=team-docs` (or any of the aliases) instead of a raw GUID.

`sites.local.json` is gitignored. Set `MCP_SITES_CONFIG_PATH` to point at a different file if you want to keep the registry outside the repo.

## Site-targeted examples

These are the patterns you'll use most often. Replace the placeholder identifiers with your tenant's real values once your registry is populated.

### List the default drive of a registered site by alias

```bash
./scripts/spcall.sh list_drives site=team-docs
```

### List files at a path on a registered site

```bash
./scripts/spcall.sh list_files \
  site=team-docs \
  path=/Shared%20Documents/Reports
```

### List files in a specific document library by drive id

```bash
./scripts/spcall.sh list_files \
  driveId=b!YOUR_DRIVE_ID_HERE \
  path=/Shared%20Documents/Reports
```

### List SharePoint lists in a site

```bash
./scripts/spcall.sh list_site_lists site=team-docs
```

### List drives directly by site id

```bash
./scripts/spcall.sh list_drives \
  siteId=yourtenant.sharepoint.com,00000000-0000-0000-0000-000000000000,11111111-1111-1111-1111-111111111111
```

## Hot-path patterns

If your team keeps coming back to the same folders, document them somewhere outside this repo (a wiki page, a snippet manager, your shell rc file). A common pattern is to define shell aliases or functions next to your environment loader:

```bash
# in ~/.zshrc / ~/.bashrc / a dotfiles snippet
sp_reports() {
  ./scripts/spcall.sh list_files \
    site=team-docs \
    path=/Shared%20Documents/Reports \
    "$@"
}
```

Don't commit tenant-specific hot paths to this repo. They drift, they leak organizational structure, and they aren't portable. Keep them in your own dotfiles.

## MCP integration snippet

Use the stdio wrapper as the MCP command so the repo-local environment is loaded automatically.

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "/absolute/path/to/mcp-onedrive-sharepoint/scripts/run-stdio.sh"
    }
  }
}
```

For agent-side ad-hoc execution, prefer `spcall` when you want a one-off tool call without registering a persistent MCP server.

See [USAGE.md](USAGE.md) for the full tool reference and client-by-client wiring snippets (Claude Code, Claude Desktop, Cursor, raw `mcporter`).
