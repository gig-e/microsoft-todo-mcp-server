# Microsoft To Do MCP

[![CI](https://github.com/jordanburke/microsoft-todo-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/jordanburke/microsoft-todo-mcp-server/actions/workflows/ci.yml)
[![npm version](https://badge.fury.io/js/microsoft-todo-mcp-server.svg)](https://www.npmjs.com/package/microsoft-todo-mcp-server)

A Model Context Protocol (MCP) server that enables AI assistants like Claude and Cursor to interact with Microsoft To Do via the Microsoft Graph API. This service provides comprehensive task management capabilities through a secure OAuth 2.0 authentication flow.

## Features

- **21 MCP Tools**: Complete task management functionality including lists, tasks, checklist items, cross-list search, and organization features
- **Configurable Access Level**: Expose the server as read-only, read/write, or full control — see [Access Modes](#access-modes)
- **Seamless Authentication**: Automatic token refresh with zero manual intervention
- **OAuth 2.0 Authentication**: Secure authentication with automatic token refresh
- **Microsoft Graph API Integration**: Direct integration with Microsoft's official API
- **Multi-tenant Support**: Works with personal, work, and school Microsoft accounts
- **TypeScript**: Fully typed for reliability and developer experience
- **ESM Modules**: Modern JavaScript module system

## Prerequisites

- Node.js 22.13 or higher
- pnpm package manager
- A Microsoft account (personal, work, or school)
- Azure App Registration (see setup below)

## Installation

### Option 1: Global Installation (Recommended)

```bash
# Install globally using npm
npm install -g microsoft-todo-mcp-server

# Or using pnpm
pnpm install -g microsoft-todo-mcp-server

# Or run directly with npx (no installation)
npx microsoft-todo-mcp-server
```

The package provides these commands:

- `microsoft-todo-mcp-server` - Full package name
- `mstodo` - Short alias for the MCP server
- `mstodo-setup` - Interactive setup: credentials, sign-in, access mode, client config
- `mstodo-auth` - One-shot interactive sign-in on its own
- `mstodo-config` - Configuration helper tool

With a global install there is no repo to `cd` into, so `mstodo-setup` writes `.env` to the
per-user config directory (`%APPDATA%\microsoft-todo-mcp\.env` /
`~/.config/microsoft-todo-mcp/.env`), which survives `npm i -g` upgrades. In a git
checkout it writes to the repo root instead. Both locations are read at startup, and real
environment variables — such as the `env` block in your MCP client config — win over both.

### Option 2: Clone and Run Locally

```bash
git clone https://github.com/jordanburke/microsoft-todo-mcp-server.git
cd microsoft-todo-mcp-server
pnpm install
pnpm run build
```

## Azure App Registration

This app authenticates as a **public client** using the OAuth 2.0 authorization code
flow with PKCE — there is no client secret.

### Option A: Azure Portal

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to "App registrations" and create a new registration
3. Name your application (e.g., "To Do MCP")
4. For "Supported account types", select one of the following based on your needs:
   - **Accounts in this organizational directory only (Single tenant)** - For use within a single organization
   - **Accounts in any organizational directory (Any Azure AD directory - Multitenant)** - For use across multiple organizations
   - **Accounts in any organizational directory and personal Microsoft accounts** - For both work accounts and personal accounts
5. Under "Authentication", add a platform of type **"Mobile and desktop applications"** and add
   `http://localhost` as a redirect URI. Do **not** use the "Web" platform type — that requires a
   client secret, which this app does not use.
6. Go to "API permissions" and add the following permissions:
   - Microsoft Graph > Delegated permissions:
     - Tasks.Read
     - Tasks.Read.Shared
     - Tasks.ReadWrite
     - Tasks.ReadWrite.Shared
     - User.Read
7. Click "Grant admin consent" for these permissions (org tenants only — see note below)

> Migrating an existing "Web" platform registration? Add a "Mobile and desktop applications"
> platform alongside it (or switch to it) — a client secret is no longer read or required.

### Option B: Azure CLI

Requires `az login` first. This creates the same registration as Option A.

```bash
# 1. Create the app as a public client (no secret). Pick the sign-in audience that
#    matches who should be able to use it:
#      AzureADMyOrg                        -> single tenant  (TENANT_ID=<your-tenant-id>)
#      AzureADMultipleOrgs                 -> multi-tenant   (TENANT_ID=organizations)
#      AzureADandPersonalMicrosoftAccount  -> both           (TENANT_ID=common)
#      PersonalMicrosoftAccount            -> personal only  (TENANT_ID=consumers)
APP_ID=$(az ad app create \
  --display-name "To Do MCP" \
  --sign-in-audience AzureADandPersonalMicrosoftAccount \
  --is-fallback-public-client true \
  --public-client-redirect-uris "http://localhost" \
  --query appId -o tsv)

echo "CLIENT_ID=$APP_ID"

# 2. Add the required Microsoft Graph delegated permissions.
#    (Microsoft Graph's resource appId, 00000003-0000-0000-c000-000000000000, and the
#    scope GUIDs below are fixed platform values — the same for every tenant/app.)
az ad app permission add --id "$APP_ID" \
  --api 00000003-0000-0000-c000-000000000000 \
  --api-permissions \
    e1fe6dd8-ba31-4d61-89e7-88639da4683d=Scope \
    f45671fb-e0fe-4b4b-be20-3d3ce43f1bcb=Scope \
    2219042f-cab5-40cc-b0d2-16b1540b4c5f=Scope \
    88d21fd4-8e5a-4c32-b5e2-4a1c95f34f72=Scope \
    c5ddf11b-c114-4886-8558-8a4e557cd52b=Scope
# (User.Read, Tasks.Read, Tasks.ReadWrite, Tasks.Read.Shared, Tasks.ReadWrite.Shared, respectively)

# 3. Grant admin consent — only applicable/needed for AzureADMyOrg / AzureADMultipleOrgs
#    audiences, and only if you're a tenant admin. For AzureADandPersonalMicrosoftAccount
#    or PersonalMicrosoftAccount audiences, personal-account users consent themselves the
#    first time they sign in (during `pnpm run auth`) — skip this step for those.
az ad app permission admin-consent --id "$APP_ID"
```

### Option C: Microsoft Graph API directly

Equivalent to Option B, via a raw Graph call (e.g. through `az rest`, `curl` with a bearer
token, or Graph Explorer) — useful if you're not using the Azure CLI:

```bash
az rest --method POST \
  --uri https://graph.microsoft.com/v1.0/applications \
  --body '{
    "displayName": "To Do MCP",
    "signInAudience": "AzureADandPersonalMicrosoftAccount",
    "isFallbackPublicClient": true,
    "publicClient": { "redirectUris": ["http://localhost"] },
    "requiredResourceAccess": [
      {
        "resourceAppId": "00000003-0000-0000-c000-000000000000",
        "resourceAccess": [
          { "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d", "type": "Scope" },
          { "id": "f45671fb-e0fe-4b4b-be20-3d3ce43f1bcb", "type": "Scope" },
          { "id": "2219042f-cab5-40cc-b0d2-16b1540b4c5f", "type": "Scope" },
          { "id": "88d21fd4-8e5a-4c32-b5e2-4a1c95f34f72", "type": "Scope" },
          { "id": "c5ddf11b-c114-4886-8558-8a4e557cd52b", "type": "Scope" }
        ]
      }
    ]
  }'
```

This requires `Application.ReadWrite.All` (or equivalent) on whatever credential is
calling Graph. Admin consent still needs a separate call/step as in Option B if the
audience includes org accounts.

## Configuration

### Environment Setup

Create a `.env` file in the project root (required for authentication):

```env
CLIENT_ID=your_client_id
TENANT_ID=your_tenant_setting
```

### TENANT_ID Options

- `organizations` - For multi-tenant organizational accounts (default if not specified)
- `consumers` - For personal Microsoft accounts only
- `common` - For both organizational and personal accounts
- `your-specific-tenant-id` - For single-tenant configurations

**Examples:**

```env
# For multi-tenant organizational accounts (default)
TENANT_ID=organizations

# For personal Microsoft accounts
TENANT_ID=consumers

# For both organizational and personal accounts
TENANT_ID=common

# For a specific organization tenant
TENANT_ID=00000000-0000-0000-0000-000000000000
```

### Access Modes

Authentication grants this server your full `Tasks.ReadWrite` scope, so by default an
assistant can delete lists and tasks as easily as it can read them. `MSTODO_ACCESS_MODE`
narrows the tool surface:

| Mode    | Tools | What the assistant can do                                                    |
| ------- | ----- | ---------------------------------------------------------------------------- |
| `read`  | 10    | Query lists, tasks, checklists, search and agenda. Changes nothing.          |
| `write` | 17    | Everything in `read`, plus create and update. Cannot delete.                 |
| `full`  | 21    | Everything, including `delete-*` and `archive-completed-tasks`. **Default.** |

Set it in the MCP server config alongside the command:

```json
{
  "mcpServers": {
    "microsoftTodo": {
      "command": "npx",
      "args": ["--yes", "microsoft-todo-mcp-server"],
      "env": { "MSTODO_ACCESS_MODE": "write" }
    }
  }
}
```

Withheld tools are never registered, so they don't appear in `tools/list` at all — the
assistant can't attempt them and be refused, it simply doesn't know they exist. Run
`auth-status` to see the active mode and exactly which tools are being withheld.

Notes:

- Unset means `full`, so upgrading an existing install doesn't silently lose tools.
- An unrecognised value is fatal at startup rather than falling back — a typo like
  `MSTODO_ACCESS_MODE=raed` must not quietly grant full control.
- `archive-completed-tasks` counts as destructive: it deletes from the source list after
  copying, so it needs `full` even though it's framed as a move.
- Aliases are accepted: `read-only` → `read`, `read-write`/`update` → `write`, `all` → `full`.
- This is a guard rail against accidents, not a security boundary — the token in the
  encrypted cache still carries full scope, and anything that can edit the MCP config can
  change the mode.

### Token Storage

Authentication tokens are never written into this project directory (which matters if it's
synced by OneDrive, Dropbox, etc.). Instead, the MSAL token cache is stored **encrypted at
rest** in a per-user, per-machine location:

- **Windows**: `%APPDATA%\microsoft-todo-mcp\token-cache.bin`
- **macOS/Linux**: `~/.config/microsoft-todo-mcp/token-cache.bin`

The encryption key is derived from a machine identifier, so the file is only usable on the
machine that created it — copying it to another machine (e.g. via cloud sync) yields
undecryptable bytes, not a portable credential. Refresh happens silently and automatically;
there's nothing to configure. You can override the cache file location:

```bash
export MSTODO_TOKEN_FILE=/path/to/custom/token-cache.bin
```

Because credentials aren't portable between machines, run `pnpm run auth` once on each
machine you use this server from.

## Usage

### Complete Setup Workflow

#### Step 1: Authenticate with Microsoft

```bash
# In a git checkout
pnpm run auth

# Installed globally
mstodo-auth
```

This opens a browser window for Microsoft authentication and stores your session in the
encrypted per-machine token cache described above. It's one-shot — it exits once you've
signed in, and binds no port of its own (MSAL's loopback callback listens on 127.0.0.1
only). Run it once on every machine, since the cache can't be copied between them.

`mstodo-setup` runs this step for you as part of the guided flow, so use one or the other.

#### Step 2: Create MCP Configuration

```bash
# Generate MCP configuration file
pnpm run create-config

# Or use the global helper (if installed globally)
mstodo-config

# Optionally pick an output path and access mode (see Access Modes above)
mstodo-config ./mcp.json write
```

This creates an `mcp.json` file. No tokens are embedded in it — the server reads its own
encrypted per-machine token store automatically, so `pnpm run auth` must be run once on
each machine you deploy this to. `mstodo-setup` asks for the access mode interactively.

#### Step 3: Configure Your AI Assistant

**For Claude Desktop:**

Add to your configuration file:

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "microsoftTodo": {
      "command": "npx",
      "args": ["--yes", "microsoft-todo-mcp-server"],
      "env": { "MSTODO_ACCESS_MODE": "full" }
    }
  }
}
```

No tokens go in this config — run `pnpm run auth` once on this machine first. Change
`MSTODO_ACCESS_MODE` to `read` or `write` to narrow what the assistant can do.

**For Cursor:**

```bash
# Copy to Cursor's global configuration
cp mcp.json ~/.cursor/mcp-servers.json
```

### Available Scripts

```bash
# Development & Building
pnpm run build        # Build TypeScript to JavaScript
pnpm run dev          # Build and run CLI in one command

# Running the Server
pnpm start            # Run MCP server directly
pnpm run cli          # Run MCP server via CLI wrapper
npx microsoft-todo-mcp-server  # Run globally installed version

# Authentication & Configuration
pnpm run auth         # Run interactive PKCE sign-in (opens browser)
pnpm run create-config # Generate mcp.json (no tokens embedded)

# Code Quality
pnpm run format       # Format code with Prettier
pnpm run format:check # Check code formatting
pnpm run lint         # Run linting checks
pnpm run typecheck    # TypeScript type checking
```

## MCP Tools

The server provides 21 tools for comprehensive Microsoft To Do and Planner management. The
badge after each tool is the [access mode](#access-modes) it requires: 🟢 `read`,
🟡 `write`, 🔴 `full`.

### Authentication

- **`auth-status`** 🟢 - Check authentication status, token expiration, account type, and the active access mode

### Task Lists (Top-level Containers)

- **`get-task-lists`** 🟢 - Retrieve all task lists with metadata (default, shared, etc.)
- **`get-task-lists-organized`** 🟢 - Group lists into folders inferred from naming patterns, emoji prefixes, and sharing status
- **`create-task-list`** 🟡 - Create a new task list
- **`update-task-list`** 🟡 - Rename an existing task list
- **`delete-task-list`** 🔴 - Delete a task list and all its contents

### Tasks (Main Todo Items)

- **`get-tasks`** 🟢 - Get tasks from a list with filtering, sorting, and pagination
  - Supports OData query parameters: `$filter`, `$select`, `$orderby`, `$top`, `$skip`, `$count`
- **`create-task`** 🟡 - Create a new task with full property support
  - Title, description, due date, start date, importance, reminders, status, categories
- **`update-task`** 🟡 - Update any task properties
- **`delete-task`** 🔴 - Delete a task and all its checklist items

### Cross-List Views

Graph scopes every task query to a single list, so these two fan out across your lists and
combine the results — no `listId` needed. Both accept list IDs _or_ list names in
`listIds`, and both report any lists that couldn't be read so an empty result is never
mistaken for "nothing there".

- **`search-tasks`** 🟢 - Find tasks by text across every list at once
  - Matches title and categories (and descriptions with `searchBody`), case-insensitively
  - All whitespace-separated terms must appear, in any order: `tax invoice` finds "Invoice for tax return"
  - Filters: `importance`, `dueBefore`/`dueAfter`, `includeCompleted`, `limit`
- **`get-agenda`** 🟢 - Roll everything due into overdue / today / tomorrow / upcoming sections
  - `days` sets how far ahead to look (default 7); overdue is never windowed out
  - `timeZone` (IANA name) decides which day is "today", defaulting to `MSTODO_TIMEZONE` then the server's local zone
  - Due dates are read as the calendar day To Do stored, not converted from the instant — so a task due Aug 5 doesn't show as Aug 4 in a negative-offset zone

### Checklist Items (Subtasks)

- **`get-checklist-items`** 🟢 - Get subtasks for a specific task
- **`create-checklist-item`** 🟡 - Add a new subtask to a task
- **`update-checklist-item`** 🟡 - Update subtask text or completion status
- **`delete-checklist-item`** 🔴 - Remove a specific subtask

### Maintenance

- **`archive-completed-tasks`** 🔴 - Move completed tasks older than N days into an archive list (copies, then deletes from the source — supports `dryRun`)
- **`test-graph-api-exploration`** 🟢 - Probe Graph for undocumented folder/grouping properties

### Microsoft Planner

Powers the "Assigned to me" view in the To Do app — a separate API from native To Do
lists, covering tasks from Planner plans (e.g. project boards) you're assigned to. Uses
the same `Tasks.Read`/`Tasks.ReadWrite` scopes already granted; no extra consent needed.

- **`get-assigned-planner-tasks`** 🟢 - Get Planner tasks assigned to you across all plans
- **`get-planner-task-details`** 🟢 - Get the description and checklist for a specific Planner task
- **`update-planner-task`** 🟡 - Update progress, title, priority, or dates on a Planner task (handles Planner's required ETag concurrency check automatically)

## Architecture

### Project Structure

- **MCP Server** (`src/todo-index.ts`) - Core server implementing the MCP protocol
- **Access Modes** (`src/access-mode.ts`) - Parses `MSTODO_ACCESS_MODE` and decides which tools get registered
- **Cross-List Helpers** (`src/agenda.ts`) - Due-date bucketing, query matching, and the fan-out concurrency cap
- **CLI Wrapper** (`src/cli.ts`) - Executable entry point
- **Auth Flow** (`src/auth-server.ts`) - One-shot interactive PKCE sign-in via MSAL Node
- **MSAL Client** (`src/msal-client.ts`) - Shared `PublicClientApplication` factory and encrypted cache plugin
- **Token Manager** (`src/token-manager.ts`) - Silent token acquisition against the encrypted cache
- **Crypto Store** (`src/crypto-store.ts`) - AES-256-GCM encryption of the token cache, machine-bound key
- **Config Generator** (`src/create-mcp-config.ts`) - Helper to create MCP configurations

### Technical Details

- **Microsoft Graph API**: Uses v1.0 endpoints, with bounded-concurrency fan-out and Retry-After-aware retries on 429/503 throttling
- **Authentication**: MSAL Node `PublicClientApplication`, authorization code flow with PKCE (no client secret)
- **Token Storage**: Encrypted at rest, machine-bound, stored outside the project directory; automatic silent refresh
- **Build System**: ts-builds (tsdown) for fast TypeScript compilation
- **Module System**: ESM (ECMAScript modules)

## Limitations & Known Issues

### Personal Microsoft Accounts

- **MailboxNotEnabledForRESTAPI Error**: Personal Microsoft accounts (outlook.com, hotmail.com, live.com) have limited access to the To Do API through Microsoft Graph
- This is a Microsoft service limitation, not an issue with this application
- Work/school accounts have full API access

### API Limitations

- Rate limits apply according to Microsoft's policies
- Some features may be unavailable for personal accounts
- Shared lists have limited functionality

## Troubleshooting

### Authentication Issues

**Token acquisition failures**

- Verify `CLIENT_ID` and `TENANT_ID` in your `.env` file
- Ensure the Azure app registration has a "Mobile and desktop applications" platform with
  `http://localhost` as a redirect URI (not a "Web" platform — that expects a client secret)
- Check Azure App permissions are granted with admin consent

**Permission issues**

- Ensure all required Graph API permissions are added and consented
- For organizational accounts, admin consent may be required

### Account Type Configuration

**Work/School Accounts**

```env
TENANT_ID=organizations  # Multi-tenant
# Or use your specific tenant ID
```

**Personal Accounts**

```env
TENANT_ID=consumers  # Personal only
# Or TENANT_ID=common for both types
```

### Debugging

**Check authentication status:**

```bash
# Using the MCP tool
# In your AI assistant: "Check auth status"
```

The token cache is encrypted at rest and machine-bound, so it can't be inspected directly
with `cat`/`jq` — use the `auth-status` MCP tool, or re-run `pnpm run auth` if in doubt.

**Enable verbose logging:**

```bash
# The server logs to stderr for debugging
mstodo 2> debug.log
```

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Run `pnpm run lint` and `pnpm run typecheck` before submitting
4. Submit a pull request

## License

MIT License - See [LICENSE](LICENSE) file for details

## Acknowledgments

- Fork of [@jhirono/todomcp](https://github.com/jhirono/todomcp)
- Built on the [Model Context Protocol SDK](https://github.com/modelcontextprotocol/sdk)
- Uses [Microsoft Graph API](https://developer.microsoft.com/en-us/graph)

## Support

- [GitHub Issues](https://github.com/jordanburke/microsoft-todo-mcp-server/issues)
- [npm Package](https://www.npmjs.com/package/microsoft-todo-mcp-server)
