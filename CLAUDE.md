# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Common Development Commands

### Build and Development

```bash
pnpm install         # Install dependencies (pnpm 11)
pnpm run build       # Build with ts-builds (tsdown) to dist/ directory
pnpm run dev         # Build and run CLI in one command
pnpm run validate    # Full chain: format, lint, typecheck, test, build
```

Build tooling is provided by **ts-builds** (3.x); package scripts delegate to the
`ts-builds` CLI. Output extension is forced to `.js` via `tsdown.config.ts` so the
`bin`/`main` paths and `./*.js` source imports keep resolving.

### Authentication and Setup

```bash
pnpm run auth        # Run interactive PKCE sign-in (opens browser, no server/port)
pnpm run setup       # Guided setup: credentials, sign-in, access mode, client config
pnpm run create-config # Generate mcp.json (no tokens embedded)
```

Published bins: `mstodo`/`microsoft-todo-mcp-server` (server), `mstodo-setup`,
`mstodo-auth`, `mstodo-config`. Anything reachable as a global bin must resolve paths
against its own module location, never `process.cwd()`. `setup.ts` locates
`auth-server.js` as a sibling and spawns it with `process.execPath` (no shell; install
paths contain spaces).

`.env` is read from two places, package root first then `<configDir>/.env`; dotenv never
overrides an already-set variable, so the MCP client's `env` block beats both. `setup.ts`
writes to the repo root in a checkout and to the config dir for a global install, where
the package root is wiped on upgrade.

### Running the Server

```bash
pnpm run cli         # Run MCP server via CLI wrapper
pnpm start           # Same thing; both run dist/cli.js
```

`dist/cli.js` is the **only** runnable stdio entry; `bin`, `pnpm start` and `pnpm run cli`
all point at it. (The cloud connector's `dist/http-server.js` is a separate HTTP entry; see
below.) `dist/todo-index.js` is an import-only module that exports `startServer`;
running it directly does nothing on purpose. It cannot self-start: tsdown splits shared
code, so that file is a re-export shim around a hashed chunk whose `import.meta.url` never
equals `process.argv[1]`, and an `import.meta.url === process.argv[1]` guard there is dead
code that silently starts no server. Don't re-add one.

### Cloud Connector

```bash
pnpm run cloud-server     # HTTP server (dist/http-server.js), needs MSTODO_BEARER_TOKEN
pnpm run seed-cloud-token # Re-encrypt the local token cache with MSTODO_ENCRYPTION_KEY for Key Vault
```

Opt-in Azure Container Apps deployment of the same tool set over Streamable HTTP. Full
walkthrough in `docs/CLOUD_CONNECTOR.md`. Neither `http-server`, `seed-cloud-token` nor
`local-proxy` is a published bin.

## ts-builds & pnpm 11 Notes

Built with **ts-builds 3.2.0** (tsdown) on **pnpm 11**. Non-obvious, load-bearing constraints:

- **Node >= 22.13 required.** pnpm 11 crashes on Node 20 (`ERR_UNKNOWN_BUILTIN_MODULE`). CI runs Node 22.x/24.x; the publish workflow uses Node 24 (`.nvmrc`).
- **Keep `globals` in `publicHoistPattern`** (`pnpm-workspace.yaml`). `eslint.config.js` imports `globals`, which `*eslint*` does not match. Without the explicit hoist line, resolution falls back to pnpm's incidental `.pnpm` virtual-store hoist, which passes locally but **fails CI** with `Cannot find package 'globals'`. It is load-bearing, not redundant; ts-builds 3.2.0 only adds it for _new_ `init`s, it does not retro-fit this file.
- **First-party release-age excludes are globs, not pins.** `minimumReleaseAgeExclude` is `*functype*` + `ts-builds`. `pnpm add` auto-adds per-version pins for first-party packages within the 24h cooldown. Collapse them back to the glob (pins re-trip on every release).
- **`allowBuilds: esbuild: false`** is required under pnpm 11 `strictDepBuilds` (esbuild's binary ships via `@esbuild/<platform>` optional deps; build script not needed).
- **tsdown emits `.js`, not `.mjs`.** `tsdown.config.ts` forces `outExtensions: () => ({ js: ".js" })` so the `bin`/`main` paths and `./*.js` imports resolve. Don't drop it.
- **`pnpm install` in a non-TTY shell** may abort with `ERR_PNPM_ABORTED_REMOVE_MODULES_DIR_NO_TTY`; prefix with `CI=true`. Note `CI=true` makes `--frozen-lockfile` the default, so add `--no-frozen-lockfile` when intentionally updating the lockfile.

## Architecture Overview

This is a Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft To Do via the Microsoft Graph API. The codebase follows a modular architecture:

1. **MCP Server** (`src/todo-index.ts`): Core server implementing the MCP protocol with 21 tools for Microsoft To Do and Planner operations
2. **CLI Wrapper** (`src/cli.ts`): Executable entry point
3. **Auth Flow** (`src/auth-server.ts`): One-shot interactive sign-in via MSAL Node `PublicClientApplication.acquireTokenInteractive()`: authorization code + PKCE, no client secret, no hand-rolled HTTP server (MSAL's own loopback server, bound to 127.0.0.1, handles the callback)
4. **MSAL Client** (`src/msal-client.ts`): Shared `PublicClientApplication` factory, scopes, and the encrypted MSAL cache plugin
5. **Token Manager** (`src/token-manager.ts`): Silent token acquisition (`acquireTokenSilent`) against the encrypted cache; no raw refresh token is ever held by app code
6. **Crypto Store** (`src/crypto-store.ts`): AES-256-GCM encryption of the token cache with a key derived from a machine identifier (`node-machine-id`), or from `MSTODO_ENCRYPTION_KEY` when set (cloud mode). No native/compiled dependency
7. **Config Generator** (`src/create-mcp-config.ts`): Utility to create MCP configuration files (no tokens embedded)
8. **Access Modes** (`src/access-mode.ts`): Parses `MSTODO_ACCESS_MODE` and classifies each tool as read/write/destructive
9. **Cross-List Helpers** (`src/agenda.ts`): Pure due-date bucketing, query matching, and the fan-out concurrency cap behind `search-tasks` and `get-agenda`
10. **Token Store** (`src/token-store.ts`): `TokenStore` behind the MSAL cache plugin: `LocalFileTokenStore` (default) or `KeyVaultTokenStore` (`MSTODO_TOKEN_STORE=key-vault`, via `DefaultAzureCredential`)
11. **HTTP Server** (`src/http-server.ts`): Cloud connector entry: `StreamableHTTPServerTransport` on `node:http`, `/mcp` gated by a constant-time bearer check, unauthenticated `GET /health`
12. **Cloud Seeder** (`src/seed-cloud-token.ts`): Decrypts the local cache with the machine-ID key and re-encrypts it with `MSTODO_ENCRYPTION_KEY`; no Azure SDK, upload is a manual `az keyvault secret set`
13. **Local Proxy** (`src/local-proxy.ts`): stdio ↔ HTTP relay for Claude Desktop (replaces `mcp-remote`); tracks `Mcp-Session-Id` and unwraps SSE-framed responses

### Key Architectural Patterns

- **Access-mode gating**: tools are registered through `registerTool(access, ...)` in
  `todo-index.ts`, which drops any tool the configured mode disallows. Withheld tools never
  appear in `tools/list`. Gating at registration, not at call time, is deliberate.
  **Every new tool must be classified** `read` | `write` | `destructive`; "destructive"
  means it can delete user data, including indirectly (`archive-completed-tasks` deletes
  from the source list after copying). Unset mode = `full`, for backwards compatibility;
  an unparseable value is fatal at startup rather than defaulting.
- **Cross-list fan-out**: Graph scopes task queries to one list, so `search-tasks` and
  `get-agenda` query each list and merge. Concurrency is capped (`LIST_FETCH_CONCURRENCY`)
  because an unbounded `Promise.all` over ~12 lists reliably trips Graph's throttle, and a
  throttled list looks like an empty one. `makeGraphRequest` retries 429/503 honouring
  `Retry-After`.
- **Due dates are calendar days, not instants**: To Do stores a due date as midnight in the
  payload's own `timeZone`, so `dueDayKey()` reads the date off the wall clock instead of
  converting. Converting would shift the day for any negative UTC offset.
- **Token Management**: MSAL's token cache is persisted encrypted at rest, outside the project directory, in a per-user app-data location (`%APPDATA%\microsoft-todo-mcp` / `~/.config/microsoft-todo-mcp`); automatic silent refresh via `acquireTokenSilent`
- **Machine-bound encryption**: the cache decryption key is derived from a machine identifier, so the ciphertext is not portable to another machine (relevant since this directory may be cloud-synced), but it is not equivalent to OS-keychain-backed storage (DPAPI/Keychain/libsecret), which would require a native dependency
- **Cloud connector is single-session by design**: the SDK `Server` can only be connected to
  one transport, so `http-server.ts` tracks one active session and close-then-reconnects the
  shared `server` when a new client initializes (a second concurrent client evicts the
  first). This was chosen over refactoring `todo-index.ts` into a per-session server factory.
  Don't switch to the stateless transport; it can't be reused across requests.
- **`/health` must stay unauthenticated**: Container Apps' probes can't send the bearer token,
  and the default probe targets `/`; the app's probes are explicitly pointed at `/health`.
- **Claude Desktop must spawn `local-proxy.js` with `node`, not `npx`**: on Windows `npx` goes
  through `cmd.exe`, which mangles Program Files paths (the reason `mcp-remote` was dropped).
- **Multi-tenant Support**: Configurable for different Microsoft account types via TENANT_ID
- **Error Handling**: Special handling for personal Microsoft accounts (MailboxNotEnabledForRESTAPI)
- **Type Safety**: Strict TypeScript with Zod schemas for parameter validation

### Microsoft Graph API Integration

The server communicates with Microsoft Graph API v1.0:

- Base URL: `https://graph.microsoft.com/v1.0`
- Three-level hierarchy: Lists → Tasks → Checklist Items
- Supports OData query parameters for filtering and sorting

### Environment Configuration

- `MSTODO_TOKEN_FILE`: Custom path for the encrypted token cache (defaults to the per-user app-data location; see above)
- `MSTODO_ACCESS_MODE`: `read` | `write` | `full`. Sets how much of the tool surface to expose (defaults to `full`; invalid values exit non-zero)
- `MSTODO_TIMEZONE`: IANA zone used by `get-agenda` to decide which day is "today" (defaults to the server's local zone; the tool's `timeZone` param overrides it)
- `.env` file required for authentication with CLIENT_ID, TENANT_ID (no CLIENT_SECRET; public client)

Cloud connector only (see `docs/CLOUD_CONNECTOR.md`):

- `MSTODO_BEARER_TOKEN`: shared secret `http-server.ts` requires on `/mcp` (fatal if unset)
- `MSTODO_ENCRYPTION_KEY`: fixed key material replacing the machine ID. **Never set it in a local `.env`**. The local cache would switch keys and `seed-cloud-token` (which always decrypts with the machine ID) would fail
- `MSTODO_TOKEN_STORE=key-vault` + `MSTODO_KEYVAULT_URL` + `MSTODO_TOKEN_SECRET_NAME`: keep the cache in Key Vault (the managed identity needs read **and** write, since refreshes are written back)
- `MSTODO_AUTH_TOKEN`: client side, read by `local-proxy.js`; same value as `MSTODO_BEARER_TOKEN`
- `PORT`: HTTP listen port (Container Apps injects it; defaults to 8080)

## Important Notes

- Always run `pnpm run build` (or `pnpm run validate`) after modifying TypeScript files (ts-builds/tsdown bundling)
- `pnpm run auth` is a one-shot interactive sign-in (opens a browser, exits when done). It is not a long-running server and does not bind any port itself; MSAL's internal loopback callback server binds to `127.0.0.1` only and exits after receiving the redirect
- Tokens are refreshed automatically and silently via MSAL's cache; app code never sees a raw refresh token
- Personal Microsoft accounts have limited API access compared to work/school accounts
