// src/load-env.ts
//
// MCP clients (Claude Desktop, Cursor, MCP Inspector, etc.) spawn this server with an
// arbitrary working directory — dotenv's default `.env` lookup (relative to
// process.cwd()) can't be relied on. Resolve `.env` against locations we actually know
// instead, so CLIENT_ID/TENANT_ID load regardless of who launched us.
import dotenv from "dotenv"
import { dirname, join } from "path"
import { fileURLToPath } from "url"

import { getConfigDirPath } from "./msal-client.js"

/** The installed package root — one level up from the built file's own directory. */
export const packageRoot = join(dirname(fileURLToPath(import.meta.url)), "..")

/** `.env` beside the source, i.e. the repo root in a git checkout. */
export const packageEnvPath = join(packageRoot, ".env")

/**
 * `.env` in the per-user config directory that already holds the token cache. This is the
 * durable home for a global install, where the package root lives under npm's global
 * node_modules and is wiped on every upgrade.
 */
export const userEnvPath = join(getConfigDirPath(), ".env")

// Package root wins: dotenv never overwrites an already-set variable, and a checkout's
// own .env should beat a stale user-level one. Real environment variables (e.g. the MCP
// config's `env` block) still take precedence over both.
dotenv.config({ path: packageEnvPath })
dotenv.config({ path: userEnvPath })
