// src/load-env.ts
//
// MCP clients (Claude Desktop, Cursor, MCP Inspector, etc.) spawn this server with an
// arbitrary working directory — dotenv's default `.env` lookup (relative to
// process.cwd()) can't be relied on. Load .env relative to where this module actually
// lives instead, so CLIENT_ID/TENANT_ID resolve regardless of who launched us.
import dotenv from "dotenv"
import { dirname, join } from "path"
import { fileURLToPath } from "url"

const packageRoot = join(dirname(fileURLToPath(import.meta.url)), "..")

dotenv.config({ path: join(packageRoot, ".env") })
