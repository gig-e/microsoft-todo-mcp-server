#!/usr/bin/env node

import fs from "fs"
import path from "path"

import { type AccessMode, DEFAULT_ACCESS_MODE, describeAccessMode, parseAccessMode } from "./access-mode.js"

// No tokens are embedded in this file: the server reads its own encrypted, per-machine
// token store (see src/msal-client.ts) and needs no secrets passed via env.
const outputPath = process.argv[2] || path.join(process.cwd(), "mcp.json")

// The access mode is always written out, even at its default, so it's visible and
// editable in the generated config rather than being an env var users must discover.
let accessMode: AccessMode
try {
  accessMode = parseAccessMode(process.argv[3] ?? DEFAULT_ACCESS_MODE)
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error))
  console.error("Usage: mstodo-config [output-path] [read|write|full]")
  process.exit(1)
}

const mcpConfig = {
  mcpServers: {
    microsoftTodo: {
      command: "npx",
      args: ["--yes", "microsoft-todo-mcp-server"],
      env: {
        MSTODO_ACCESS_MODE: accessMode,
      },
    },
  },
}

try {
  fs.writeFileSync(outputPath, JSON.stringify(mcpConfig, null, 2), "utf8")
  console.log(`MCP configuration written to: ${outputPath}`)
  console.log(`Access mode: ${describeAccessMode(accessMode)}`)
  console.log("Run 'pnpm run auth' once per machine to authenticate — no tokens are stored in this file.")
} catch (error) {
  const errorMessage = error instanceof Error ? error.message : String(error)
  console.error("Error creating MCP config:", errorMessage)
  process.exit(1)
}
