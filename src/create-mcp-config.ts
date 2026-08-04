#!/usr/bin/env node

import fs from "fs"
import path from "path"

// No tokens are embedded in this file: the server reads its own encrypted, per-machine
// token store (see src/msal-client.ts) and needs no secrets passed via env.
const outputPath = process.argv[2] || path.join(process.cwd(), "mcp.json")

const mcpConfig = {
  mcpServers: {
    microsoftTodo: {
      command: "npx",
      args: ["--yes", "microsoft-todo-mcp-server"],
      env: {},
    },
  },
}

try {
  fs.writeFileSync(outputPath, JSON.stringify(mcpConfig, null, 2), "utf8")
  console.log(`MCP configuration written to: ${outputPath}`)
  console.log("Run 'pnpm run auth' once per machine to authenticate — no tokens are stored in this file.")
} catch (error) {
  const errorMessage = error instanceof Error ? error.message : String(error)
  console.error("Error creating MCP config:", errorMessage)
  process.exit(1)
}
