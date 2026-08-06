#!/usr/bin/env node

import { spawn } from "child_process"
import { existsSync, readFileSync, writeFileSync } from "fs"
import { homedir } from "os"
import { join } from "path"
import readline from "readline"

import { type AccessMode, DEFAULT_ACCESS_MODE, describeAccessMode, parseAccessMode } from "./access-mode.js"
import { getCacheFilePath } from "./msal-client.js"

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
})

const question = (query: string): Promise<string> => {
  return new Promise((resolve) => rl.question(query, resolve))
}

const ACCESS_MODE_CHOICES: Record<string, AccessMode> = { "1": "read", "2": "write", "3": "full" }

async function askAccessMode(): Promise<AccessMode> {
  console.log("\n🔒 How much control should the assistant have?")
  console.log("  1) read  - read-only; browse lists and tasks, change nothing")
  console.log("  2) write - read plus create and update, but never delete")
  console.log("  3) full  - everything, including delete and archive")

  for (let attempt = 0; attempt < 3; attempt++) {
    const answer = (await question("Choose 1-3 (press Enter for full): ")).trim().toLowerCase()
    if (answer === "") return DEFAULT_ACCESS_MODE
    if (ACCESS_MODE_CHOICES[answer]) return ACCESS_MODE_CHOICES[answer]

    try {
      // Also accept the mode spelled out, e.g. "read-only".
      return parseAccessMode(answer)
    } catch {
      console.log("Please enter 1, 2 or 3.")
    }
  }

  console.log(`Falling back to '${DEFAULT_ACCESS_MODE}'. You can change MSTODO_ACCESS_MODE in the config later.`)
  return DEFAULT_ACCESS_MODE
}

async function setup() {
  console.log("🚀 Microsoft To Do MCP Server Setup")
  console.log("==================================\n")

  // Check if already configured
  const tokenPath = getCacheFilePath()

  if (existsSync(tokenPath)) {
    const answer = await question("Existing session found. Reconfigure? (y/N): ")
    if (answer.toLowerCase() !== "y") {
      console.log("Setup cancelled.")
      process.exit(0)
    }
  }

  // Check for Azure app credentials
  const hasEnvFile = existsSync(".env")

  if (!hasEnvFile) {
    console.log("\n📋 Azure App Registration Required")
    console.log("You need to create an app registration in Azure Portal first.")
    console.log("\nSteps:")
    console.log("1. Go to https://portal.azure.com")
    console.log("2. Navigate to 'App registrations' and create a new registration")
    console.log("3. Under 'Authentication', add a platform: 'Mobile and desktop applications'")
    console.log("4. Add redirect URI: http://localhost")
    console.log("5. Add these API permissions: Tasks.Read, Tasks.ReadWrite, User.Read")
    console.log("6. Do NOT create a client secret — this app authenticates as a public client (PKCE)\n")

    const clientId = await question("Enter your CLIENT_ID: ")
    const tenantId = (await question("Enter your TENANT_ID (press Enter for 'organizations'): ")) || "organizations"

    // Create .env file
    const envContent = `CLIENT_ID=${clientId}
TENANT_ID=${tenantId}
`
    writeFileSync(".env", envContent)
    console.log("✅ Created .env file")
  }

  // Asked before the browser opens so every prompt happens up front.
  const accessMode = await askAccessMode()

  console.log("\n🔐 Starting authentication flow...")
  console.log("A browser window will open. Please sign in with your Microsoft account.\n")

  // Run the interactive sign-in; it writes directly to the encrypted per-machine token store
  const authProcess = spawn("node", ["dist/auth-server.js"], {
    stdio: "inherit",
    shell: true,
  })

  authProcess.on("close", async (code) => {
    if (code === 0) {
      console.log("\n✅ Authentication successful!")
      console.log(`📁 Session stored securely at: ${tokenPath}`)

      await updateClaudeConfig(accessMode)

      console.log("\n🎉 Setup complete! Microsoft To Do MCP is ready to use.")
      console.log("Restart Claude Desktop to activate the integration.")
    } else {
      console.error("\n❌ Authentication failed. Please try again.")
    }

    rl.close()
  })
}

async function updateClaudeConfig(accessMode: AccessMode) {
  const serverEntry = {
    command: "npx",
    args: ["microsoft-todo-mcp-server"],
    // No tokens here — the server reads its own encrypted, per-machine store.
    env: { MSTODO_ACCESS_MODE: accessMode },
  }

  const claudeConfigPath =
    process.platform === "win32"
      ? join(process.env.APPDATA || "", "Claude", "claude_desktop_config.json")
      : process.platform === "darwin"
        ? join(homedir(), "Library", "Application Support", "Claude", "claude_desktop_config.json")
        : join(homedir(), ".config", "Claude", "claude_desktop_config.json")

  if (!existsSync(claudeConfigPath)) {
    console.log("\n⚠️  Claude config not found. Add this to your Claude desktop config manually:")
    console.log(JSON.stringify({ "microsoft-todo": serverEntry }, null, 2))
    return
  }

  try {
    const config = JSON.parse(readFileSync(claudeConfigPath, "utf8"))

    // Add or update the microsoft-todo server config
    if (!config.mcpServers) {
      config.mcpServers = {}
    }

    config.mcpServers["microsoft-todo"] = serverEntry

    writeFileSync(claudeConfigPath, JSON.stringify(config, null, 2))
    console.log("\n✅ Updated Claude Desktop configuration")
    console.log(`   Access mode: ${describeAccessMode(accessMode)}`)
  } catch (error) {
    console.error("\n⚠️  Could not update Claude config automatically:", error)
  }
}

// Run setup
setup().catch(console.error)
