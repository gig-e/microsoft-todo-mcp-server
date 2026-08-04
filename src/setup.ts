#!/usr/bin/env node

import { spawn } from "child_process"
import { existsSync, readFileSync, writeFileSync } from "fs"
import { homedir } from "os"
import { join } from "path"
import readline from "readline"

import { getCacheFilePath } from "./msal-client.js"

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
})

const question = (query: string): Promise<string> => {
  return new Promise((resolve) => rl.question(query, resolve))
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

      await updateClaudeConfig()

      console.log("\n🎉 Setup complete! Microsoft To Do MCP is ready to use.")
      console.log("Restart Claude Desktop to activate the integration.")
    } else {
      console.error("\n❌ Authentication failed. Please try again.")
    }

    rl.close()
  })
}

async function updateClaudeConfig() {
  const claudeConfigPath =
    process.platform === "win32"
      ? join(process.env.APPDATA || "", "Claude", "claude_desktop_config.json")
      : process.platform === "darwin"
        ? join(homedir(), "Library", "Application Support", "Claude", "claude_desktop_config.json")
        : join(homedir(), ".config", "Claude", "claude_desktop_config.json")

  if (!existsSync(claudeConfigPath)) {
    console.log("\n⚠️  Claude config not found. Add this to your Claude desktop config manually:")
    console.log(
      JSON.stringify(
        {
          "microsoft-todo": {
            command: "npx",
            args: ["microsoft-todo-mcp-server"],
            env: {},
          },
        },
        null,
        2,
      ),
    )
    return
  }

  try {
    const config = JSON.parse(readFileSync(claudeConfigPath, "utf8"))

    // Add or update the microsoft-todo server config
    if (!config.mcpServers) {
      config.mcpServers = {}
    }

    config.mcpServers["microsoft-todo"] = {
      command: "npx",
      args: ["microsoft-todo-mcp-server"],
      env: {}, // No need for tokens in env anymore!
    }

    writeFileSync(claudeConfigPath, JSON.stringify(config, null, 2))
    console.log("\n✅ Updated Claude Desktop configuration")
  } catch (error) {
    console.error("\n⚠️  Could not update Claude config automatically:", error)
  }
}

// Run setup
setup().catch(console.error)
