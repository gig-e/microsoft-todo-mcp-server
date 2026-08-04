// src/msal-client.ts
import { ICachePlugin, LogLevel, PublicClientApplication } from "@azure/msal-node"
import { existsSync, mkdirSync } from "fs"
import { homedir } from "os"
import { join } from "path"

import { readEncryptedFile, writeEncryptedFile } from "./crypto-store.js"

export const scopes = [
  "offline_access", // keep first so it isn't dropped from consent
  "openid",
  "profile",
  "Tasks.Read",
  "Tasks.Read.Shared",
  "Tasks.ReadWrite",
  "Tasks.ReadWrite.Shared",
  "User.Read",
]

export function getConfigDir(): string {
  const configDir =
    process.platform === "win32"
      ? join(process.env.APPDATA || join(homedir(), "AppData", "Roaming"), "microsoft-todo-mcp")
      : join(homedir(), ".config", "microsoft-todo-mcp")

  if (!existsSync(configDir)) {
    mkdirSync(configDir, { recursive: true })
  }

  return configDir
}

export function getCacheFilePath(): string {
  return process.env.MSTODO_TOKEN_FILE || join(getConfigDir(), "token-cache.bin")
}

function createCachePlugin(cacheFilePath: string): ICachePlugin {
  return {
    beforeCacheAccess: async (cacheContext) => {
      const data = readEncryptedFile(cacheFilePath)
      if (data) {
        cacheContext.tokenCache.deserialize(data)
      }
    },
    afterCacheAccess: async (cacheContext) => {
      if (cacheContext.cacheHasChanged) {
        writeEncryptedFile(cacheFilePath, cacheContext.tokenCache.serialize())
      }
    },
  }
}

export function getTenantId(): string {
  return process.env.TENANT_ID || "organizations"
}

export function createPublicClientApplication(): PublicClientApplication {
  const clientId = process.env.CLIENT_ID
  if (!clientId) {
    throw new Error("CLIENT_ID is required. Set it in your .env file.")
  }

  return new PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${getTenantId()}`,
    },
    cache: {
      cachePlugin: createCachePlugin(getCacheFilePath()),
    },
    system: {
      loggerOptions: {
        loggerCallback: (_level, message) => {
          console.error(`MSAL: ${message}`)
        },
        // Must stay false: PII logging can include tokens and account claims.
        piiLoggingEnabled: false,
        logLevel: LogLevel.Warning,
      },
    },
  })
}
