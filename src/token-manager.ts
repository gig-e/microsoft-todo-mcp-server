// src/token-manager.ts
import { PublicClientApplication } from "@azure/msal-node"

import { createPublicClientApplication, scopes } from "./msal-client.js"

interface TokenData {
  accessToken: string
  expiresAt: number
}

export class TokenManager {
  private pca: PublicClientApplication

  constructor() {
    this.pca = createPublicClientApplication()
  }

  // Acquires a token silently against MSAL's own encrypted cache. As a public client,
  // MSAL refreshes using its internally-cached refresh token — we never see or store
  // a raw refresh token ourselves.
  async getTokens(): Promise<TokenData | null> {
    const accounts = await this.pca.getTokenCache().getAllAccounts()
    const account = accounts[0]

    if (!account) {
      return null
    }

    try {
      const result = await this.pca.acquireTokenSilent({ account, scopes })

      if (!result) {
        return null
      }

      return {
        accessToken: result.accessToken,
        expiresAt: result.expiresOn ? result.expiresOn.getTime() : Date.now() + 3600 * 1000,
      }
    } catch (error) {
      console.error("Silent token acquisition failed:", error instanceof Error ? error.message : error)
      this.promptForReauth()
      return null
    }
  }

  promptForReauth(): void {
    console.error(`
=================================================================
RE-AUTHENTICATION REQUIRED

Your Microsoft To Do session has expired or could not be refreshed silently.

To fix this:
1. Open a new terminal
2. Navigate to the microsoft-todo-mcp-server directory
3. Run: pnpm run auth
4. Complete the authentication in your browser
5. Restart Claude Desktop to use the new session
=================================================================
    `)
  }
}

export const tokenManager = new TokenManager()
