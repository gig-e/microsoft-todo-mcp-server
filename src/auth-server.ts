#!/usr/bin/env node

// One-shot interactive sign-in for Microsoft To Do MCP.
//
// Uses MSAL Node's PublicClientApplication.acquireTokenInteractive(), which handles
// PKCE (code_verifier/code_challenge), the OAuth `state`/nonce values, and the loopback
// callback server (bound to 127.0.0.1 only) internally. There is no client secret and
// no hand-rolled Express server here on purpose.
import "./load-env.js"

import open from "open"

import { createPublicClientApplication, getCacheFilePath, scopes } from "./msal-client.js"

console.error("Microsoft To Do MCP — authentication")
console.error("CLIENT_ID:", process.env.CLIENT_ID ? "Present" : "Missing")
console.error("TENANT_ID:", process.env.TENANT_ID || 'Not specified, using "organizations" (multi-tenant)')

async function main() {
  const pca = createPublicClientApplication()

  console.error("Opening your browser to sign in with Microsoft...")

  const result = await pca.acquireTokenInteractive({
    scopes,
    prompt: "consent",
    openBrowser: async (url: string) => {
      await open(url)
    },
    successTemplate: "<h1>Authentication successful</h1><p>You can close this window and return to the terminal.</p>",
    errorTemplate: "<h1>Authentication failed</h1><p>Please check the terminal for details and try again.</p>",
  })

  if (!result) {
    throw new Error("No result returned from interactive authentication")
  }

  console.error(`Authentication successful for ${result.account?.username ?? "your account"}.`)
  console.error(`Session stored securely at: ${getCacheFilePath()}`)
}

main().catch((error) => {
  console.error("Authentication failed:", error instanceof Error ? error.message : error)
  process.exit(1)
})
