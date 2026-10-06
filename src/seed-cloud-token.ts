#!/usr/bin/env node

// One-time local step to move an already-signed-in token cache into the cloud deployment.
// Deliberately has no Azure SDK dependency — it's a pure crypto transform (decrypt with
// this machine's key, re-encrypt with the server's key); uploading the result to Key Vault
// happens via a plain `az keyvault secret set` command, not from here.
import "./load-env.js"

import { writeFileSync } from "node:fs"

import machineIdModule from "node-machine-id"

import { encrypt, readEncryptedFile } from "./crypto-store.js"
import { getCacheFilePath } from "./msal-client.js"

const { machineIdSync } = machineIdModule

const OUTPUT_FILE = ".cloud-token-cache.bin"

function main() {
  const localCachePath = getCacheFilePath()
  const plaintext = readEncryptedFile(localCachePath, machineIdSync())

  if (!plaintext) {
    console.error(`No local token cache found at ${localCachePath}. Run 'pnpm run auth' first.`)
    process.exit(1)
  }

  const serverKey = process.env.MSTODO_ENCRYPTION_KEY
  if (!serverKey) {
    console.error("MSTODO_ENCRYPTION_KEY is required — this is the key the cloud deployment will decrypt with.")
    process.exit(1)
  }

  // Base64 text, not raw bytes: Key Vault secrets are UTF-8 strings, and KeyVaultTokenStore
  // (src/token-store.ts) expects to base64-decode what it reads back.
  const ciphertext = encrypt(plaintext, serverKey).toString("base64")
  writeFileSync(OUTPUT_FILE, ciphertext, "utf8")

  console.error(`Wrote ${OUTPUT_FILE}. Upload it with:`)
  console.error(`  az keyvault secret set --vault-name <vault> --name <secret-name> --file ${OUTPUT_FILE}`)
}

main()
