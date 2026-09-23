// src/token-store.ts
//
// Where the encrypted MSAL cache lives. Local mode (default) writes it to a file on this
// machine, unchanged from before. Server mode (MSTODO_TOKEN_STORE=key-vault) reads/writes
// it as a Key Vault secret instead, since Container Apps' local disk isn't durable across
// replica restarts/redeploys.
import { DefaultAzureCredential } from "@azure/identity"
import { SecretClient } from "@azure/keyvault-secrets"

import { decrypt, encrypt, readEncryptedFile, resolveKeyMaterial, writeEncryptedFile } from "./crypto-store.js"

export interface TokenStore {
  read(): Promise<string | null>
  write(content: string): Promise<void>
}

export class LocalFileTokenStore implements TokenStore {
  constructor(private readonly filePath: string) {}

  async read(): Promise<string | null> {
    return readEncryptedFile(this.filePath, resolveKeyMaterial())
  }

  async write(content: string): Promise<void> {
    writeEncryptedFile(this.filePath, content, resolveKeyMaterial())
  }
}

function isSecretNotFound(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    "statusCode" in error &&
    (error as { statusCode: unknown }).statusCode === 404
  )
}

export class KeyVaultTokenStore implements TokenStore {
  private readonly client: SecretClient

  constructor(
    vaultUrl: string,
    private readonly secretName: string,
  ) {
    this.client = new SecretClient(vaultUrl, new DefaultAzureCredential())
  }

  async read(): Promise<string | null> {
    try {
      const secret = await this.client.getSecret(this.secretName)
      if (!secret.value) return null
      return decrypt(Buffer.from(secret.value, "base64"), resolveKeyMaterial())
    } catch (error) {
      if (isSecretNotFound(error)) return null
      throw error
    }
  }

  async write(content: string): Promise<void> {
    const ciphertext = encrypt(content, resolveKeyMaterial())
    await this.client.setSecret(this.secretName, ciphertext.toString("base64"))
  }
}

/** localFilePath is passed in rather than resolved here to avoid a circular import with msal-client.ts. */
export function resolveTokenStore(localFilePath: string): TokenStore {
  if (process.env.MSTODO_TOKEN_STORE === "key-vault") {
    const vaultUrl = process.env.MSTODO_KEYVAULT_URL
    const secretName = process.env.MSTODO_TOKEN_SECRET_NAME
    if (!vaultUrl || !secretName) {
      throw new Error("MSTODO_TOKEN_STORE=key-vault requires MSTODO_KEYVAULT_URL and MSTODO_TOKEN_SECRET_NAME.")
    }
    return new KeyVaultTokenStore(vaultUrl, secretName)
  }

  return new LocalFileTokenStore(localFilePath)
}
