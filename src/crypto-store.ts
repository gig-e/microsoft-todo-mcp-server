// src/crypto-store.ts
import { createCipheriv, createDecipheriv, randomBytes, scryptSync } from "crypto"
import { existsSync, readFileSync, writeFileSync } from "fs"
import machineIdModule from "node-machine-id"

const { machineIdSync } = machineIdModule

// Bumping this salt invalidates every previously-encrypted cache file on next auth.
const KEY_SALT = "microsoft-todo-mcp-server:token-store:v1"
const ALGORITHM = "aes-256-gcm"
const IV_LENGTH = 12
const AUTH_TAG_LENGTH = 16

// Deriving the key from the machine id (rather than a stored/random key) means the
// ciphertext only decrypts on the machine that wrote it — copying the file elsewhere
// (e.g. via OneDrive sync) yields unusable bytes instead of portable credentials.
function deriveKey(): Buffer {
  const machineId = machineIdSync()
  return scryptSync(machineId, KEY_SALT, 32)
}

export function encrypt(plaintext: string): Buffer {
  const key = deriveKey()
  const iv = randomBytes(IV_LENGTH)
  const cipher = createCipheriv(ALGORITHM, key, iv)
  const ciphertext = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()])
  const authTag = cipher.getAuthTag()
  return Buffer.concat([iv, authTag, ciphertext])
}

export function decrypt(data: Buffer): string {
  const key = deriveKey()
  const iv = data.subarray(0, IV_LENGTH)
  const authTag = data.subarray(IV_LENGTH, IV_LENGTH + AUTH_TAG_LENGTH)
  const ciphertext = data.subarray(IV_LENGTH + AUTH_TAG_LENGTH)
  const decipher = createDecipheriv(ALGORITHM, key, iv)
  decipher.setAuthTag(authTag)
  return Buffer.concat([decipher.update(ciphertext), decipher.final()]).toString("utf8")
}

export function readEncryptedFile(path: string): string | null {
  if (!existsSync(path)) {
    return null
  }

  try {
    return decrypt(readFileSync(path))
  } catch (error) {
    console.error(`Could not decrypt token store at ${path}; treating as absent.`, error)
    return null
  }
}

export function writeEncryptedFile(path: string, content: string): void {
  writeFileSync(path, encrypt(content))
}
