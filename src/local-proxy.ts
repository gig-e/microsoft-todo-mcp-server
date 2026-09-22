#!/usr/bin/env node

// Minimal stdio <-> Streamable HTTP proxy for Claude Desktop's remote connector config.
// Replaces the third-party mcp-remote package, which hit a Windows-specific bug when
// Claude Desktop spawned it via cmd.exe (`'C:\Program' is not recognized as an internal
// or external command`, immediately after the client's `initialize` message — every time,
// regardless of --transport flags). Spawned directly with `node` (no npx/cmd.exe in the
// path), with no OAuth probing or SSE-fallback complexity: just a static bearer token and
// a dumb relay that tracks this server's Mcp-Session-Id and unwraps its SSE-framed
// single responses back into plain JSON lines for the stdio side.
import { createInterface } from "node:readline"

const REMOTE_URL = process.argv[2]
const BEARER_TOKEN = process.env.MSTODO_AUTH_TOKEN

if (!REMOTE_URL) {
  console.error("Usage: local-proxy.js <remote-url>")
  process.exit(1)
}
if (!BEARER_TOKEN) {
  console.error("MSTODO_AUTH_TOKEN is required.")
  process.exit(1)
}

let sessionId: string | undefined

async function forward(line: string): Promise<void> {
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    Accept: "application/json, text/event-stream",
    Authorization: `Bearer ${BEARER_TOKEN}`,
  }
  if (sessionId) headers["Mcp-Session-Id"] = sessionId

  const response = await fetch(REMOTE_URL, { method: "POST", headers, body: line })

  const newSessionId = response.headers.get("mcp-session-id")
  if (newSessionId) sessionId = newSessionId

  // Notifications get a 202 with no body — nothing to relay back to the client.
  if (response.status === 202) {
    return
  }

  const text = await response.text()
  if (!response.ok) {
    console.error(`Remote server returned ${response.status}: ${text}`)
    return
  }

  const contentType = response.headers.get("content-type") ?? ""
  if (contentType.includes("text/event-stream")) {
    for (const rawEvent of text.split("\n\n")) {
      const dataLine = rawEvent.split("\n").find((eventLine) => eventLine.startsWith("data:"))
      if (dataLine) {
        process.stdout.write(dataLine.slice("data:".length).trim() + "\n")
      }
    }
    return
  }

  if (text.trim()) {
    process.stdout.write(text.trim() + "\n")
  }
}

const rl = createInterface({ input: process.stdin })
rl.on("line", (line) => {
  if (!line.trim()) return
  forward(line).catch((error) => {
    console.error("Proxy error:", error instanceof Error ? error.message : error)
  })
})
