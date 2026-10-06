#!/usr/bin/env node

// Remote entry point for running this server as an Azure Container Apps connector, so
// the same account's To Do data is reachable from any machine without a local install.
// src/cli.ts (stdio) is unaffected — this is a second, opt-in transport over the same
// tool set.
import "./load-env.js"

import { randomUUID, timingSafeEqual } from "node:crypto"
import { createServer as createHttpServer, type IncomingMessage, type ServerResponse } from "node:http"

import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js"

import { describeAccessMode } from "./access-mode.js"
import { accessMode, isPersonalMicrosoftAccount, server, withheldTools } from "./todo-index.js"

const PORT = Number(process.env.PORT) || 8080
const BEARER_TOKEN = process.env.MSTODO_BEARER_TOKEN

if (!BEARER_TOKEN) {
  console.error("MSTODO_BEARER_TOKEN is required to run the HTTP server.")
  process.exit(1)
}

// Constant-time comparison: this header is the only thing standing between a public
// *.azurecontainerapps.io URL and full Tasks.ReadWrite access, so a timing side-channel
// on the check would defeat the point of it.
function isAuthorized(header: string | undefined): boolean {
  if (!header) return false
  const expected = Buffer.from(`Bearer ${BEARER_TOKEN}`)
  const actual = Buffer.from(header)
  return expected.length === actual.length && timingSafeEqual(expected, actual)
}

// Stateful mode (a session ID per client), not stateless: the SDK's stateless transport
// throws "Stateless transport cannot be reused across requests" on the second call to
// handleRequest — it's meant to be constructed fresh per request. Stateful mode instead
// requires one transport per session, and the MCP SDK's Server can only ever be connected
// to one transport at a time ("Already connected to a transport. Call close() before
// connecting to a new transport, or use a separate Protocol instance per connection.") —
// so a second client's `initialize` fails outright unless the previous session is closed
// first. This is a single-user tool (not a multi-tenant service), so rather than
// refactoring todo-index.ts into a per-session server factory, the simpler fit is: track
// the one active session, and close-then-reconnect the same shared `server` when a new
// session starts. That does mean a second concurrent client would evict the first, which
// is an acceptable trade for this use case.
let activeTransport: StreamableHTTPServerTransport | undefined
let activeSessionId: string | undefined

async function getTransportForRequest(req: IncomingMessage): Promise<StreamableHTTPServerTransport> {
  const rawSessionId = req.headers["mcp-session-id"]
  const sessionIdHeader = Array.isArray(rawSessionId) ? rawSessionId[0] : rawSessionId

  if (activeTransport && sessionIdHeader !== undefined && sessionIdHeader === activeSessionId) {
    return activeTransport
  }

  if (activeTransport) {
    await server.close()
    activeTransport = undefined
    activeSessionId = undefined
  }

  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: randomUUID,
    onsessioninitialized: (newSessionId) => {
      activeSessionId = newSessionId
    },
  })
  transport.onclose = () => {
    activeTransport = undefined
    activeSessionId = undefined
  }
  await server.connect(transport)
  activeTransport = transport
  return transport
}

async function main() {
  console.error(`Access mode: ${describeAccessMode(accessMode)}`)
  if (withheldTools.length > 0) {
    console.error(`Withheld ${withheldTools.length} tool(s): ${withheldTools.join(", ")}`)
  }

  await isPersonalMicrosoftAccount()

  const httpServer = createHttpServer((req: IncomingMessage, res: ServerResponse) => {
    // Dedicated, unauthenticated health endpoint for Container Apps' startup/liveness
    // probes (configured explicitly to hit this path — see the deploy notes). Azure's
    // default probe otherwise hits "/" (which we'd 404) or would need to carry the bearer
    // token to get anything but 401 from "/mcp", so the platform would never consider a
    // replica healthy even though the app is fine — this is a documented Container-Apps-
    // plus-MCP-server gotcha, not something specific to this server's auth design.
    if (req.url === "/health") {
      res.writeHead(200, { "content-type": "text/plain" }).end("ok")
      return
    }

    // GET is allowed alongside POST: the transport's standalone SSE stream, and some MCP
    // clients (e.g. mcp-remote) probe/fall back to a GET-based transport strategy when an
    // OAuth discovery request 404s, which it does here since we don't implement OAuth.
    if (req.url !== "/mcp") {
      res.writeHead(404).end()
      return
    }

    if (!isAuthorized(req.headers.authorization)) {
      res.writeHead(401, { "content-type": "application/json" }).end(JSON.stringify({ error: "Unauthorized" }))
      return
    }

    getTransportForRequest(req)
      .then((transport) => transport.handleRequest(req, res))
      .catch((error) => {
        console.error("Error handling request:", error)
        if (!res.headersSent) {
          res.writeHead(500).end()
        }
      })
  })

  httpServer.listen(PORT, () => {
    console.error(`Microsoft To Do MCP (HTTP) listening on port ${PORT}`)
  })
}

main().catch((error) => {
  console.error("Error starting HTTP server:", error)
  process.exit(1)
})
