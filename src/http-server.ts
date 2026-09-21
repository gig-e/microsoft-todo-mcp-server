#!/usr/bin/env node

// Remote entry point for running this server as an Azure Container Apps connector, so
// the same account's To Do data is reachable from any machine without a local install.
// src/cli.ts (stdio) is unaffected — this is a second, opt-in transport over the same
// tool set.
import "./load-env.js"

import { timingSafeEqual } from "node:crypto"
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

async function main() {
  console.error(`Access mode: ${describeAccessMode(accessMode)}`)
  if (withheldTools.length > 0) {
    console.error(`Withheld ${withheldTools.length} tool(s): ${withheldTools.join(", ")}`)
  }

  await isPersonalMicrosoftAccount()

  // Stateless mode: Container Apps can run multiple replicas with no sticky sessions, so
  // per-request statelessness avoids a class of multi-instance bugs. Deploy with
  // --max-replicas 1 anyway, since this is a single-user tool.
  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined })
  await server.connect(transport)

  const httpServer = createHttpServer((req: IncomingMessage, res: ServerResponse) => {
    if (req.method !== "POST" || req.url !== "/mcp") {
      res.writeHead(404).end()
      return
    }

    if (!isAuthorized(req.headers.authorization)) {
      res.writeHead(401, { "content-type": "application/json" }).end(JSON.stringify({ error: "Unauthorized" }))
      return
    }

    transport.handleRequest(req, res).catch((error) => {
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
