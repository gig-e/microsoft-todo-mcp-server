#!/usr/bin/env node

import { startServer } from "./todo-index.js"

console.error("Microsoft Todo MCP CLI")

startServer().catch((error) => {
  const errorMessage = error instanceof Error ? error.message : String(error)
  console.error("Error starting server:", errorMessage)
  process.exit(1)
})
