import "./load-env.js"

import { McpServer, type ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import { z, type ZodRawShape } from "zod"

import { type AccessMode, describeAccessMode, isToolAllowed, parseAccessMode, type ToolAccess } from "./access-mode.js"
import {
  type AgendaBucket,
  type AgendaTask,
  bucketForTask,
  compareTasksByDue,
  dayKeyInZone,
  dueDayKey,
  isValidTimeZone,
  mapWithConcurrency,
  matchesQuery,
} from "./agenda.js"
import { tokenManager } from "./token-manager.js"

// Microsoft Graph API endpoints
const MS_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
const USER_AGENT = "microsoft-todo-mcp-server/1.0"

// Create server instance. Exported so src/http-server.ts can connect a second (HTTP)
// transport to the same instance without duplicating tool registration.
export const server = new McpServer({
  name: "mstodo",
  version: "1.0.0",
})

// A bad MSTODO_ACCESS_MODE is fatal by design — see access-mode.ts. Exit with just the
// message rather than a stack trace, since MCP clients surface stderr to the user raw.
export const accessMode: AccessMode = ((): AccessMode => {
  try {
    return parseAccessMode(process.env.MSTODO_ACCESS_MODE)
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exit(1)
  }
})()

export const withheldTools: string[] = []

/**
 * Register a tool only if the configured access mode permits it. Withheld tools never
 * reach tools/list, so an assistant in read mode can't call them and be refused — from
 * its point of view they don't exist.
 */
function registerTool<Args extends ZodRawShape>(
  access: ToolAccess,
  name: string,
  description: string,
  paramsSchema: Args,
  cb: ToolCallback<Args>,
): void {
  if (!isToolAllowed(access, accessMode)) {
    withheldTools.push(`${name} (${access})`)
    return
  }

  server.tool(name, description, paramsSchema, cb)
}

// Graph throttles bursts with 429 and sheds load with 503, asking callers to honour
// Retry-After. Querying every list at once (search-tasks, get-agenda) hits this routinely,
// and an unretried throttle looks like an empty list rather than an error.
const MAX_THROTTLE_RETRIES = 3
const MAX_RETRY_DELAY_MS = 10_000

function isThrottled(status: number): boolean {
  return status === 429 || status === 503
}

function retryDelayMs(response: Response, attempt: number): number {
  const retryAfter = Number(response.headers.get("Retry-After"))
  // Fall back to exponential backoff when the header is absent or non-numeric (it may
  // also be an HTTP date, which we don't try to parse).
  const suggested = Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter * 1000 : 2 ** attempt * 500
  return Math.min(suggested, MAX_RETRY_DELAY_MS)
}

const delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))

// Helper function for making Microsoft Graph API requests
async function makeGraphRequest<T>(url: string, token: string, method = "GET", body?: any): Promise<T | null> {
  const headers = {
    "User-Agent": USER_AGENT,
    Accept: "application/json",
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  }

  try {
    const options: RequestInit = {
      method,
      headers,
    }

    if (body && (method === "POST" || method === "PATCH")) {
      options.body = JSON.stringify(body)
    }

    console.error(`Making request to: ${url}`)
    console.error(
      `Request options: ${JSON.stringify({
        method,
        headers: {
          ...headers,
          Authorization: "Bearer [REDACTED]",
        },
      })}`,
    )

    let response = await fetch(url, options)

    // If we get a 401, try to refresh the token and retry once
    if (response.status === 401) {
      console.error("Got 401, attempting token refresh...")
      const newToken = await getAccessToken() // This will trigger refresh
      if (newToken && newToken !== token) {
        // Retry with new token
        headers.Authorization = `Bearer ${newToken}`
        response = await fetch(url, { ...options, headers })
      }
    }

    for (let attempt = 0; isThrottled(response.status) && attempt < MAX_THROTTLE_RETRIES; attempt++) {
      const wait = retryDelayMs(response, attempt)
      console.error(`Throttled with ${response.status}; retrying in ${wait}ms (attempt ${attempt + 1})`)
      await delay(wait)
      response = await fetch(url, { ...options, headers })
    }

    if (!response.ok) {
      const errorText = await response.text()
      console.error(`HTTP error! status: ${response.status}, body: ${errorText}`)

      // Check for the specific MailboxNotEnabledForRESTAPI error
      if (errorText.includes("MailboxNotEnabledForRESTAPI")) {
        console.error(`
=================================================================
ERROR: MailboxNotEnabledForRESTAPI

The Microsoft To Do API is not available for personal Microsoft accounts 
(outlook.com, hotmail.com, live.com, etc.) through the Graph API.

This is a limitation of the Microsoft Graph API, not an authentication issue.
Microsoft only allows To Do API access for Microsoft 365 business accounts.

You can still use Microsoft To Do through the web interface or mobile apps,
but API access is restricted for personal accounts.
=================================================================
        `)

        throw new Error(
          "Microsoft To Do API is not available for personal Microsoft accounts. See console for details.",
        )
      }

      throw new Error(`HTTP error! status: ${response.status}, body: ${errorText}`)
    }

    const data = await response.json()
    console.error(`Response received: ${JSON.stringify(data).substring(0, 200)}...`)
    return data as T
  } catch (error) {
    console.error("Error making Graph API request:", error)
    return null
  }
}

// Authentication helper using delegated flow with token manager
async function getAccessToken(): Promise<string | null> {
  try {
    console.error("getAccessToken called")

    // Use the token manager to get tokens (handles all sources and refresh)
    const tokens = await tokenManager.getTokens()

    if (tokens) {
      console.error(`Successfully retrieved valid token`)
      return tokens.accessToken
    }

    console.error("No valid tokens available")
    return null
  } catch (error) {
    console.error("Error getting access token:", error)
    return null
  }
}

// Function to check if the account is a personal Microsoft account
export async function isPersonalMicrosoftAccount(): Promise<boolean> {
  try {
    const token = await getAccessToken()
    if (!token) return false

    // Make a request to get user info
    const url = `${MS_GRAPH_BASE}/me`
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    })

    if (!response.ok) {
      console.error(`Error getting user info: ${response.status}`)
      return false
    }

    const userData = await response.json()
    const email = userData.mail || userData.userPrincipalName || ""

    // Check if the email domain indicates a personal account
    const personalDomains = ["outlook.com", "hotmail.com", "live.com", "msn.com", "passport.com"]
    const domain = email.split("@")[1]?.toLowerCase()

    if (domain && personalDomains.some((d) => domain.includes(d))) {
      console.error(`
=================================================================
WARNING: Personal Microsoft Account Detected

Your Microsoft account (${email}) appears to be a personal account.
Microsoft To Do API access is typically not available for personal accounts
through the Microsoft Graph API, only for Microsoft 365 business accounts.

You may encounter the "MailboxNotEnabledForRESTAPI" error when trying to
access To Do lists or tasks. This is a limitation of the Microsoft Graph API,
not an issue with your authentication or this application.

You can still use Microsoft To Do through the web interface or mobile apps,
but API access is restricted for personal accounts.
=================================================================
      `)
      return true
    }

    return false
  } catch (error) {
    console.error("Error checking account type:", error)
    return false
  }
}

// Server tool to check authentication status
registerTool(
  "read",
  "auth-status",
  "Check if you're authenticated with Microsoft Graph API. Shows current token status and expiration time, and indicates if the token needs to be refreshed.",
  {},
  async () => {
    const tokens = await tokenManager.getTokens()

    if (!tokens) {
      return {
        content: [
          {
            type: "text",
            text: "Not authenticated. Please run 'npx microsoft-todo-mcp-server setup' to authenticate with Microsoft.",
          },
        ],
      }
    }

    const isExpired = Date.now() > tokens.expiresAt
    const expiryTime = new Date(tokens.expiresAt).toLocaleString()

    // Surface the access mode here too: if a tool the user expects is missing from the
    // list, this is where they'll look to find out why.
    let modeMessage = `\n\nAccess mode: ${describeAccessMode(accessMode)}.`
    if (withheldTools.length > 0) {
      modeMessage += `\nWithheld by this mode (${withheldTools.length}): ${withheldTools.join(", ")}.`
      modeMessage += `\nSet MSTODO_ACCESS_MODE=full in the MCP server config to enable them.`
    }

    // Check if it's a personal account
    const isPersonal = await isPersonalMicrosoftAccount()
    let accountMessage = ""

    if (isPersonal) {
      accountMessage =
        "\n\n⚠️ WARNING: You are using a personal Microsoft account. " +
        "Microsoft To Do API access is typically not available for personal accounts " +
        "through the Microsoft Graph API. You may encounter 'MailboxNotEnabledForRESTAPI' errors. " +
        "This is a Microsoft limitation, not an authentication issue."
    }

    if (isExpired) {
      return {
        content: [
          {
            type: "text",
            text: `Authentication expired at ${expiryTime}. Will attempt to refresh when you call any API.${accountMessage}${modeMessage}`,
          },
        ],
      }
    } else {
      return {
        content: [
          {
            type: "text",
            text: `Authenticated. Token expires at ${expiryTime}.${accountMessage}${modeMessage}`,
          },
        ],
      }
    }
  },
)

interface TaskList {
  id: string
  displayName: string
  isOwner?: boolean
  isShared?: boolean
  wellknownListName?: string // 'none', 'defaultList', 'flaggedEmails', 'unknownFutureValue'
}

interface Task {
  id: string
  title: string
  status: string
  importance: string
  dueDateTime?: {
    dateTime: string
    timeZone: string
  }
  completedDateTime?: {
    dateTime: string
    timeZone: string
  }
  reminderDateTime?: {
    dateTime: string
    timeZone: string
  }
  body?: {
    content: string
    contentType: string
  }
  categories?: string[]
}

interface ChecklistItem {
  id: string
  displayName: string
  isChecked: boolean
  createdDateTime?: string
}

interface PlannerTask {
  id: string
  planId: string
  bucketId: string
  title: string
  percentComplete: number
  priority: number
  dueDateTime?: string
  startDateTime?: string
  completedDateTime?: string | null
  hasDescription: boolean
  checklistItemCount: number
  activeChecklistItemCount: number
}

interface PlannerPlan {
  id: string
  title: string
}

interface PlannerChecklistItem {
  title: string
  isChecked: boolean
}

interface PlannerTaskDetails {
  description?: string
  checklist?: Record<string, PlannerChecklistItem>
}

function formatPlannerPriority(priority: number): string {
  if (priority <= 1) return "Urgent"
  if (priority <= 4) return "Important"
  if (priority <= 6) return "Medium"
  return "Low"
}

function formatPlannerStatus(percentComplete: number): string {
  if (percentComplete >= 100) return "✓"
  if (percentComplete > 0) return "◐"
  return "○"
}

// Register tools
registerTool(
  "read",
  "get-task-lists",
  "Get all Microsoft Todo task lists (the top-level containers that organize your tasks). Shows list names, IDs, and indicates default or shared lists.",
  {},
  async () => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to retrieve task lists",
            },
          ],
        }
      }

      const lists = response.value || []
      if (lists.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No task lists found.",
            },
          ],
        }
      }

      const formattedLists = lists.map((list) => {
        // Add well-known list name if applicable
        let wellKnownInfo = ""
        if (list.wellknownListName && list.wellknownListName !== "none") {
          if (list.wellknownListName === "defaultList") {
            wellKnownInfo = " (Default Tasks List)"
          } else if (list.wellknownListName === "flaggedEmails") {
            wellKnownInfo = " (Flagged Emails)"
          }
        }

        // Add sharing info if applicable
        let sharingInfo = ""
        if (list.isShared) {
          sharingInfo = list.isOwner ? " (Shared by you)" : " (Shared with you)"
        }

        return `ID: ${list.id}\nName: ${list.displayName}${wellKnownInfo}${sharingInfo}\n---`
      })

      return {
        content: [
          {
            type: "text",
            text: `Your task lists:\n\n${formattedLists.join("\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching task lists: ${error}`,
          },
        ],
      }
    }
  },
)

// Enhanced organized view of task lists
registerTool(
  "read",
  "get-task-lists-organized",
  "Get all task lists organized into logical folders/categories based on naming patterns, emoji prefixes, and sharing status. Provides a hierarchical view similar to folder organization.",
  {
    includeIds: z.boolean().optional().describe("Include list IDs in output (default: false)"),
    groupBy: z
      .enum(["category", "shared", "type"])
      .optional()
      .describe("Grouping strategy - 'category' (default), 'shared', or 'type'"),
  },
  async ({ includeIds, groupBy }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to retrieve task lists",
            },
          ],
        }
      }

      const lists = response.value || []
      if (lists.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No task lists found.",
            },
          ],
        }
      }

      // Group by shared status
      if (groupBy === "shared") {
        const sharedLists = lists.filter((l) => l.isShared)
        const personalLists = lists.filter((l) => !l.isShared)

        let output = "📂 Microsoft To Do Lists - By Sharing Status\n"
        output += "=".repeat(50) + "\n\n"

        output += `👥 Shared Lists (${sharedLists.length})\n`
        sharedLists.forEach((list) => {
          const ownership = list.isOwner ? "Shared by you" : "Shared with you"
          output += `   ├─ ${list.displayName} [${ownership}]\n`
        })

        output += `\n🔒 Personal Lists (${personalLists.length})\n`
        personalLists.forEach((list) => {
          output += `   ├─ ${list.displayName}\n`
        })

        return { content: [{ type: "text", text: output }] }
      }

      // Helper function to organize lists
      const organizeLists = (lists: TaskList[]): { [category: string]: TaskList[] } => {
        const organized: { [category: string]: TaskList[] } = {}

        // Patterns for categorizing lists
        const patterns = {
          archived: /\(([^)]+)\s*-\s*Archived\)$/i,
          archive: /^📦\s*Archive/i,
          shopping: /^🛒/,
          property: /^🏡/,
          family: /^👪/,
          seasonal: /^(🎄|🎉)/,
          work: /^(Work|SBIR)/i,
          travel: /^(🚗|Rangeley)/i,
          reading: /^📰/,
        }

        lists.forEach((list) => {
          // Check archived pattern
          const archiveMatch = list.displayName.match(patterns.archived)
          if (archiveMatch) {
            const category = `📦 Archived - ${archiveMatch[1]}`
            if (!organized[category]) organized[category] = []
            organized[category].push(list)
          }
          // Check archive prefix
          else if (patterns.archive.test(list.displayName)) {
            if (!organized["📦 Archives"]) organized["📦 Archives"] = []
            organized["📦 Archives"].push(list)
          }
          // Check shopping lists
          else if (patterns.shopping.test(list.displayName)) {
            if (!organized["🛒 Shopping Lists"]) organized["🛒 Shopping Lists"] = []
            organized["🛒 Shopping Lists"].push(list)
          }
          // Check property lists
          else if (patterns.property.test(list.displayName)) {
            if (!organized["🏡 Properties"]) organized["🏡 Properties"] = []
            organized["🏡 Properties"].push(list)
          }
          // Check family lists
          else if (patterns.family.test(list.displayName)) {
            if (!organized["👪 Family"]) organized["👪 Family"] = []
            organized["👪 Family"].push(list)
          }
          // Check seasonal lists
          else if (patterns.seasonal.test(list.displayName)) {
            if (!organized["🎉 Seasonal & Events"]) organized["🎉 Seasonal & Events"] = []
            organized["🎉 Seasonal & Events"].push(list)
          }
          // Check work lists
          else if (patterns.work.test(list.displayName)) {
            if (!organized["💼 Work"]) organized["💼 Work"] = []
            organized["💼 Work"].push(list)
          }
          // Check travel lists
          else if (patterns.travel.test(list.displayName)) {
            if (!organized["🚗 Travel & Rangeley"]) organized["🚗 Travel & Rangeley"] = []
            organized["🚗 Travel & Rangeley"].push(list)
          }
          // Check reading lists
          else if (patterns.reading.test(list.displayName)) {
            if (!organized["📚 Reading"]) organized["📚 Reading"] = []
            organized["📚 Reading"].push(list)
          }
          // Special lists
          else if (list.wellknownListName && list.wellknownListName !== "none") {
            if (!organized["⭐ Special Lists"]) organized["⭐ Special Lists"] = []
            organized["⭐ Special Lists"].push(list)
          }
          // Shared lists
          else if (list.isShared) {
            if (!organized["👥 Shared Lists"]) organized["👥 Shared Lists"] = []
            organized["👥 Shared Lists"].push(list)
          }
          // Everything else
          else {
            if (!organized["📋 Other Lists"]) organized["📋 Other Lists"] = []
            organized["📋 Other Lists"].push(list)
          }
        })

        return organized
      }

      // Default: organize by category
      const organized = organizeLists(lists)

      let output = "📂 Microsoft To Do Lists - Organized View\n"
      output += "=".repeat(50) + "\n\n"

      // Sort categories for consistent display
      const sortedCategories = Object.keys(organized).sort((a, b) => {
        // Priority order for categories
        const priority: { [key: string]: number } = {
          "⭐ Special Lists": 1,
          "👥 Shared Lists": 2,
          "💼 Work": 3,
          "👪 Family": 4,
          "🏡 Properties": 5,
          "🛒 Shopping Lists": 6,
          "🚗 Travel & Rangeley": 7,
          "🎉 Seasonal & Events": 8,
          "📚 Reading": 9,
          "📋 Other Lists": 10,
          "📦 Archives": 11,
        }

        // Check if categories start with "📦 Archived -"
        const aIsArchived = a.startsWith("📦 Archived -")
        const bIsArchived = b.startsWith("📦 Archived -")

        if (aIsArchived && !bIsArchived) return 1
        if (!aIsArchived && bIsArchived) return -1
        if (aIsArchived && bIsArchived) return a.localeCompare(b)

        const aPriority = priority[a] || 999
        const bPriority = priority[b] || 999

        if (aPriority !== bPriority) return aPriority - bPriority
        return a.localeCompare(b)
      })

      sortedCategories.forEach((category) => {
        const categoryLists = organized[category]
        output += `${category} (${categoryLists.length})\n`

        categoryLists.forEach((list, index) => {
          const isLast = index === categoryLists.length - 1
          const prefix = isLast ? "└─" : "├─"

          let listInfo = `${prefix} ${list.displayName}`

          // Add metadata
          const metadata: string[] = []
          if (list.wellknownListName === "defaultList") metadata.push("Default")
          if (list.wellknownListName === "flaggedEmails") metadata.push("Flagged Emails")
          if (list.isShared && list.isOwner) metadata.push("Shared by you")
          if (list.isShared && !list.isOwner) metadata.push("Shared with you")

          if (metadata.length > 0) {
            listInfo += ` [${metadata.join(", ")}]`
          }

          output += `   ${listInfo}\n`

          if (!isLast) {
            output += "   │\n"
          }
        })

        output += "\n"
      })

      // Add summary
      const totalLists = Object.values(organized).reduce((sum, l) => sum + l.length, 0)
      const totalCategories = Object.keys(organized).length

      output += "-".repeat(50) + "\n"
      output += `Summary: ${totalLists} lists in ${totalCategories} categories\n`

      if (includeIds) {
        // Add a section with IDs
        output += "\n\n📋 List IDs Reference:\n" + "-".repeat(50) + "\n"
        lists.forEach((list) => {
          output += `${list.displayName}: ${list.id}\n`
        })
      }

      return { content: [{ type: "text", text: output }] }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching organized task lists: ${error}`,
          },
        ],
      }
    }
  },
)

registerTool(
  "write",
  "create-task-list",
  "Create a new task list (top-level container) in Microsoft Todo to help organize your tasks into categories or projects.",
  {
    displayName: z.string().describe("Name of the new task list"),
  },
  async ({ displayName }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Prepare the request body
      const requestBody = {
        displayName,
      }

      // Make the API request to create the task list
      const response = await makeGraphRequest<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists`, token, "POST", requestBody)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create task list: ${displayName}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task list created successfully!\nName: ${response.displayName}\nID: ${response.id}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating task list: ${error}`,
          },
        ],
      }
    }
  },
)

registerTool(
  "write",
  "update-task-list",
  "Update the name of an existing task list (top-level container) in Microsoft Todo.",
  {
    listId: z.string().describe("ID of the task list to update"),
    displayName: z.string().describe("New name for the task list"),
  },
  async ({ listId, displayName }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Prepare the request body
      const requestBody = {
        displayName,
      }

      // Make the API request to update the task list
      const response = await makeGraphRequest<TaskList>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}`,
        token,
        "PATCH",
        requestBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to update task list with ID: ${listId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task list updated successfully!\nNew name: ${response.displayName}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating task list: ${error}`,
          },
        ],
      }
    }
  },
)

registerTool(
  "destructive",
  "delete-task-list",
  "Delete a task list (top-level container) from Microsoft Todo. This will remove the list and all tasks within it.",
  {
    listId: z.string().describe("ID of the task list to delete"),
  },
  async ({ listId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Make a DELETE request to the Microsoft Graph API
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}`
      console.error(`Deleting task list: ${url}`)

      // The DELETE method doesn't return a response body, so we expect null
      await makeGraphRequest<null>(url, token, "DELETE")

      // If we get here, the delete was successful (204 No Content)
      return {
        content: [
          {
            type: "text",
            text: `Task list with ID: ${listId} was successfully deleted.`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting task list: ${error}`,
          },
        ],
      }
    }
  },
)

registerTool(
  "read",
  "get-tasks",
  "Get tasks from a specific Microsoft Todo list. These are the main todo items that can contain checklist items (subtasks).",
  {
    listId: z.string().describe("ID of the task list"),
    filter: z.string().optional().describe("OData $filter query (e.g., 'status eq \\'completed\\'')"),
    select: z.string().optional().describe("Comma-separated list of properties to include (e.g., 'id,title,status')"),
    orderby: z.string().optional().describe("Property to sort by (e.g., 'createdDateTime desc')"),
    top: z.number().optional().describe("Maximum number of tasks to retrieve"),
    skip: z.number().optional().describe("Number of tasks to skip"),
    count: z.boolean().optional().describe("Whether to include a count of tasks"),
  },
  async ({ listId, filter, select, orderby, top, skip, count }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Build the query parameters
      const queryParams = new URLSearchParams()

      if (filter) queryParams.append("$filter", filter)
      if (select) queryParams.append("$select", select)
      if (orderby) queryParams.append("$orderby", orderby)
      if (top !== undefined) queryParams.append("$top", top.toString())
      if (skip !== undefined) queryParams.append("$skip", skip.toString())
      if (count !== undefined) queryParams.append("$count", count.toString())

      // Construct the URL with query parameters
      const queryString = queryParams.toString()
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks${queryString ? "?" + queryString : ""}`

      console.error(`Making request to: ${url}`)

      const response = await makeGraphRequest<{ value: Task[]; "@odata.count"?: number }>(url, token)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve tasks for list: ${listId}`,
            },
          ],
        }
      }

      const tasks = response.value || []
      if (tasks.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No tasks found in list with ID: ${listId}`,
            },
          ],
        }
      }

      // Format the tasks based on available properties
      const formattedTasks = tasks.map((task) => {
        // Default format
        let taskInfo = `ID: ${task.id}\nTitle: ${task.title}`

        // Add status if available
        if (task.status) {
          const status = task.status === "completed" ? "✓" : "○"
          taskInfo = `${status} ${taskInfo}`
        }

        // Add due date if available
        if (task.dueDateTime) {
          taskInfo += `\nDue: ${new Date(task.dueDateTime.dateTime).toLocaleDateString()}`
        }

        // Add importance if available
        if (task.importance) {
          taskInfo += `\nImportance: ${task.importance}`
        }

        // Add categories if available
        if (task.categories && task.categories.length > 0) {
          taskInfo += `\nCategories: ${task.categories.join(", ")}`
        }

        // Add body content if available and not empty
        if (task.body && task.body.content && task.body.content.trim() !== "") {
          const previewLength = 50
          const contentPreview =
            task.body.content.length > previewLength
              ? task.body.content.substring(0, previewLength) + "..."
              : task.body.content
          taskInfo += `\nDescription: ${contentPreview}`
        }

        return `${taskInfo}\n---`
      })

      // Add count information if requested and available
      let countInfo = ""
      if (count && response["@odata.count"] !== undefined) {
        countInfo = `Total count: ${response["@odata.count"]}\n\n`
      }

      return {
        content: [
          {
            type: "text",
            text: `Tasks in list ${listId}:\n\n${countInfo}${formattedTasks.join("\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching tasks: ${error}`,
          },
        ],
      }
    }
  },
)

registerTool(
  "write",
  "create-task",
  "Create a new task in a specific Microsoft Todo list. A task is the main todo item that can have a title, description, due date, and other properties.",
  {
    listId: z.string().describe("ID of the task list"),
    title: z.string().describe("Title of the task"),
    body: z.string().optional().describe("Description or body content of the task"),
    dueDateTime: z.string().optional().describe("Due date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    startDateTime: z.string().optional().describe("Start date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    importance: z.enum(["low", "normal", "high"]).optional().describe("Task importance"),
    isReminderOn: z.boolean().optional().describe("Whether to enable reminder for this task"),
    reminderDateTime: z.string().optional().describe("Reminder date and time in ISO format"),
    status: z
      .enum(["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"])
      .optional()
      .describe("Status of the task"),
    categories: z.array(z.string()).optional().describe("Categories associated with the task"),
  },
  async ({
    listId,
    title,
    body,
    dueDateTime,
    startDateTime,
    importance,
    isReminderOn,
    reminderDateTime,
    status,
    categories,
  }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Construct the task body with all supported properties
      const taskBody: any = { title }

      // Add optional properties if provided
      if (body) {
        taskBody.body = {
          content: body,
          contentType: "text",
        }
      }

      if (dueDateTime) {
        taskBody.dueDateTime = {
          dateTime: dueDateTime,
          timeZone: "UTC",
        }
      }

      if (startDateTime) {
        taskBody.startDateTime = {
          dateTime: startDateTime,
          timeZone: "UTC",
        }
      }

      if (importance) {
        taskBody.importance = importance
      }

      if (isReminderOn !== undefined) {
        taskBody.isReminderOn = isReminderOn
      }

      if (reminderDateTime) {
        taskBody.reminderDateTime = {
          dateTime: reminderDateTime,
          timeZone: "UTC",
        }
      }

      if (status) {
        taskBody.status = status
      }

      if (categories && categories.length > 0) {
        taskBody.categories = categories
      }

      const response = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks`,
        token,
        "POST",
        taskBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create task in list: ${listId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task created successfully!\nID: ${response.id}\nTitle: ${response.title}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating task: ${error}`,
          },
        ],
      }
    }
  },
)

registerTool(
  "write",
  "update-task",
  "Update an existing task in Microsoft Todo. Allows changing any properties of the task including title, due date, importance, etc.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task to update"),
    title: z.string().optional().describe("New title of the task"),
    body: z.string().optional().describe("New description or body content of the task"),
    dueDateTime: z.string().optional().describe("New due date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    startDateTime: z.string().optional().describe("New start date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    importance: z.enum(["low", "normal", "high"]).optional().describe("New task importance"),
    isReminderOn: z.boolean().optional().describe("Whether to enable reminder for this task"),
    reminderDateTime: z.string().optional().describe("New reminder date and time in ISO format"),
    status: z
      .enum(["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"])
      .optional()
      .describe("New status of the task"),
    categories: z.array(z.string()).optional().describe("New categories associated with the task"),
  },
  async ({
    listId,
    taskId,
    title,
    body,
    dueDateTime,
    startDateTime,
    importance,
    isReminderOn,
    reminderDateTime,
    status,
    categories,
  }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Construct the task update body with all provided properties
      const taskBody: any = {}

      // Add optional properties if provided
      if (title !== undefined) {
        taskBody.title = title
      }

      if (body !== undefined) {
        taskBody.body = {
          content: body,
          contentType: "text",
        }
      }

      if (dueDateTime !== undefined) {
        if (dueDateTime === "") {
          // Remove the due date by setting it to null
          taskBody.dueDateTime = null
        } else {
          taskBody.dueDateTime = {
            dateTime: dueDateTime,
            timeZone: "UTC",
          }
        }
      }

      if (startDateTime !== undefined) {
        if (startDateTime === "") {
          // Remove the start date by setting it to null
          taskBody.startDateTime = null
        } else {
          taskBody.startDateTime = {
            dateTime: startDateTime,
            timeZone: "UTC",
          }
        }
      }

      if (importance !== undefined) {
        taskBody.importance = importance
      }

      if (isReminderOn !== undefined) {
        taskBody.isReminderOn = isReminderOn
      }

      if (reminderDateTime !== undefined) {
        if (reminderDateTime === "") {
          // Remove the reminder date by setting it to null
          taskBody.reminderDateTime = null
        } else {
          taskBody.reminderDateTime = {
            dateTime: reminderDateTime,
            timeZone: "UTC",
          }
        }
      }

      if (status !== undefined) {
        taskBody.status = status
      }

      if (categories !== undefined) {
        taskBody.categories = categories
      }

      // Make sure we have at least one property to update
      if (Object.keys(taskBody).length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No properties provided for update. Please specify at least one property to change.",
            },
          ],
        }
      }

      const response = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
        token,
        "PATCH",
        taskBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to update task with ID: ${taskId} in list: ${listId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task updated successfully!\nID: ${response.id}\nTitle: ${response.title}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating task: ${error}`,
          },
        ],
      }
    }
  },
)

registerTool(
  "destructive",
  "delete-task",
  "Delete a task from a Microsoft Todo list. This will remove the task and all its checklist items (subtasks).",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task to delete"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Make a DELETE request to the Microsoft Graph API
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`
      console.error(`Deleting task: ${url}`)

      // The DELETE method doesn't return a response body, so we expect null
      await makeGraphRequest<null>(url, token, "DELETE")

      // If we get here, the delete was successful (204 No Content)
      return {
        content: [
          {
            type: "text",
            text: `Task with ID: ${taskId} was successfully deleted from list: ${listId}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting task: ${error}`,
          },
        ],
      }
    }
  },
)

registerTool(
  "read",
  "get-checklist-items",
  "Get checklist items (subtasks) for a specific task. Checklist items are smaller steps or components that belong to a parent task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Fetch the task first to get its title
      const taskResponse = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
        token,
      )

      const taskTitle = taskResponse ? taskResponse.title : "Unknown Task"

      // Fetch the checklist items
      const response = await makeGraphRequest<{ value: ChecklistItem[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
        token,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve checklist items for task: ${taskId}`,
            },
          ],
        }
      }

      const items = response.value || []
      if (items.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No checklist items found for task "${taskTitle}" (ID: ${taskId})`,
            },
          ],
        }
      }

      const formattedItems = items.map((item) => {
        const status = item.isChecked ? "✓" : "○"
        let itemInfo = `${status} ${item.displayName} (ID: ${item.id})`

        // Add creation date if available
        if (item.createdDateTime) {
          const createdDate = new Date(item.createdDateTime).toLocaleString()
          itemInfo += `\nCreated: ${createdDate}`
        }

        return itemInfo
      })

      return {
        content: [
          {
            type: "text",
            text: `Checklist items for task "${taskTitle}" (ID: ${taskId}):\n\n${formattedItems.join("\n\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching checklist items: ${error}`,
          },
        ],
      }
    }
  },
)

registerTool(
  "write",
  "create-checklist-item",
  "Create a new checklist item (subtask) for a task. Checklist items help break down a task into smaller, manageable steps.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    displayName: z.string().describe("Text content of the checklist item"),
    isChecked: z.boolean().optional().describe("Whether the item is checked off"),
  },
  async ({ listId, taskId, displayName, isChecked }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Prepare the request body
      const requestBody: any = {
        displayName,
      }

      if (isChecked !== undefined) {
        requestBody.isChecked = isChecked
      }

      // Make the API request to create the checklist item
      const response = await makeGraphRequest<ChecklistItem>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
        token,
        "POST",
        requestBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create checklist item for task: ${taskId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Checklist item created successfully!\nContent: ${response.displayName}\nID: ${response.id}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating checklist item: ${error}`,
          },
        ],
      }
    }
  },
)

registerTool(
  "write",
  "update-checklist-item",
  "Update an existing checklist item (subtask). Allows changing the text content or completion status of the subtask.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    checklistItemId: z.string().describe("ID of the checklist item to update"),
    displayName: z.string().optional().describe("New text content of the checklist item"),
    isChecked: z.boolean().optional().describe("Whether the item is checked off"),
  },
  async ({ listId, taskId, checklistItemId, displayName, isChecked }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Prepare the update body, including only the fields that are provided
      const requestBody: any = {}

      if (displayName !== undefined) {
        requestBody.displayName = displayName
      }

      if (isChecked !== undefined) {
        requestBody.isChecked = isChecked
      }

      // Make sure we have at least one property to update
      if (Object.keys(requestBody).length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No properties provided for update. Please specify either displayName or isChecked.",
            },
          ],
        }
      }

      // Make the API request to update the checklist item
      const response = await makeGraphRequest<ChecklistItem>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
        token,
        "PATCH",
        requestBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to update checklist item with ID: ${checklistItemId}`,
            },
          ],
        }
      }

      const statusText = response.isChecked ? "Checked" : "Not checked"

      return {
        content: [
          {
            type: "text",
            text: `Checklist item updated successfully!\nContent: ${response.displayName}\nStatus: ${statusText}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating checklist item: ${error}`,
          },
        ],
      }
    }
  },
)

registerTool(
  "destructive",
  "delete-checklist-item",
  "Delete a checklist item (subtask) from a task. This removes just the specific subtask, not the parent task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    checklistItemId: z.string().describe("ID of the checklist item to delete"),
  },
  async ({ listId, taskId, checklistItemId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Make a DELETE request to the Microsoft Graph API
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`
      console.error(`Deleting checklist item: ${url}`)

      // The DELETE method doesn't return a response body, so we expect null
      await makeGraphRequest<null>(url, token, "DELETE")

      // If we get here, the delete was successful (204 No Content)
      return {
        content: [
          {
            type: "text",
            text: `Checklist item with ID: ${checklistItemId} was successfully deleted from task: ${taskId}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting checklist item: ${error}`,
          },
        ],
      }
    }
  },
)

// Cross-list queries — Graph's To Do endpoint scopes every task query to a single list,
// so "search everything" and "what's due today" both mean fanning out over the lists and
// combining the results here. Shared plumbing for search-tasks and get-agenda:

const TASK_PAGE_SIZE = 100
// Cap the paging so one enormous list can't hang the whole fan-out.
const MAX_TASK_PAGES_PER_LIST = 10
// Querying a dozen lists at once reliably trips Graph's throttle, so keep the burst small.
const LIST_FETCH_CONCURRENCY = 4

interface TaskPage {
  value: Task[]
  "@odata.nextLink"?: string
}

/** Page through one list's tasks. Returns null only if the list couldn't be read at all. */
async function fetchTasksInList(token: string, list: TaskList): Promise<AgendaTask[] | null> {
  let url: string | undefined = `${MS_GRAPH_BASE}/me/todo/lists/${list.id}/tasks?$top=${TASK_PAGE_SIZE}`
  const collected: AgendaTask[] = []

  for (let page = 0; page < MAX_TASK_PAGES_PER_LIST && url; page++) {
    const response: TaskPage | null = await makeGraphRequest<TaskPage>(url, token)
    if (!response) return collected.length > 0 ? collected : null

    for (const task of response.value ?? []) {
      collected.push({ ...task, listId: list.id, listName: list.displayName })
    }

    url = response["@odata.nextLink"]
  }

  return collected
}

/** Resolve `listIds` entries against list IDs first, then display names — IDs are opaque
 * base64 blobs, so letting callers say "Work" instead saves a get-task-lists round trip. */
function selectLists(lists: TaskList[], selectors?: string[]): TaskList[] {
  if (!selectors || selectors.length === 0) return lists

  const wanted = new Set(selectors.map((selector) => selector.trim().toLowerCase()))
  return lists.filter((list) => wanted.has(list.id.toLowerCase()) || wanted.has(list.displayName.toLowerCase()))
}

interface CollectedTasks {
  tasks: AgendaTask[]
  lists: TaskList[]
  /** Lists that errored — reported to the caller so an empty result isn't read as "none". */
  failedLists: string[]
}

async function collectTasksAcrossLists(token: string, listSelectors?: string[]): Promise<CollectedTasks | null> {
  const listsResponse = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)
  if (!listsResponse?.value) return null

  const lists = selectLists(listsResponse.value, listSelectors)
  const results = await mapWithConcurrency(lists, LIST_FETCH_CONCURRENCY, async (list) => ({
    list,
    tasks: await fetchTasksInList(token, list),
  }))

  const tasks: AgendaTask[] = []
  const failedLists: string[] = []

  for (const result of results) {
    if (result.tasks === null) failedLists.push(result.list.displayName)
    else tasks.push(...result.tasks)
  }

  return { tasks, lists, failedLists }
}

/** One task as three lines: title, context, and the IDs a follow-up tool call needs. */
function formatCrossListTask(task: AgendaTask, options: { showDue?: boolean } = {}): string {
  const marker = task.status === "completed" ? "✓" : "○"
  const details = [`list: ${task.listName}`]

  const dueDay = dueDayKey(task.dueDateTime)
  if (options.showDue && dueDay) details.push(`due: ${dueDay}`)
  if (task.importance && task.importance !== "normal") details.push(`importance: ${task.importance}`)
  if (task.categories && task.categories.length > 0) details.push(`categories: ${task.categories.join(", ")}`)

  return `${marker} ${task.title}\n    ${details.join(" · ")}\n    id: ${task.id}\n    listId: ${task.listId}`
}

function formatFailedLists(failedLists: string[]): string {
  if (failedLists.length === 0) return ""
  return `\n\n⚠️ Could not read ${failedLists.length} list(s), so results may be incomplete: ${failedLists.join(", ")}`
}

registerTool(
  "read",
  "search-tasks",
  "Search tasks by text across every Microsoft Todo list at once, or a chosen subset. Use this when you know roughly what a task is called but not which list holds it — get-tasks requires a listId, this does not. Matches the title and categories (and optionally the description) case-insensitively; every whitespace-separated term must appear, in any order.",
  {
    query: z
      .string()
      .describe("Text to match. Multiple words must all appear somewhere, in any order (e.g. 'tax invoice')"),
    listIds: z
      .array(z.string())
      .optional()
      .describe("Restrict the search to these lists — either list IDs or exact list names. Default: all lists"),
    searchBody: z.boolean().optional().default(false).describe("Also search task descriptions (default: false)"),
    includeCompleted: z.boolean().optional().default(false).describe("Include completed tasks (default: false)"),
    importance: z.enum(["low", "normal", "high"]).optional().describe("Only return tasks with this importance"),
    dueBefore: z.string().optional().describe("Only tasks due on or before this date (YYYY-MM-DD)"),
    dueAfter: z.string().optional().describe("Only tasks due on or after this date (YYYY-MM-DD)"),
    limit: z.number().min(1).max(200).optional().default(50).describe("Maximum results to return (default: 50)"),
  },
  async ({ query, listIds, searchBody, includeCompleted, importance, dueBefore, dueAfter, limit }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return { content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }] }
      }

      const collected = await collectTasksAcrossLists(token, listIds)
      if (!collected) {
        return { content: [{ type: "text", text: "Failed to retrieve task lists" }] }
      }

      if (collected.lists.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No lists matched ${JSON.stringify(listIds)}. Run get-task-lists to see the available lists.`,
            },
          ],
        }
      }

      // Day keys are YYYY-MM-DD, so plain string comparison orders them correctly.
      const before = dueBefore?.slice(0, 10)
      const after = dueAfter?.slice(0, 10)

      const matches = collected.tasks
        .filter((task) => {
          if (!includeCompleted && task.status === "completed") return false
          if (importance && task.importance !== importance) return false
          if (!matchesQuery(task, query, { searchBody })) return false

          if (before || after) {
            const dueDay = dueDayKey(task.dueDateTime)
            if (!dueDay) return false
            if (before && dueDay > before) return false
            if (after && dueDay < after) return false
          }

          return true
        })
        .sort(compareTasksByDue)

      const shown = matches.slice(0, limit)
      const scope = listIds?.length ? `${collected.lists.length} selected list(s)` : `${collected.lists.length} lists`

      if (shown.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No tasks matching "${query}" in ${scope}.${
                includeCompleted ? "" : " Completed tasks were excluded — set includeCompleted to include them."
              }${formatFailedLists(collected.failedLists)}`,
            },
          ],
        }
      }

      const truncation =
        matches.length > shown.length
          ? `\n\nShowing the first ${shown.length} of ${matches.length} matches — raise limit or narrow the query for more.`
          : ""

      return {
        content: [
          {
            type: "text",
            text:
              `Found ${matches.length} task(s) matching "${query}" in ${scope}:\n\n` +
              shown.map((task) => formatCrossListTask(task, { showDue: true })).join("\n\n") +
              truncation +
              formatFailedLists(collected.failedLists),
          },
        ],
      }
    } catch (error) {
      return { content: [{ type: "text", text: `Error searching tasks: ${error}` }] }
    }
  },
)

const AGENDA_SECTIONS: Array<{ bucket: AgendaBucket; heading: string }> = [
  { bucket: "overdue", heading: "⚠️ Overdue" },
  { bucket: "today", heading: "📅 Today" },
  { bucket: "tomorrow", heading: "➡️ Tomorrow" },
  { bucket: "upcoming", heading: "🗓️ Upcoming" },
  { bucket: "noDueDate", heading: "📥 No due date" },
]

registerTool(
  "read",
  "get-agenda",
  "Roll up everything due across all Microsoft Todo lists into overdue / today / tomorrow / upcoming sections. Answers 'what's on my plate?' in one call, without needing a listId or one get-tasks call per list. Overdue tasks are always included regardless of how far back they go.",
  {
    days: z
      .number()
      .min(0)
      .max(90)
      .optional()
      .default(7)
      .describe("How many days ahead to include beyond today (default: 7)"),
    listIds: z
      .array(z.string())
      .optional()
      .describe("Restrict to these lists — either list IDs or exact list names. Default: all lists"),
    includeNoDueDate: z
      .boolean()
      .optional()
      .default(false)
      .describe("Add a section for tasks with no due date (default: false)"),
    includeCompleted: z.boolean().optional().default(false).describe("Include completed tasks (default: false)"),
    timeZone: z
      .string()
      .optional()
      .describe(
        "IANA time zone deciding which day counts as 'today', e.g. 'America/Chicago'. Defaults to MSTODO_TIMEZONE, then the server's local zone",
      ),
  },
  async ({ days, listIds, includeNoDueDate, includeCompleted, timeZone }) => {
    try {
      const zone = timeZone ?? process.env.MSTODO_TIMEZONE
      if (zone && !isValidTimeZone(zone)) {
        return {
          content: [
            {
              type: "text",
              text: `Unknown time zone: "${zone}". Use an IANA name such as 'America/Chicago' or 'Europe/London'.`,
            },
          ],
        }
      }

      const token = await getAccessToken()
      if (!token) {
        return { content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }] }
      }

      const collected = await collectTasksAcrossLists(token, listIds)
      if (!collected) {
        return { content: [{ type: "text", text: "Failed to retrieve task lists" }] }
      }

      if (collected.lists.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No lists matched ${JSON.stringify(listIds)}. Run get-task-lists to see the available lists.`,
            },
          ],
        }
      }

      const todayKey = dayKeyInZone(new Date(), zone)
      const buckets = new Map<AgendaBucket, AgendaTask[]>()

      for (const task of collected.tasks) {
        if (!includeCompleted && task.status === "completed") continue

        const bucket = bucketForTask(task, todayKey, days)
        if (!bucket) continue
        if (bucket === "noDueDate" && !includeNoDueDate) continue

        const existing = buckets.get(bucket)
        if (existing) existing.push(task)
        else buckets.set(bucket, [task])
      }

      const zoneLabel = zone ?? Intl.DateTimeFormat().resolvedOptions().timeZone
      const header = `Agenda for ${todayKey} (${zoneLabel}) — ${collected.lists.length} list(s), next ${days} day(s)`

      const sections = AGENDA_SECTIONS.flatMap(({ bucket, heading }) => {
        const tasks = buckets.get(bucket)
        if (!tasks || tasks.length === 0) return []

        const body = tasks.sort(compareTasksByDue).map((task) => formatCrossListTask(task, { showDue: true }))
        return [`${heading} (${tasks.length})\n\n${body.join("\n\n")}`]
      })

      if (sections.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `${header}\n\nNothing due. 🎉${formatFailedLists(collected.failedLists)}`,
            },
          ],
        }
      }

      const total = [...buckets.values()].reduce((sum, tasks) => sum + tasks.length, 0)

      return {
        content: [
          {
            type: "text",
            text: `${header}\n${total} task(s)\n\n` + sections.join("\n\n") + formatFailedLists(collected.failedLists),
          },
        ],
      }
    } catch (error) {
      return { content: [{ type: "text", text: `Error building agenda: ${error}` }] }
    }
  },
)

// Microsoft Planner — powers the "Assigned to me" view in the To Do app, which draws
// from Planner plans (e.g. a project's Planner board), not the native To Do lists above.
registerTool(
  "read",
  "get-assigned-planner-tasks",
  "Get Microsoft Planner tasks assigned to you across all plans (e.g. project boards). This is what powers the 'Assigned to me' view in the Microsoft To Do app — distinct from your native To Do lists.",
  {},
  async () => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<{ value: PlannerTask[] }>(`${MS_GRAPH_BASE}/me/planner/tasks`, token)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to retrieve Planner tasks",
            },
          ],
        }
      }

      const tasks = response.value || []
      if (tasks.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No Planner tasks assigned to you.",
            },
          ],
        }
      }

      // Plan titles aren't included on the task objects — resolve each unique plan once
      const uniquePlanIds = [...new Set(tasks.map((t) => t.planId))]
      const plans = await Promise.all(
        uniquePlanIds.map((planId) => makeGraphRequest<PlannerPlan>(`${MS_GRAPH_BASE}/planner/plans/${planId}`, token)),
      )
      const planNames = new Map<string, string>()
      uniquePlanIds.forEach((planId, i) => planNames.set(planId, plans[i]?.title || planId))

      const formattedTasks = tasks.map((task) => {
        let taskInfo = `${formatPlannerStatus(task.percentComplete)} ${task.title}`
        taskInfo += `\nPlan: ${planNames.get(task.planId)}`
        if (task.dueDateTime) {
          taskInfo += `\nDue: ${new Date(task.dueDateTime).toLocaleDateString()}`
        }
        taskInfo += `\nPriority: ${task.priority} (${formatPlannerPriority(task.priority)})`
        if (task.checklistItemCount > 0) {
          taskInfo += `\nChecklist: ${task.activeChecklistItemCount}/${task.checklistItemCount} remaining`
        }
        taskInfo += `\nID: ${task.id}`
        return `${taskInfo}\n---`
      })

      return {
        content: [
          {
            type: "text",
            text: `Planner tasks assigned to you:\n\n${formattedTasks.join("\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching Planner tasks: ${error}`,
          },
        ],
      }
    }
  },
)

registerTool(
  "read",
  "get-planner-task-details",
  "Get the full description and checklist for a specific Microsoft Planner task. Use get-assigned-planner-tasks first to find the task ID.",
  {
    taskId: z.string().describe("ID of the Planner task"),
  },
  async ({ taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const [task, details] = await Promise.all([
        makeGraphRequest<PlannerTask>(`${MS_GRAPH_BASE}/planner/tasks/${taskId}`, token),
        makeGraphRequest<PlannerTaskDetails>(`${MS_GRAPH_BASE}/planner/tasks/${taskId}/details`, token),
      ])

      if (!task) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve Planner task with ID: ${taskId}`,
            },
          ],
        }
      }

      let output = `${formatPlannerStatus(task.percentComplete)} ${task.title}\n`
      output += `Priority: ${task.priority} (${formatPlannerPriority(task.priority)})\n`
      if (task.dueDateTime) {
        output += `Due: ${new Date(task.dueDateTime).toLocaleDateString()}\n`
      }

      if (details?.description) {
        output += `\nDescription:\n${details.description}\n`
      }

      const checklistEntries = Object.values(details?.checklist || {})
      if (checklistEntries.length > 0) {
        output += `\nChecklist:\n`
        checklistEntries.forEach((item) => {
          output += `${item.isChecked ? "✓" : "○"} ${item.title}\n`
        })
      }

      return { content: [{ type: "text", text: output }] }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching Planner task details: ${error}`,
          },
        ],
      }
    }
  },
)

// Unlike the To Do API, Planner requires an If-Match ETag on every write to prevent
// clobbering concurrent edits (tasks often live on shared/team plans). We fetch the
// current ETag immediately before each PATCH and retry once on a 412 conflict.
async function updatePlannerTaskWithRetry(
  taskId: string,
  patchBody: Record<string, unknown>,
  token: string,
): Promise<PlannerTask | null> {
  for (let attempt = 0; attempt < 2; attempt++) {
    const current = await makeGraphRequest<PlannerTask & { "@odata.etag"?: string }>(
      `${MS_GRAPH_BASE}/planner/tasks/${taskId}`,
      token,
    )
    const etag = current?.["@odata.etag"]
    if (!etag) {
      return null
    }

    const response = await fetch(`${MS_GRAPH_BASE}/planner/tasks/${taskId}`, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        "If-Match": etag,
      },
      body: JSON.stringify(patchBody),
    })

    if (response.ok) {
      // Planner's PATCH returns 204 No Content on success; refetch to confirm the result.
      return await makeGraphRequest<PlannerTask>(`${MS_GRAPH_BASE}/planner/tasks/${taskId}`, token)
    }

    if (response.status !== 412) {
      const errorText = await response.text()
      console.error(`Planner task update failed: ${response.status} ${errorText}`)
      return null
    }

    console.error("Planner task ETag conflict (412), retrying with a fresh ETag...")
  }

  return null
}

registerTool(
  "write",
  "update-planner-task",
  "Update a Microsoft Planner task — progress, title, priority, or dates. Handles the ETag concurrency check Planner requires automatically. Use get-assigned-planner-tasks first to find the task ID.",
  {
    taskId: z.string().describe("ID of the Planner task to update"),
    title: z.string().optional().describe("New title of the task"),
    percentComplete: z
      .number()
      .min(0)
      .max(100)
      .optional()
      .describe("Progress percentage: 0 = not started, 1-99 = in progress, 100 = completed"),
    priority: z
      .number()
      .min(0)
      .max(10)
      .optional()
      .describe("Priority 0-10: 0-1 Urgent, 2-4 Important, 5-6 Medium, 7-10 Low"),
    dueDateTime: z
      .string()
      .optional()
      .describe("New due date in ISO format (e.g., 2026-12-31T23:59:59Z), or empty string to remove"),
    startDateTime: z
      .string()
      .optional()
      .describe("New start date in ISO format (e.g., 2026-12-31T23:59:59Z), or empty string to remove"),
  },
  async ({ taskId, title, percentComplete, priority, dueDateTime, startDateTime }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const patchBody: Record<string, unknown> = {}
      if (title !== undefined) patchBody.title = title
      if (percentComplete !== undefined) patchBody.percentComplete = percentComplete
      if (priority !== undefined) patchBody.priority = priority
      if (dueDateTime !== undefined) patchBody.dueDateTime = dueDateTime === "" ? null : dueDateTime
      if (startDateTime !== undefined) patchBody.startDateTime = startDateTime === "" ? null : startDateTime

      if (Object.keys(patchBody).length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No properties provided for update. Please specify at least one property to change.",
            },
          ],
        }
      }

      const updated = await updatePlannerTaskWithRetry(taskId, patchBody, token)

      if (!updated) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to update Planner task with ID: ${taskId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Planner task updated successfully!\n${formatPlannerStatus(updated.percentComplete)} ${updated.title}\nPriority: ${updated.priority} (${formatPlannerPriority(updated.priority)})`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating Planner task: ${error}`,
          },
        ],
      }
    }
  },
)

// Bulk archive completed tasks
registerTool(
  "destructive",
  "archive-completed-tasks",
  "Move completed tasks older than a specified number of days from one list to another (archive) list. Useful for cleaning up active lists while preserving historical tasks.",
  {
    sourceListId: z.string().describe("ID of the source list to archive tasks from"),
    targetListId: z.string().describe("ID of the target archive list"),
    olderThanDays: z
      .number()
      .min(0)
      .default(90)
      .describe("Archive tasks completed more than this many days ago (default: 90)"),
    dryRun: z
      .boolean()
      .optional()
      .default(false)
      .describe("If true, only preview what would be archived without making changes"),
  },
  async ({ sourceListId, targetListId, olderThanDays, dryRun }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Calculate cutoff date
      const cutoffDate = new Date()
      cutoffDate.setDate(cutoffDate.getDate() - olderThanDays)

      // Get all completed tasks from source list
      const tasksResponse = await makeGraphRequest<{ value: Task[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks?$filter=status eq 'completed'`,
        token,
      )

      if (!tasksResponse || !tasksResponse.value) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to retrieve tasks from source list",
            },
          ],
        }
      }

      // Filter tasks older than cutoff
      const tasksToArchive = tasksResponse.value.filter((task) => {
        if (!task.completedDateTime?.dateTime) return false
        const completedDate = new Date(task.completedDateTime.dateTime)
        return completedDate < cutoffDate
      })

      if (tasksToArchive.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No completed tasks found older than ${olderThanDays} days.`,
            },
          ],
        }
      }

      if (dryRun) {
        // Preview mode - just show what would be archived
        let preview = `📋 Archive Preview\n`
        preview += `Would archive ${tasksToArchive.length} tasks completed before ${cutoffDate.toLocaleDateString()}\n\n`

        tasksToArchive.forEach((task) => {
          const completedDate = task.completedDateTime?.dateTime
            ? new Date(task.completedDateTime.dateTime).toLocaleDateString()
            : "Unknown"
          preview += `- ${task.title} (completed: ${completedDate})\n`
        })

        return { content: [{ type: "text", text: preview }] }
      }

      // Actually archive the tasks
      let successCount = 0
      const failedTasks: string[] = []

      for (const task of tasksToArchive) {
        try {
          // Create task in target list
          const createResponse = await makeGraphRequest(
            `${MS_GRAPH_BASE}/me/todo/lists/${targetListId}/tasks`,
            token,
            "POST",
            {
              title: task.title,
              status: "completed",
              body: task.body,
              importance: task.importance,
              completedDateTime: task.completedDateTime,
              dueDateTime: task.dueDateTime,
              reminderDateTime: task.reminderDateTime,
              categories: task.categories,
            },
          )

          if (createResponse) {
            // Delete from source list
            await makeGraphRequest(`${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks/${task.id}`, token, "DELETE")
            successCount++
          } else {
            failedTasks.push(task.title)
          }
        } catch (error) {
          failedTasks.push(task.title)
        }
      }

      let result = `📦 Archive Complete\n`
      result += `Successfully archived ${successCount} of ${tasksToArchive.length} tasks\n`
      result += `Tasks completed before ${cutoffDate.toLocaleDateString()} were moved.\n`

      if (failedTasks.length > 0) {
        result += `\n⚠️ Failed to archive ${failedTasks.length} tasks:\n`
        failedTasks.forEach((title) => {
          result += `- ${title}\n`
        })
      }

      return { content: [{ type: "text", text: result }] }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error archiving tasks: ${error}`,
          },
        ],
      }
    }
  },
)

// Test tool to explore Graph API for hidden properties
registerTool(
  "read",
  "test-graph-api-exploration",
  "Test various Graph API queries to discover hidden properties or endpoints for folder/group organization in Microsoft To Do.",
  {
    testType: z.enum(["odata-select", "odata-expand", "headers", "extensions", "all"]).describe("Type of test to run"),
  },
  async ({ testType }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      let results = "🔍 Graph API Exploration Results\n" + "=".repeat(50) + "\n\n"

      // Test 1: Try with $select=* to get all properties
      if (testType === "odata-select" || testType === "all") {
        results += "📊 Test 1: Using $select=* to retrieve all properties\n"
        try {
          const response = await makeGraphRequest<any>(`${MS_GRAPH_BASE}/me/todo/lists?$select=*`, token)
          if (response && response.value && response.value.length > 0) {
            const firstList = response.value[0]
            const properties = Object.keys(firstList)
            results += `Found ${properties.length} properties: ${properties.join(", ")}\n`

            // Show full first list as example
            results += "\nExample list object:\n"
            results += JSON.stringify(firstList, null, 2).substring(0, 1000) + "...\n"
          }
        } catch (error) {
          results += `Error: ${error}\n`
        }
        results += "\n"
      }

      // Test 2: Try various $expand options
      if (testType === "odata-expand" || testType === "all") {
        results += "📊 Test 2: Using $expand to retrieve related data\n"
        const expandOptions = [
          "extensions",
          "singleValueExtendedProperties",
          "multiValueExtendedProperties",
          "openExtensions",
          "parent",
          "children",
          "folder",
          "parentFolder",
          "group",
          "category",
        ]

        for (const expand of expandOptions) {
          try {
            const response = await makeGraphRequest<any>(
              `${MS_GRAPH_BASE}/me/todo/lists?$expand=${expand}&$top=1`,
              token,
            )
            if (response && response.value) {
              results += `✓ $expand=${expand}: Success - `
              if (response.value.length > 0 && response.value[0][expand]) {
                results += `Found data!\n`
                results += JSON.stringify(response.value[0][expand], null, 2).substring(0, 500) + "...\n"
              } else {
                results += `No additional data returned\n`
              }
            }
          } catch (error: any) {
            results += `✗ $expand=${expand}: ${error.message || "Failed"}\n`
          }
        }
        results += "\n"
      }

      // Test 3: Check response headers for additional info
      if (testType === "headers" || testType === "all") {
        results += "📊 Test 3: Checking response headers\n"
        try {
          const response = await fetch(`${MS_GRAPH_BASE}/me/todo/lists`, {
            headers: {
              Authorization: `Bearer ${token}`,
              Accept: "application/json",
              Prefer: "return=representation",
            },
          })

          results += "Response headers:\n"
          response.headers.forEach((value, key) => {
            results += `${key}: ${value}\n`
          })
        } catch (error) {
          results += `Error: ${error}\n`
        }
        results += "\n"
      }

      // Test 4: Try extensions endpoint
      if (testType === "extensions" || testType === "all") {
        results += "📊 Test 4: Checking for extensions\n"
        try {
          const listsResponse = await makeGraphRequest<{ value: TaskList[] }>(
            `${MS_GRAPH_BASE}/me/todo/lists?$top=1`,
            token,
          )

          if (listsResponse && listsResponse.value && listsResponse.value.length > 0) {
            const listId = listsResponse.value[0].id

            // Try to get extensions
            try {
              const extResponse = await makeGraphRequest<any>(
                `${MS_GRAPH_BASE}/me/todo/lists/${listId}/extensions`,
                token,
              )
              results += `Extensions found: ${JSON.stringify(extResponse, null, 2)}\n`
            } catch (error: any) {
              results += `No extensions endpoint: ${error.message}\n`
            }
          }
        } catch (error) {
          results += `Error: ${error}\n`
        }
        results += "\n"
      }

      // Test 5: Check if there's a separate folders or groups endpoint
      if (testType === "all") {
        results += "📊 Test 5: Checking for folder/group endpoints\n"
        const endpoints = [
          "/me/todo/folders",
          "/me/todo/groups",
          "/me/todo/listGroups",
          "/me/todo/listFolders",
          "/me/todo/categories",
        ]

        for (const endpoint of endpoints) {
          try {
            const response = await makeGraphRequest<any>(`${MS_GRAPH_BASE}${endpoint}`, token)
            results += `✓ ${endpoint}: Found! Response: ${JSON.stringify(response).substring(0, 200)}...\n`
          } catch (error: any) {
            results += `✗ ${endpoint}: Not found (${error.message || "Failed"})\n`
          }
        }
      }

      results += "\n" + "=".repeat(50) + "\n"
      results += "Analysis complete. Check results above for any discovered properties or endpoints."

      return {
        content: [
          {
            type: "text",
            text: results,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error during Graph API exploration: ${error}`,
          },
        ],
      }
    }
  },
)

// Main function to start the server
export async function startServer(): Promise<void> {
  try {
    console.error(`Access mode: ${describeAccessMode(accessMode)}`)
    if (withheldTools.length > 0) {
      console.error(`Withheld ${withheldTools.length} tool(s): ${withheldTools.join(", ")}`)
    }

    // Check if using a personal Microsoft account and show warning if needed
    await isPersonalMicrosoftAccount()

    // Start the server
    const transport = new StdioServerTransport()
    await server.connect(transport)

    console.error("Server started and listening")
  } catch (error) {
    console.error("Error starting server:", error)
    throw error
  }
}

// This module deliberately has no "am I the entry point, then start" block. Two things
// make that unworkable here: tsdown bundles the entries and splits shared code, so
// `dist/todo-index.js` is only a re-export shim around a hashed chunk whose
// `import.meta.url` can never equal `process.argv[1]`; and `main` should stay importable
// without the side effect of connecting a stdio transport. `src/cli.ts` (`dist/cli.js`)
// is the single runnable entry — it's what `bin` and `pnpm start` both point at.
