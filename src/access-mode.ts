// src/access-mode.ts
//
// Coarse capability gating for the MCP surface. Authentication hands this server the
// user's full Tasks.ReadWrite scope, so without this everything an assistant can reach is
// all-or-nothing. MSTODO_ACCESS_MODE narrows it at *registration* time: tools the mode
// doesn't permit are never registered, so they don't appear in tools/list and the model
// can't attempt them at all — which is more reliable than refusing the call afterwards.

/** How much of the tool surface to expose, set via MSTODO_ACCESS_MODE. */
export type AccessMode = "read" | "write" | "full"

/** What a tool does to the user's data, independent of the configured mode. */
export type ToolAccess = "read" | "write" | "destructive"

/**
 * Unset means `full` — every prior version of this server exposed the delete tools, so
 * defaulting to anything narrower would silently break existing installs on upgrade.
 */
export const DEFAULT_ACCESS_MODE: AccessMode = "full"

const MODE_ALIASES: Record<string, AccessMode> = {
  read: "read",
  readonly: "read",
  "read-only": "read",
  write: "write",
  readwrite: "write",
  "read-write": "write",
  update: "write",
  full: "full",
  all: "full",
}

// A mode permits every access level at or below its own rank.
const ACCESS_RANK: Record<ToolAccess, number> = { read: 0, write: 1, destructive: 2 }
const MODE_RANK: Record<AccessMode, number> = { read: 0, write: 1, full: 2 }

/**
 * Resolve the configured mode. A value we don't recognise throws rather than falling back
 * to the default: a typo in `MSTODO_ACCESS_MODE=raed` quietly granting full control is the
 * one failure mode this feature exists to prevent.
 */
export function parseAccessMode(raw: string | undefined | null): AccessMode {
  if (raw === undefined || raw === null || raw.trim() === "") return DEFAULT_ACCESS_MODE

  const mode = MODE_ALIASES[raw.trim().toLowerCase()]
  if (!mode) {
    throw new Error(
      `Invalid MSTODO_ACCESS_MODE: "${raw}". Expected one of: ` +
        `"read" (read-only), "write" (read plus create/update), "full" (adds delete/archive).`,
    )
  }

  return mode
}

export function isToolAllowed(access: ToolAccess, mode: AccessMode): boolean {
  return ACCESS_RANK[access] <= MODE_RANK[mode]
}

export function describeAccessMode(mode: AccessMode): string {
  switch (mode) {
    case "read":
      return "read — queries only; creating, updating and deleting are unavailable"
    case "write":
      return "write — read, create and update; deleting and archiving are unavailable"
    case "full":
      return "full — read, create, update, delete and archive"
  }
}
