import { describe, expect, it } from "vitest"

import { type AccessMode, describeAccessMode, isToolAllowed, parseAccessMode, type ToolAccess } from "./access-mode.js"

describe("parseAccessMode", () => {
  it("defaults to full when unset, so upgrades don't silently lose the delete tools", () => {
    expect(parseAccessMode(undefined)).toBe("full")
    expect(parseAccessMode(null)).toBe("full")
    expect(parseAccessMode("")).toBe("full")
    expect(parseAccessMode("   ")).toBe("full")
  })

  it("accepts the canonical modes", () => {
    expect(parseAccessMode("read")).toBe("read")
    expect(parseAccessMode("write")).toBe("write")
    expect(parseAccessMode("full")).toBe("full")
  })

  it("accepts aliases and is case/whitespace insensitive", () => {
    expect(parseAccessMode("READ-ONLY")).toBe("read")
    expect(parseAccessMode("readonly")).toBe("read")
    expect(parseAccessMode(" Read-Write ")).toBe("write")
    expect(parseAccessMode("readwrite")).toBe("write")
    expect(parseAccessMode("update")).toBe("write")
    expect(parseAccessMode("ALL")).toBe("full")
  })

  it("throws on an unrecognised value rather than falling back", () => {
    // A typo quietly granting full control is the failure this feature exists to prevent.
    expect(() => parseAccessMode("raed")).toThrow(/Invalid MSTODO_ACCESS_MODE: "raed"/)
    expect(() => parseAccessMode("none")).toThrow(/Expected one of/)
  })
})

describe("isToolAllowed", () => {
  const matrix: Array<[AccessMode, ToolAccess, boolean]> = [
    ["read", "read", true],
    ["read", "write", false],
    ["read", "destructive", false],
    ["write", "read", true],
    ["write", "write", true],
    ["write", "destructive", false],
    ["full", "read", true],
    ["full", "write", true],
    ["full", "destructive", true],
  ]

  it.each(matrix)("mode %s %s a %s tool", (mode, access, expected) => {
    expect(isToolAllowed(access, mode)).toBe(expected)
  })

  it("keeps read tools available in every mode", () => {
    const modes: AccessMode[] = ["read", "write", "full"]
    expect(modes.every((mode) => isToolAllowed("read", mode))).toBe(true)
  })
})

describe("describeAccessMode", () => {
  it("returns a distinct human-readable line per mode", () => {
    const descriptions = (["read", "write", "full"] as AccessMode[]).map(describeAccessMode)
    expect(new Set(descriptions).size).toBe(3)
    expect(descriptions.every((text) => text.length > 0)).toBe(true)
  })
})
