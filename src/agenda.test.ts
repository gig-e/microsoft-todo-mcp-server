import { describe, expect, it } from "vitest"

import {
  type AgendaTask,
  bucketForTask,
  compareTasksByDue,
  dayKeyInZone,
  daysBetweenDayKeys,
  dueDayKey,
  isValidTimeZone,
  mapWithConcurrency,
  matchesQuery,
} from "./agenda.js"

function task(overrides: Partial<AgendaTask> = {}): AgendaTask {
  return {
    id: "task-1",
    title: "Write the report",
    status: "notStarted",
    importance: "normal",
    listId: "list-1",
    listName: "Work",
    ...overrides,
  }
}

function due(dateTime: string, timeZone = "UTC") {
  return { dateTime, timeZone }
}

describe("dueDayKey", () => {
  it("reads the calendar date off the wall clock, not the converted instant", () => {
    // To Do writes a due date as midnight in the stated zone. Parsing this as an instant
    // and rendering it in a negative-offset zone would report Aug 4.
    expect(dueDayKey(due("2026-08-05T00:00:00.0000000"))).toBe("2026-08-05")
  })

  it("handles a real Graph payload with a trailing Z and sub-second precision", () => {
    expect(dueDayKey(due("2026-12-31T23:59:59.9999999Z"))).toBe("2026-12-31")
  })

  it("returns null when there is no due date or the value is unparseable", () => {
    expect(dueDayKey(undefined)).toBeNull()
    expect(dueDayKey(null)).toBeNull()
    expect(dueDayKey(due(""))).toBeNull()
    expect(dueDayKey(due("not a date"))).toBeNull()
  })
})

describe("dayKeyInZone", () => {
  it("resolves the same instant to different days either side of the date line", () => {
    const instant = new Date("2026-08-05T02:00:00Z")
    expect(dayKeyInZone(instant, "UTC")).toBe("2026-08-05")
    expect(dayKeyInZone(instant, "America/New_York")).toBe("2026-08-04")
    expect(dayKeyInZone(instant, "Asia/Tokyo")).toBe("2026-08-05")
  })

  it("zero-pads month and day", () => {
    expect(dayKeyInZone(new Date("2026-01-02T12:00:00Z"), "UTC")).toBe("2026-01-02")
  })
})

describe("isValidTimeZone", () => {
  it("accepts IANA zones and rejects junk", () => {
    expect(isValidTimeZone("America/Chicago")).toBe(true)
    expect(isValidTimeZone("UTC")).toBe(true)
    expect(isValidTimeZone("Mars/Olympus_Mons")).toBe(false)
  })
})

describe("daysBetweenDayKeys", () => {
  it("counts whole days in both directions", () => {
    expect(daysBetweenDayKeys("2026-08-05", "2026-08-05")).toBe(0)
    expect(daysBetweenDayKeys("2026-08-05", "2026-08-06")).toBe(1)
    expect(daysBetweenDayKeys("2026-08-05", "2026-08-01")).toBe(-4)
  })

  it("crosses month and year boundaries", () => {
    expect(daysBetweenDayKeys("2026-01-31", "2026-02-01")).toBe(1)
    expect(daysBetweenDayKeys("2026-12-31", "2027-01-01")).toBe(1)
  })

  it("is unaffected by DST, since day keys are compared as UTC midnights", () => {
    // US DST starts 2026-03-08; a naive local-midnight subtraction would give 0.958 days.
    expect(daysBetweenDayKeys("2026-03-07", "2026-03-09")).toBe(2)
  })
})

describe("bucketForTask", () => {
  const today = "2026-08-05"

  it("sorts tasks into the expected sections", () => {
    expect(bucketForTask(task({ dueDateTime: due("2026-08-01T00:00:00Z") }), today, 7)).toBe("overdue")
    expect(bucketForTask(task({ dueDateTime: due("2026-08-05T00:00:00Z") }), today, 7)).toBe("today")
    expect(bucketForTask(task({ dueDateTime: due("2026-08-06T00:00:00Z") }), today, 7)).toBe("tomorrow")
    expect(bucketForTask(task({ dueDateTime: due("2026-08-09T00:00:00Z") }), today, 7)).toBe("upcoming")
    expect(bucketForTask(task(), today, 7)).toBe("noDueDate")
  })

  it("excludes tasks beyond the window", () => {
    expect(bucketForTask(task({ dueDateTime: due("2026-08-12T00:00:00Z") }), today, 7)).toBe("upcoming")
    expect(bucketForTask(task({ dueDateTime: due("2026-08-13T00:00:00Z") }), today, 7)).toBeNull()
  })

  it("never windows out overdue tasks, however late", () => {
    expect(bucketForTask(task({ dueDateTime: due("2019-01-01T00:00:00Z") }), today, 1)).toBe("overdue")
  })

  it("with days=0 keeps today but drops tomorrow", () => {
    expect(bucketForTask(task({ dueDateTime: due("2026-08-05T00:00:00Z") }), today, 0)).toBe("today")
    expect(bucketForTask(task({ dueDateTime: due("2026-08-06T00:00:00Z") }), today, 0)).toBeNull()
  })
})

describe("compareTasksByDue", () => {
  it("orders by due date, then importance, then title", () => {
    const tasks = [
      task({ id: "d", title: "Undated" }),
      task({ id: "c", title: "Zebra", dueDateTime: due("2026-08-06T00:00:00Z") }),
      task({ id: "b", title: "Apple", dueDateTime: due("2026-08-05T00:00:00Z") }),
      task({ id: "a", title: "Banana", dueDateTime: due("2026-08-05T00:00:00Z"), importance: "high" }),
    ]

    expect([...tasks].sort(compareTasksByDue).map((entry) => entry.id)).toEqual(["a", "b", "c", "d"])
  })

  it("puts undated tasks last even against far-future due dates", () => {
    const undated = task({ id: "undated" })
    const future = task({ id: "future", dueDateTime: due("2099-01-01T00:00:00Z") })
    expect([undated, future].sort(compareTasksByDue).map((entry) => entry.id)).toEqual(["future", "undated"])
  })
})

describe("matchesQuery", () => {
  const subject = task({
    title: "Invoice for tax return",
    categories: ["Finance"],
    body: { content: "Send to the accountant", contentType: "text" },
  })

  it("matches case-insensitively on the title", () => {
    expect(matchesQuery(subject, "INVOICE")).toBe(true)
  })

  it("requires every term but not adjacency or order", () => {
    expect(matchesQuery(subject, "tax invoice")).toBe(true)
    expect(matchesQuery(subject, "invoice receipt")).toBe(false)
  })

  it("matches categories", () => {
    expect(matchesQuery(subject, "finance")).toBe(true)
  })

  it("only searches the body when asked", () => {
    expect(matchesQuery(subject, "accountant")).toBe(false)
    expect(matchesQuery(subject, "accountant", { searchBody: true })).toBe(true)
  })

  it("treats an empty or whitespace query as matching everything", () => {
    expect(matchesQuery(subject, "")).toBe(true)
    expect(matchesQuery(subject, "   ")).toBe(true)
  })

  it("does not throw on tasks missing categories and body", () => {
    expect(matchesQuery(task(), "report")).toBe(true)
    expect(matchesQuery(task(), "report", { searchBody: true })).toBe(true)
  })
})

describe("mapWithConcurrency", () => {
  /** Resolves on demand so a test can observe how many calls are in flight at once. */
  function deferred<T>() {
    let resolve!: (value: T) => void
    const promise = new Promise<T>((r) => (resolve = r))
    return { promise, resolve }
  }

  it("returns results in input order regardless of completion order", async () => {
    const items = [30, 10, 20, 0]
    const results = await mapWithConcurrency(items, 2, async (ms) => {
      await new Promise((resolve) => setTimeout(resolve, ms))
      return ms
    })

    expect(results).toEqual(items)
  })

  it("never exceeds the concurrency limit", async () => {
    let inFlight = 0
    let peak = 0

    await mapWithConcurrency(
      Array.from({ length: 12 }, (_, i) => i),
      4,
      async (index) => {
        inFlight++
        peak = Math.max(peak, inFlight)
        await new Promise((resolve) => setTimeout(resolve, index % 3))
        inFlight--
        return index
      },
    )

    expect(peak).toBe(4)
  })

  it("starts a queued item as soon as a slot frees, rather than waiting for a batch", async () => {
    const gates = [deferred<void>(), deferred<void>(), deferred<void>()]
    const started: number[] = []

    const pending = mapWithConcurrency([0, 1, 2], 2, async (index) => {
      started.push(index)
      await gates[index].promise
      return index
    })

    await Promise.resolve()
    expect(started).toEqual([0, 1])

    // Freeing one slot admits the third item without waiting on the second.
    gates[0].resolve()
    await Promise.resolve()
    await Promise.resolve()
    expect(started).toEqual([0, 1, 2])

    gates[1].resolve()
    gates[2].resolve()
    expect(await pending).toEqual([0, 1, 2])
  })

  it("handles an empty list and a limit larger than the input", async () => {
    expect(await mapWithConcurrency([], 4, async (x) => x)).toEqual([])
    expect(await mapWithConcurrency([1, 2], 99, async (x) => x * 2)).toEqual([2, 4])
  })
})
