// src/agenda.ts
//
// Helpers behind the `search-tasks` and `get-agenda` tools. Kept free of Graph calls so
// the fiddly parts — due-date bucketing, query matching, and the fan-out concurrency cap
// — are directly testable.

export interface DateTimeTimeZone {
  dateTime: string
  timeZone: string
}

/** A Graph todoTask plus the list it came from, since both tools span every list. */
export interface AgendaTask {
  id: string
  title: string
  status?: string
  importance?: string
  categories?: string[]
  body?: {
    content: string
    contentType: string
  }
  dueDateTime?: DateTimeTimeZone
  listId: string
  listName: string
}

export type AgendaBucket = "overdue" | "today" | "tomorrow" | "upcoming" | "noDueDate"

const DAY_MS = 86_400_000
const DAY_KEY_PATTERN = /^(\d{4})-(\d{2})-(\d{2})/

/**
 * The calendar day a task is due, as `YYYY-MM-DD`.
 *
 * Graph returns `dueDateTime` as a wall-clock time plus the zone it was written in, and
 * To Do stores a due *date* as midnight in that zone. Treating it as an instant and
 * converting into the viewer's zone shifts it a day for any negative UTC offset — a task
 * due Aug 5 would read as Aug 4 in New York — so take the date straight off the wall
 * clock instead. Returns null when there is no due date or the value is unparseable.
 */
export function dueDayKey(due: DateTimeTimeZone | undefined | null): string | null {
  const match = due?.dateTime ? DAY_KEY_PATTERN.exec(due.dateTime) : null
  return match ? `${match[1]}-${match[2]}-${match[3]}` : null
}

/** Today's `YYYY-MM-DD` in `timeZone` (an IANA name); the system zone when omitted. */
export function dayKeyInZone(instant: Date, timeZone?: string): string {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(instant)

  const valueOf = (type: Intl.DateTimeFormatPartTypes) => parts.find((part) => part.type === type)?.value ?? ""

  return `${valueOf("year")}-${valueOf("month")}-${valueOf("day")}`
}

/** Whether `timeZone` is an IANA zone this runtime knows, so callers can report it clearly. */
export function isValidTimeZone(timeZone: string): boolean {
  try {
    new Intl.DateTimeFormat("en-US", { timeZone }).format(new Date(0))
    return true
  } catch {
    return false
  }
}

/** Whole days from one `YYYY-MM-DD` to another; negative when `to` is earlier. */
export function daysBetweenDayKeys(from: string, to: string): number {
  return Math.round((Date.parse(`${to}T00:00:00Z`) - Date.parse(`${from}T00:00:00Z`)) / DAY_MS)
}

/**
 * Which agenda section a task belongs to, or null when its due date falls outside the
 * `upcomingDays` window. Overdue is never windowed — a task three months late still needs
 * to surface.
 */
export function bucketForTask(task: AgendaTask, todayKey: string, upcomingDays: number): AgendaBucket | null {
  const due = dueDayKey(task.dueDateTime)
  if (!due) return "noDueDate"

  const diff = daysBetweenDayKeys(todayKey, due)
  if (diff < 0) return "overdue"
  if (diff > upcomingDays) return null
  if (diff === 0) return "today"
  if (diff === 1) return "tomorrow"
  return "upcoming"
}

const IMPORTANCE_RANK: Record<string, number> = { high: 0, normal: 1, low: 2 }

/** Due soonest first, undated last, then high importance first, then title. */
export function compareTasksByDue(a: AgendaTask, b: AgendaTask): number {
  const dueA = dueDayKey(a.dueDateTime)
  const dueB = dueDayKey(b.dueDateTime)

  if (dueA !== dueB) {
    if (dueA === null) return 1
    if (dueB === null) return -1
    return dueA < dueB ? -1 : 1
  }

  const importanceA = IMPORTANCE_RANK[a.importance ?? "normal"] ?? 1
  const importanceB = IMPORTANCE_RANK[b.importance ?? "normal"] ?? 1
  if (importanceA !== importanceB) return importanceA - importanceB

  return a.title.localeCompare(b.title)
}

/**
 * Substring match, case-insensitive, across the title and categories — and the body when
 * `searchBody` is set. Whitespace splits the query into terms that must *all* match, so
 * "tax invoice" finds "Invoice for tax return" without needing the words adjacent.
 *
 * This is client-side because the To Do endpoint's $filter doesn't support `contains` on
 * task titles; callers fetch the lists and narrow here.
 */
export function matchesQuery(task: AgendaTask, query: string, options: { searchBody?: boolean } = {}): boolean {
  const terms = query.toLowerCase().split(/\s+/).filter(Boolean)
  if (terms.length === 0) return true

  const haystack = [task.title, ...(task.categories ?? []), options.searchBody ? (task.body?.content ?? "") : ""]
    .join("\n")
    .toLowerCase()

  return terms.every((term) => haystack.includes(term))
}

/**
 * Map over `items` with at most `limit` calls in flight, preserving input order.
 *
 * Both cross-list tools query every list, and Graph throttles an unbounded fan-out with
 * 429s — which surface as lists that silently contribute no tasks, i.e. a wrong answer
 * rather than an error. Capping concurrency keeps the burst under the throttle.
 */
export async function mapWithConcurrency<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const results = new Array<R>(items.length)
  let cursor = 0

  const workers = Array.from({ length: Math.max(1, Math.min(limit, items.length)) }, async () => {
    while (cursor < items.length) {
      const index = cursor++
      results[index] = await fn(items[index])
    }
  })

  await Promise.all(workers)
  return results
}
