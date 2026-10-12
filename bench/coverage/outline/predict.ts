/**
 * An outline set's predictions, written before Excel runs the set: `npx tsx bench/coverage/outline/predict.ts <set> [grid]`
 * writes `outline/<set>.predicted.json`, each case's steps as `OutlineView`s. By default they are the engine's of the day
 * (`engineViews`: the commands as `editStructure`'s outline operations, the lines as `outlineGroups` reads them, the file
 * as `writeXlsx` writes it), so a set's committed predictions stay what the engine answered before Excel ran it. With
 * `grid`, the grid's own model (`xlcalc-sheet/src/local/outline.ts`: `groupsOf` and `outlineEdit`, ported here), for
 * `rules-held` and `ru-held`, before the engine had the commands: a group is the longest run of lines at its level or
 * deeper; Group puts each line one level deeper, refused at the seventh; Ungroup takes one off the lines that have one,
 * refused when none has; collapsing hides a group and marks its summary line; expanding shows it but for inner groups whose
 * own summary line keeps a mark; a level button k hides the lines of level k and deeper and marks the summary lines in
 * view of the collapsed groups. What the grid's model does not say was guessed: `ShowDetail` on a summary line acts on
 * its innermost group, on other lines on the innermost group holding them all (refused without one); a summary line's
 * `ShowDetail` reads False when all the lines of its innermost group are hidden; Excel saves the mark on the summary line
 * of such a group too.
 */
import { fileURLToPath } from 'node:url'
import { lines as spans } from '../../../src/layout/spans.js'
import { indexToCol, MAX_COLS, MAX_ROWS } from '../../../src/model.js'
import { setModule, writeAnswers } from '../stored.js'
import { engineViews, linesOf, windowOf, type LineView, type OutlineBook, type OutlineCase, type OutlineCommand, type OutlineSet, type OutlineView, type SavedLine } from './cases.js'

const [set, model] = process.argv.slice(2)
const { BOOKS, buildOutlineCases } = await setModule<OutlineSet>('outline', set)

interface Line {
  level: number
  hidden: boolean
  collapsed: boolean
}
interface Group {
  from: number
  to: number
  level: number
}
/** One axis: its lines with anything (1-based), whether summaries follow their groups, its last line. */
interface Axis {
  lines: Map<number, Line>
  after: boolean
  last: number
}

const MAX_LEVEL = 7
const at = (a: Axis, i: number): Line => a.lines.get(i) ?? { level: 0, hidden: false, collapsed: false }
function put(a: Axis, i: number, patch: Partial<Line>) {
  const l = { ...at(a, i), ...patch }
  if (l.level || l.hidden || l.collapsed) a.lines.set(i, l)
  else a.lines.delete(i)
}

/** `groupsOf`: for each level, the longest runs of lines next to each other at that level or deeper. */
function groupsOf(a: Axis): Group[] {
  const keys = [...a.lines.keys()].filter((i) => at(a, i).level > 0).sort((x, y) => x - y)
  const out: Group[] = []
  for (let level = 1; level <= MAX_LEVEL; level++) {
    let open: Group | null = null
    for (const i of keys) {
      if (at(a, i).level < level) continue
      if (open && i === open.to + 1) open.to = i
      else {
        if (open) out.push(open)
        open = { from: i, to: i, level }
      }
    }
    if (open) out.push(open)
  }
  return out
}

const summaryOf = (a: Axis, g: Group): number | null => {
  const s = a.after ? g.to + 1 : g.from - 1
  return s >= 1 && s <= a.last ? s : null
}
const allHidden = (a: Axis, g: Group) => {
  for (let i = g.from; i <= g.to; i++) if (!at(a, i).hidden) return false
  return true
}
/** The innermost group whose summary line is `s`. */
const innermostOf = (a: Axis, s: number) => groupsOf(a).filter((g) => summaryOf(a, g) === s).sort((x, y) => y.level - x.level)[0]

/** `outlineEdit`'s level button. */
function showLevel(a: Axis, k: number) {
  const groups = groupsOf(a)
  const shownAfter = (i: number) => (at(a, i).level > 0 ? at(a, i).level < k : !at(a, i).hidden)
  const marks = new Map<number, boolean>()
  for (const g of groups) {
    const s = summaryOf(a, g)
    if (s !== null) marks.set(s, (marks.get(s) ?? false) || (g.level >= k && shownAfter(s)))
  }
  for (const [i, l] of [...a.lines]) if (l.level > 0) put(a, i, { hidden: l.level >= k })
  for (const [s, mark] of marks) put(a, s, { collapsed: mark })
}

function collapse(a: Axis, g: Group) {
  for (let i = g.from; i <= g.to; i++) put(a, i, { hidden: true })
  const s = summaryOf(a, g)
  if (s !== null) put(a, s, { collapsed: true })
}

function expand(a: Axis, g: Group) {
  const s = summaryOf(a, g)
  const holes = groupsOf(a).filter((h) => h.level > g.level && h.from >= g.from && h.to <= g.to && summaryOf(a, h) !== s && at(a, summaryOf(a, h) ?? 0).collapsed)
  for (let i = g.from; i <= g.to; i++) if (!holes.some((h) => i >= h.from && i <= h.to)) put(a, i, { hidden: false })
  if (s !== null && at(a, s).collapsed) put(a, s, { collapsed: false })
}

/** A command on the book's axes; false when refused. */
function apply(rows: Axis, columns: Axis, c: OutlineCommand): boolean {
  if ('levels' in c) {
    if (c.levels[0]) showLevel(rows, c.levels[0])
    if (c.levels[1]) showLevel(columns, c.levels[1])
    return true
  }
  const [kind, text] = Object.entries(c)[0] as [string, string]
  const span = linesOf(text)
  const a = span.rows ? rows : columns
  const [lo, hi] = [span.lo + 1, span.hi + 1]
  const each = (f: (i: number) => void) => {
    for (let i = lo; i <= hi; i++) f(i)
  }
  if (kind === 'hidden' || kind === 'shown') return each((i) => put(a, i, { hidden: kind === 'hidden' })), true
  if (kind === 'group') {
    for (let i = lo; i <= hi; i++) if (at(a, i).level >= MAX_LEVEL) return false
    return each((i) => put(a, i, { level: at(a, i).level + 1 })), true
  }
  if (kind === 'ungroup') {
    let any = false
    each((i) => {
      if (at(a, i).level > 0) (any = true), put(a, i, { level: at(a, i).level - 1 })
    })
    return any
  }
  // ShowDetail: a summary line's innermost group, else the innermost group holding the lines
  const g = (lo === hi ? innermostOf(a, lo) : undefined) ?? groupsOf(a).filter((x) => x.from <= lo && x.to >= hi).sort((x, y) => y.level - x.level)[0]
  if (!g) return false
  if (kind === 'hide') collapse(a, g)
  else expand(a, g)
  return true
}

function axisOf(map: Record<string, { outlineLevel?: number; hidden?: boolean; collapsed?: boolean }> | undefined, axis: 'r' | 'c', after: boolean): Axis {
  const a: Axis = { lines: new Map(), after, last: axis === 'r' ? MAX_ROWS : MAX_COLS }
  for (const [lo, hi, e] of spans(map, axis)) for (let i = lo + 1; i <= hi + 1; i++) put(a, i, { level: e.outlineLevel ?? 0, hidden: e.hidden === true, collapsed: e.collapsed === true })
  return a
}

/** What Excel would read and save now. */
function view(b: OutlineBook, rows: Axis, columns: Axis, c: OutlineCase, refused: boolean): OutlineView {
  const w = windowOf(c.window)
  const read = (a: Axis, [lo, hi]: [number, number], key: (i: number) => string) => {
    const out: Record<string, LineView> = {}
    const summaries = new Set(groupsOf(a).map((g) => summaryOf(a, g)))
    for (let i = lo + 1; i <= hi + 1; i++) {
      const l = at(a, i)
      const g = summaries.has(i) ? innermostOf(a, i) : undefined
      if (l.level || l.hidden || g) out[key(i)] = [l.level, l.hidden, g ? !allHidden(a, g) : null]
    }
    return out
  }
  const saved = (a: Axis, [lo, hi]: [number, number], key: (i: number) => string) => {
    const out: Record<string, SavedLine> = {}
    const marked = new Set(groupsOf(a).filter((g) => allHidden(a, g)).map((g) => summaryOf(a, g)))
    for (let i = lo + 1; i <= hi + 1; i++) {
      const l = at(a, i)
      const mark = l.collapsed || marked.has(i)
      if (l.level || l.hidden || mark) out[key(i)] = [l.level, l.hidden, mark]
    }
    return out
  }
  const deepest = (a: Axis) => Math.max(0, ...[...a.lines.values()].map((l) => l.level))
  const col = (i: number) => indexToCol(i - 1)
  return {
    ...(refused ? { refused: true as const } : {}),
    rows: read(rows, w.rows, String),
    columns: read(columns, w.columns, col),
    saved: {
      rows: saved(rows, w.rows, String),
      columns: saved(columns, w.columns, col),
      summary: { below: b.outline?.summaryBelow !== false, right: b.outline?.summaryRight !== false },
      levels: [deepest(rows), deepest(columns)],
    },
  }
}

/** The grid's model of a case. */
function grid(b: OutlineBook, c: OutlineCase): OutlineView[] {
  const rows = axisOf(b.rows, 'r', b.outline?.summaryBelow !== false)
  const columns = axisOf(b.columns, 'c', b.outline?.summaryRight !== false)
  return [view(b, rows, columns, c, false), ...c.commands.map((x) => view(b, rows, columns, c, !apply(rows, columns, x)))]
}

const out: Record<string, OutlineView[]> = {}
for (const c of buildOutlineCases()) out[c.id] = model === 'grid' ? grid(BOOKS[c.book], c) : await engineViews(BOOKS[c.book], c)
writeAnswers(fileURLToPath(new URL(`./${set}.predicted.json`, import.meta.url)), { model: model === 'grid' ? 'xlcalc-sheet src/local/outline.ts' : 'editStructure', cases: out })
console.log(`${set}: ${Object.keys(out).length} cases predicted`)
