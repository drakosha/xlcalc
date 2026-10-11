/**
 * A window over a saved version of a book (`viewportOf`): the book and its result as a host keeps them, read as a
 * session's window reads its state (`viewport.ts`), without a session and without computing. The values, the text, the
 * colour and the reasons of errors come from the result, the formats, the styles and the layout from the book; the
 * cells of a large spill that the result leaves out show its `spillFill`. A result with the old names (`issues`, a
 * cell's `format`) reads alike. What a book and a result give beside their cells (the checked book, its tables, each
 * sheet's large spills and extent, the reasons by cell) is found once and kept with the two objects, so a host that
 * serves several windows of one version passes the same objects.
 */
import type { BookInput, CellIssue, CellValue, SavedBook, SavedCellResult, SavedResult, SavedViewport, SheetExtent, ViewportCell, ViewportOptions } from '../contract.js'
import { holds } from '../array-areas.js'
import { epochOf } from '../dates.js'
import { run } from '../exec/budget.js'
import { array, checkBook, checkClock, checkSheet, checkString, InputError, object } from '../input.js'
import { issueWith, reasonSource, spilledReason, wayBack, type Place } from '../issues.js'
import { sheetLines } from '../layout/lines.js'
import { parseMerge } from '../layout/merges.js'
import { displayRules } from '../locale/rules.js'
import { formatAddress, parseAddress, readAddress } from '../model.js'
import { bookLooks, type Looks } from '../shown.js'
import { hasCell, parseRange, type Rect } from '../structure/area.js'
import { quoteSheet, sheetKey } from '../syntax.js'
import { tablesOf, type Tables } from '../tables/model.js'
import { checkRange, checkViewportOptions, windowLines, windowOf, type WindowSource } from './viewport.js'

/** The cells of a large spill that its result leaves out (`CellResult.spillFill`): the spill's range, its formula's cell and their value. */
interface Fill {
  x: Rect
  from: string
  value: CellValue
}

/** What a saved sheet gives beside its cells: its large spills and where its cells end. */
interface SheetFound {
  fills: Fill[]
  extent: SheetExtent
}

/** What a saved book and its result give beside their cells. */
interface Found {
  book: BookInput
  result: SavedResult
  tables: Tables
  /** The sheets' indexes by key (`sheetKey`). */
  index: Map<string, number>
  /** What the cells show, when the result has their text (it was computed with the display). */
  looks: Looks | null
  sheets: Map<number, SheetFound>
  /** The reasons of the cells with an error, by `s:address`, once a window asks for them. */
  reasons?: Map<string, CellIssue>
}

/** What each pair of a book and a result gave, by the book, then the result. */
const kept = new WeakMap<object, WeakMap<object, Found>>()

/** The result checked as one of `book`: its sheets the book's, in order, each with its cells; its reasons lists; its clock. */
function checkResult(v: unknown, book: BookInput, path: string): SavedResult {
  const r = object(v, path, 'a result such as compute or getResult gives')
  const sheets = array(r.sheets, `${path}.sheets`, 'an array of the sheets of the result')
  if (sheets.length !== book.sheets.length) throw new InputError(`${path}.sheets`, `expected the ${book.sheets.length} sheets of the book, got ${sheets.length}: the result is not of this book`)
  sheets.forEach((x, i) => {
    const p = `${path}.sheets[${i}]`
    const sheet = object(x, p, 'a sheet such as { name: "Sheet1", cells: {} }')
    if (sheetKey(checkString(sheet.name, `${p}.name`)) !== sheetKey(book.sheets[i].name))
      throw new InputError(`${p}.name`, `expected the book's sheet ${JSON.stringify(book.sheets[i].name)}, got ${JSON.stringify(sheet.name)}: the result is not of this book`)
    object(sheet.cells, `${p}.cells`, 'an object of cells by address')
  })
  for (const k of ['diagnostics', 'issues']) if (r[k] !== undefined) array(r[k], `${path}.${k}`, 'an array of reasons')
  checkClock(r.clock, `${path}.clock`)
  return r as unknown as SavedResult
}

/** Whether a result has its cells' text: its first cell has it, or it has no cell (the display is on by default). */
function displayed(result: SavedResult): boolean {
  for (const sheet of result.sheets) for (const a in sheet.cells) return sheet.cells[a].text !== undefined
  return true
}

/** What `saved` gives beside its cells, checked and found once for its book and result. */
function found(saved: unknown): Found {
  const o = object(saved, 'saved', 'a saved book such as { book, result }')
  const byResult = typeof o.book === 'object' && o.book !== null ? kept.get(o.book) : undefined
  const hit = typeof o.result === 'object' && o.result !== null ? byResult?.get(o.result) : undefined
  if (hit) return hit
  const book = checkBook(o.book, 'saved.book')
  const result = checkResult(o.result, book, 'saved.result')
  const looks = displayed(result) ? bookLooks(book, displayRules(book.locale, epochOf(book.dateSystem))) : null
  const f: Found = { book, result, tables: tablesOf(book), index: new Map(book.sheets.map((x, i) => [sheetKey(x.name), i])), looks, sheets: new Map() }
  if (byResult) byResult.set(o.result as object, f)
  else kept.set(o.book as object, new WeakMap([[o.result as object, f]]))
  return f
}

/** Sheet `s` of a saved book: its large spills, and where its cells end as a session's `getUsedRange` finds it. */
function sheetFound(f: Found, s: number): SheetFound {
  let out = f.sheets.get(s)
  if (out) return out
  const input = f.book.sheets[s]
  const cells = f.result.sheets[s].cells
  const fills: Fill[] = []
  let [rows, columns] = [0, 0]
  const reach = (r: number, c: number) => {
    if (r >= rows) rows = r + 1
    if (c >= columns) columns = c + 1
  }
  const at = { r: 0, c: 0 }
  for (const a in cells) {
    const cell = cells[a]
    const own = input.cells[a]
    // a value, or an input cell that holds none (a style, a format, a note alone)
    if (cell.value !== null || (own && !holds(own))) {
      readAddress(a, at)
      reach(at.r, at.c)
    }
    // a spill's whole range, the cells its result leaves out too
    const x = cell.spill === undefined ? null : parseRange(cell.spill)
    if (!x) continue
    reach(x.r2, x.c2)
    if (cell.spillFill !== undefined) fills.push({ x, from: a, value: cell.spillFill })
  }
  for (const m of input.merges ?? []) {
    const x = parseMerge(m)
    if (x) reach(x.r2, x.c2)
  }
  for (const t of f.tables.list) if (t.s === s) reach(t.r2, t.c2)
  f.sheets.set(s, (out = { fills, extent: { rows, columns } }))
  return out
}

/** The result's reasons of the cells with an error, by `s:address`. */
function reasons(f: Found): Map<string, CellIssue> {
  if (f.reasons) return f.reasons
  const list: CellIssue[] = f.result.diagnostics?.flatMap((d) => (d.kind === 'cell' ? [d] : [])) ?? f.result.issues ?? []
  const out = new Map<string, CellIssue>()
  for (const i of list) {
    const s = f.index.get(sheetKey(i.sheet))
    if (s !== undefined) out.set(`${s}:${i.cell}`, i)
  }
  return (f.reasons = out)
}

/** The issue of a saved cell: the result's, or a session's for a cell of a large spill that the result leaves out. */
function issueAt(f: Found, p: Place): CellIssue | null {
  const addr = formatAddress(p.r, p.c)
  const known = reasons(f).get(`${p.s}:${addr}`)
  if (known) return known
  if (addr in f.result.sheets[p.s].cells) return null
  const fill = sheetFound(f, p.s).fills.find((x) => hasCell(x.x, p.r, p.c))
  const v = fill?.value
  if (!fill || v === null || typeof v !== 'object') return null
  const sheet = f.book.sheets[p.s].name
  return { sheet, cell: addr, error: v.error, reason: spilledReason(v.error, `${quoteSheet(sheet)}!${fill.from}`) }
}

/** The cell a saved cell's reason names, the next on the way back to where its error began; null where it began. */
function stepBack(f: Found, p: Place): Place | null {
  const i = issueAt(f, p)
  const from = i && reasonSource(i.reason)
  const s = from ? f.index.get(sheetKey(from.sheet)) : undefined
  if (s === undefined) return null
  const { row, col } = parseAddress(from!.cell)
  return { s, r: row, c: col }
}

/** A saved cell as a window gives it: its value and what the result shows of it, without its own number format. */
function resultCell(x: SavedCellResult): ViewportCell {
  const out: Record<string, unknown> = { value: x.value }
  for (const k of ['text', 'color', 'formula', 'meta', 'spill', 'spillFill', 'spillOf'] as const) if (x[k] !== undefined) out[k] = x[k]
  return out as unknown as ViewportCell
}

/** Sheet `s` of a saved book as a window reads it. */
function savedSource(f: Found, s: number): WindowSource {
  const sheet = sheetFound(f, s)
  const cells = f.result.sheets[s].cells
  const look = f.looks
  return {
    book: f.book,
    s,
    lines: sheetLines(f.book.sheets[s]),
    tables: f.tables,
    now: f.result.clock.now,
    cell(r, c, addr, own, styled) {
      const x = cells[addr]
      if (x) return resultCell(x)
      const fill = sheet.fills.find((y) => hasCell(y.x, r, c))
      if (fill) return { value: fill.value, ...look?.(s, r, c, fill.value, own), spillOf: fill.from }
      return own || styled ? { value: null, ...look?.(s, r, c, null, own) } : undefined
    },
    // a result holds no conditional appearance
    appearance() {},
    issue(r, c, origins) {
      const i = issueAt(f, { s, r, c })
      if (!i) return null
      const o = wayBack({ s, r, c }, i.error, (p) => stepBack(f, p), origins)
      return issueWith(i, { sheet: f.book.sheets[o.s].name, cell: formatAddress(o.r, o.c) })
    },
    usedRange: () => ({ ...sheet.extent }),
  }
}

/**
 * A window over a saved book and its result, as `Session.getViewport` gives it over a session of that book, without
 * computing anything: the result's values, text and colour, the book's formats, styles and layout.
 */
export function viewportOf(saved: SavedBook, sheet: string, range: string, opts?: ViewportOptions): SavedViewport {
  const f = found(saved)
  const s = checkSheet(sheet, (name) => f.index.get(sheetKey(name)), 'sheet')
  const asked = checkRange(range, 'range')
  const o = checkViewportOptions(opts)
  const w = windowLines(asked, o)
  return run(windowOf(savedSource(f, s), w, o))
}
