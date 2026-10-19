/**
 * A window over a saved version of a book (`viewportOf`): the book and its result as a host keeps them, read as a
 * session's window reads its state (`viewport.ts`), without a session and without computing. The values, the text, the
 * colour and the reasons of errors come from the result, the formats, the styles and the layout from the book; the
 * cells of a large spill that the result leaves out show its `spillFill`. With `appearance`, the conditional formats
 * are computed as a session computes them (`windowLooks`), over a state of the book whose cells hold the result's
 * values and no formula. What a book and a result give beside their cells (the checked book, its tables, each sheet's
 * large spills and extent, the reasons by cell, that state) is found once and kept with the two objects, so a host
 * that serves several windows of one version passes the same objects.
 */
import type { BookInput, BookSnapshot, CellIssue, CellResult, CellValue, ComputeResult, ConditionalAppearance, SavedViewport, SavedViewportOptions, SheetExtent, ViewportCell } from '../contract.js'
import { holds } from '../array-areas.js'
import { appearanceOf, windowLooks } from '../cond/looks.js'
import { epochOf } from '../dates.js'
import { createState, type EngineState } from '../engine.js'
import { put } from '../evaluate.js'
import { run } from '../exec/budget.js'
import { checkBook, checkClock, checkSheet } from '../input.js'
import { InputError } from '../input-error.js'
import { array, checkString, object } from '../input/basic.js'
import { checkRange } from '../input/range.js'
import { issueWith, reasonSource, spilledReason, wayBack, type Place } from '../issues.js'
import { sheetLines } from '../layout/lines.js'
import { parseMerge } from '../layout/merges.js'
import { displayRules } from '../locale/rules.js'
import { formatAddress, MAX_COLS, parseAddress, readAddress } from '../model.js'
import { clockSerial } from '../result.js'
import { bookLooks, type Looks } from '../shown.js'
import { hasCell, parseRange, type Rect } from '../structure/area.js'
import { quoteSheet, sheetKey } from '../syntax.js'
import { tablesOf, type Tables } from '../tables/model.js'
import { E, fromCellValue, Rng } from '../values.js'
import { layerOf } from '../looks/sheet.js'
import { checkViewportOptions, windowLines, windowOf, type WindowLines, type WindowSource } from './viewport.js'

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
  result: ComputeResult
  tables: Tables
  /** The sheets' indexes by key (`sheetKey`). */
  index: Map<string, number>
  /** What the cells show, when the result has their text (it was computed with the display). */
  looks: Looks | null
  sheets: Map<number, SheetFound>
  /** The reasons of the cells with an error, by `s:address`, once a window asks for them. */
  reasons?: Map<string, CellIssue>
  /** The state the conditional formats are computed over (`valuesState`), once a window asks for the appearance. */
  state?: EngineState
}

/** What each pair of a book and a result gave, by the book, then the result. */
const kept = new WeakMap<object, WeakMap<object, Found>>()

/** The result checked as one of `book`: its sheets the book's, in order, each with its cells; its reasons (`diagnostics`); its clock. */
function checkResult(v: unknown, book: BookInput, path: string): ComputeResult {
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
  array(r.diagnostics, `${path}.diagnostics`, 'an array of reasons')
  checkClock(r.clock, `${path}.clock`)
  return r as unknown as ComputeResult
}

/** Whether a result has its cells' text: its first cell has it, or it has no cell (the display is on by default). */
function displayed(result: ComputeResult): boolean {
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
  const looked = layerOf(input)?.bounds()
  if (looked) reach(looked.r2, looked.c2)
  f.sheets.set(s, (out = { fills, extent: { rows, columns } }))
  return out
}

/** The result's reasons of the cells with an error, by `s:address`. */
function reasons(f: Found): Map<string, CellIssue> {
  if (f.reasons) return f.reasons
  const list: CellIssue[] = f.result.diagnostics.flatMap((d) => (d.kind === 'cell' ? [d] : []))
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

/**
 * The state a session would hold over the saved book, for its conditional formats (`windowLooks`): the book's rules,
 * names, tables and layout on the result's clock, each cell holding the result's value (a large spill's unlisted cells
 * its `spillFill`) and no formula, so nothing is computed; `A1#` reads the spill the result gives A1.
 */
function valuesState(f: Found): EngineState {
  if (f.state) return f.state
  const { book, result } = f
  const state = createState({ ...book, sheets: book.sheets.map((x) => ({ ...x, cells: {} })) }, { now: clockSerial(result.clock), seed: result.clock.seed })
  const at = { r: 0, c: 0 }
  result.sheets.forEach(({ cells }, s) => {
    const data = state.ctx.sheets[s]
    for (const a in cells) {
      const x = cells[a]
      if (x.value === null) continue
      readAddress(a, at)
      put(data, at.r, at.c, fromCellValue(x.value))
      const g = x.spill === undefined || x.spillFill === undefined ? null : parseRange(x.spill)
      if (g) data.fills.push({ ...g, v: fromCellValue(x.spillFill!) })
    }
  })
  state.ctx.spillRef = (s, r, c) => {
    const a = formatAddress(r, c)
    if (book.sheets[s].cells[a]?.formula === undefined) return E.REF
    const x = result.sheets[s].cells[a]?.spill
    const g = x === undefined ? null : parseRange(x)
    return g ? new Rng(s, g.r1, g.c1, g.r2, g.c2) : new Rng(s, r, c, r, c)
  }
  return (f.state = state)
}

/** A saved cell as a window gives it: its value and what the result shows of it, without its own number format. */
function resultCell(x: CellResult): ViewportCell {
  const out: Record<string, unknown> = { value: x.value }
  for (const k of ['text', 'color', 'formula', 'meta', 'spill', 'spillFill', 'spillOf'] as const) if (x[k] !== undefined) out[k] = x[k]
  return out as unknown as ViewportCell
}

/** Sheet `s` of a saved book as a window of lines `w` reads it; with `appearance`, its cells' conditional appearance. */
function savedSource(f: Found, s: number, w: WindowLines, appearance: boolean): WindowSource {
  const sheet = sheetFound(f, s)
  const cells = f.result.sheets[s].cells
  const look = f.looks
  const cond = appearance ? run(windowLooks(valuesState(f), s, w.rects)) : new Map<number, ConditionalAppearance>()
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
      return own || styled || cond.has(r * MAX_COLS + c) ? { value: null, ...look?.(s, r, c, null, own) } : undefined
    },
    appearance: appearanceOf(cond, look, s),
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
 * computing its formulas: the result's values, text and colour, the book's formats, styles and layout; with
 * `appearance`, its conditional formats computed over the result's values.
 */
export function viewportOf(saved: BookSnapshot, sheet: string, range: string, opts?: SavedViewportOptions): SavedViewport {
  const f = found(saved)
  const s = checkSheet(sheet, (name) => f.index.get(sheetKey(name)), 'sheet')
  const asked = checkRange(range, 'range')
  const o = checkViewportOptions(opts, 'opts', true)
  const w = windowLines(asked, o)
  return run(windowOf(savedSource(f, s, w, o.appearance === true), w, o))
}
