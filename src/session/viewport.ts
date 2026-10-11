/**
 * A window over a sheet for a grid (`Session.getViewport`, `viewportOf`): the cells of a rectangle, or of the rows and
 * columns asked of it, as a result gives them, each with its look resolved as it shows it (`layout/lines.ts`), styles by
 * index with the objects of those it shows beside, the number format as a code and where it comes from, an empty cell
 * with a style of its row or column listed too; the widths and heights of the window's lines, the merged areas and
 * tables that meet its cells, the sheet's frozen panes and where its cells end. A session's window reads its state
 * (`sessionSource`), a saved one a book and its result (`saved.ts`); either reads the window's cells, never the whole sheet.
 */
import type { BookInput, CellInput, CellIssueDetail, CellStyle, NumberFormat, SavedViewportOptions, SheetExtent, Viewport, ViewportCell, ViewportOptions } from '../contract.js'
import type { EngineState } from '../engine.js'
import { checkedEntryText } from '../entry/text.js'
import { paced, type Steps } from '../exec/budget.js'
import { array, checkBoolean, checkNumber, checkString, InputError, object, onlyOptions } from '../input.js'
import { issueDetail, type Place } from '../issues.js'
import { isHiddenRow } from '../layout/hidden.js'
import { formatSource, lookOf, sheetLines, type SheetLines } from '../layout/lines.js'
import { parseMerge } from '../layout/merges.js'
import { toLocal } from '../locale/index.js'
import { DEFAULT_LOCALE, LOCALES } from '../locale/registry.js'
import { formatCodeIn } from '../locale/rules.js'
import { formatAddress, MAX_COLS, MAX_ROWS, parseAddress } from '../model.js'
import { columnPoints, widthOf } from '../numfmt.js'
import { cellResult, spilled } from '../result.js'
import { cellLooks } from '../shown.js'
import { lowerBound } from '../sorted.js'
import { parseRange, rectText, type Rect } from '../structure/area.js'
import { tableAt, type Tables } from '../tables/model.js'
import { holds } from '../array-areas.js'
import { appearanceOf, windowLooks } from '../cond/looks.js'
import { sheetExtent } from './extent.js'

/** The cells a window gives by default, as many as the window of a sheet a server sends. */
const MAX_CELLS = 20_000

/** Excel's width of a column without its own (8 digits and the padding) and height of a row without its own, in points (Calibri 11). */
const DEFAULT_WIDTH = widthOf(columnPoints(undefined))
const DEFAULT_HEIGHT = 15

/** Line numbers asked of a window: whole numbers from 1 to `max`, ascending, each once. */
function checkLineNumbers(v: unknown, path: string, max: number) {
  const xs = array(v, path, 'an array of line numbers such as [2, 5, 9]')
  xs.forEach((n, i) => {
    checkNumber(n, `${path}[${i}]`, 1, max, true)
    if (i && (n as number) <= (xs[i - 1] as number)) throw new InputError(`${path}[${i}]`, `expected numbers in ascending order, each once: ${n} comes after ${xs[i - 1]}`)
  })
}

/** A rectangle of a sheet given as text (`B2`, `B2:D9`, `B:D`, `2:9`). */
export function checkRange(text: unknown, path: string): Rect {
  const x = parseRange(checkString(text, path))
  if (!x) throw new InputError(path, `expected a range such as "B2", "B2:D9", "B:D" or "2:9" within the sheet, got ${JSON.stringify(text)}`)
  return x
}

/** `getViewport`'s options; with `saved`, `viewportOf`'s (`appearance` too). */
export function checkViewportOptions(v: unknown, path = 'opts', saved = false): SavedViewportOptions {
  if (v === undefined) return {}
  const o = object(v, path, 'options such as { formulaLocal: true, maxCells: 20000 }')
  onlyOptions(o, path, ['formulaLocal', 'diagnostics', 'maxCells', 'rows', 'columns', 'editText', ...(saved ? ['appearance'] : [])])
  for (const k of ['formulaLocal', 'diagnostics', 'editText', 'appearance']) if (o[k] !== undefined) checkBoolean(o[k], `${path}.${k}`)
  if (o.maxCells !== undefined) checkNumber(o.maxCells, `${path}.maxCells`, 1, MAX_COLS * MAX_ROWS, true)
  if (o.rows !== undefined) checkLineNumbers(o.rows, `${path}.rows`, MAX_ROWS)
  if (o.columns !== undefined) checkLineNumbers(o.columns, `${path}.columns`, MAX_COLS)
  return o as SavedViewportOptions
}

/** The lines of a window, zero-based and in order, the rectangles their runs make, and the part of the range asked it gives. */
export interface WindowLines {
  rows: number[]
  cols: number[]
  rects: Rect[]
  range: string
  truncated: boolean
}

/** The runs of consecutive numbers of ascending `xs`, each as its first and last. */
function runs(xs: number[]): [number, number][] {
  const out: [number, number][] = []
  for (const x of xs) {
    const last = out.at(-1)
    if (last && x === last[1] + 1) last[1] = x
    else out.push([x, x])
  }
  return out
}

/** The lines `xs` asks (numbers from 1, checked as options), zero-based, each from `lo` to `hi`; null when it asks none. */
function askedLines(xs: number[] | undefined, lo: number, hi: number, path: string, line: string): number[] | null {
  return (
    xs?.map((n, i) => {
      if (n - 1 < lo || n - 1 > hi) throw new InputError(`${path}[${i}]`, `${line} ${n} is outside the range, ${line}s ${lo + 1} to ${hi + 1}`)
      return n - 1
    }) ?? null
  )
}

/** The lines of window `x` that the options ask (all, or `rows` and `columns`): of more than `maxCells` cells, its first rows that fit, or the first cells of its first row. */
export function windowLines(x: Rect, o: ViewportOptions): WindowLines {
  const rows = askedLines(o.rows, x.r1, x.r2, 'opts.rows', 'row')
  const cols = askedLines(o.columns, x.c1, x.c2, 'opts.columns', 'column')
  const [nr, nc] = [rows?.length ?? x.r2 - x.r1 + 1, cols?.length ?? x.c2 - x.c1 + 1]
  const max = o.maxCells ?? MAX_CELLS
  // a row counts one cell at least, so that rows without columns are not given past `maxCells`
  const wide = Math.max(nc, 1)
  const [kr, kc] = nr * wide <= max ? [nr, nc] : wide <= max ? [Math.floor(max / wide), nc] : [1, max]
  const r = rows?.slice(0, kr) ?? Array.from({ length: kr }, (_, i) => x.r1 + i)
  const c = cols?.slice(0, kc) ?? Array.from({ length: kc }, (_, i) => x.c1 + i)
  const truncated = kr < nr || kc < nc
  const given = truncated ? { ...x, r2: r.at(-1)!, c2: kc < nc ? c.at(-1)! : x.c2 } : x
  const colRuns = runs(c)
  const rects = runs(r).flatMap(([r1, r2]) => colRuns.map(([c1, c2]) => ({ r1, r2, c1, c2 })))
  return { rows: r, cols: c, rects, range: rectText(given), truncated }
}

/** Whether ascending `xs` holds a number from `lo` to `hi`. */
const hits = (xs: number[], lo: number, hi: number) => {
  const i = lowerBound(xs, lo, (x) => x)
  return i < xs.length && xs[i] <= hi
}

/** What a window reads of a sheet: a session's state, or a saved book and its result (`saved.ts`). */
export interface WindowSource {
  /** The book, checked, and the sheet's index in it. */
  book: BookInput
  s: number
  lines: SheetLines
  tables: Tables
  /** The moment whose year completes text read as a date without one (`ViewportCell.editText`). */
  now: string
  /** The cell as a result gives it, without its own number format; undefined for one with nothing of its own (no input, no spill, no conditional look) unless `styled`, its line's style shown. */
  cell(r: number, c: number, addr: string, own: CellInput | undefined, styled: boolean): ViewportCell | undefined
  /** What the sheet's conditional formats give the cell, put on it: its `appearance` and the text by a format of theirs. */
  appearance(out: ViewportCell, r: number, c: number, own: CellInput | undefined): void
  /** The cell's error with its reason and origin (`origins` shared by the cells of one window). */
  issue(r: number, c: number, origins: Map<string, Place>): CellIssueDetail | null
  usedRange(): SheetExtent
}

/** The window of lines `w` of a sheet from `src` (options checked). The steps yield between cells. */
export function* windowOf(src: WindowSource, w: WindowLines, o: ViewportOptions): Steps<Omit<Viewport, 'revision'>> {
  const { book, s, lines, tables } = src
  const sheet = book.sheets[s]
  // each style the window's cells show, copied once: the answer's own
  const styles: Record<number, CellStyle> = {}
  const shows = (i: number) => {
    styles[i] ??= JSON.parse(JSON.stringify(book.styles![i])) as CellStyle
    return i
  }
  const code = formatCodeIn(book.locale)
  // formulas in the book's locale, each text once per table it stands in (a table's column is one formula)
  const loc = o.formulaLocal || o.editText ? LOCALES.get(book.locale ?? DEFAULT_LOCALE)! : null
  const locals = new Map<string, string>()
  const local = (formula: string, r: number, c: number) => {
    const table = tableAt(tables, s, r, c)?.name
    const k = `${table ?? ''}\n${formula}`
    let out = locals.get(k)
    if (out === undefined) {
      const t = toLocal(formula, loc!, table === undefined ? {} : { table })
      locals.set(k, (out = t.ok ? t.formula : formula))
    }
    return out
  }
  // the text a cell opens with: its formula, an old-kind array's in each cell of its area, its own value by its format (a text, which is read as typed, once per format), nothing for a cell without either
  const entry = { locale: book.locale, dateSystem: book.dateSystem, now: src.now }
  const typed = new Map<string, string>()
  const editText = (out: ViewportCell, own: CellInput | undefined, numberFormat: NumberFormat | undefined, r: number, c: number) => {
    if (out.formula !== undefined) return `=${local(out.formula, r, c)}`
    const top = out.spillOf === undefined ? undefined : sheet.cells[out.spillOf]
    if (top?.array !== undefined && top.formula !== undefined) {
      const p = parseAddress(out.spillOf!)
      return `=${local(top.formula, p.row, p.col)}`
    }
    if (!holds(own)) return ''
    if (typeof out.value !== 'string') return checkedEntryText(out.value, { ...entry, numberFormat })
    const k = `${typeof numberFormat}:${numberFormat}\n${out.value}`
    let t = typed.get(k)
    if (t === undefined) typed.set(k, (t = checkedEntryText(out.value, { ...entry, numberFormat })))
    return t
  }
  // the cells errors began in, shared by the cells on one way to them
  const origins = new Map<string, Place>()
  const cells: Record<string, ViewportCell> = {}
  for (const r of w.rows) {
    const row = lines.row(r)
    for (const c of w.cols) {
      if (paced(1)) yield
      const addr = formatAddress(r, c)
      const own = sheet.cells[addr]
      const col = lines.column(c)
      const shown = lookOf(own, row, col)
      // an empty cell is listed when it shows something: a style of its row or column, a look, a note or a link
      const out = src.cell(r, c, addr, own, shown.style !== undefined)
      if (!out) continue
      if (o.formulaLocal && out.formula !== undefined) out.formulaLocal = local(out.formula, r, c)
      if (own?.array !== undefined) out.array = own.array
      if (own?.style !== undefined) out.style = own.style
      if (shown.style !== undefined) out.effectiveStyle = shows(shown.style)
      if (shown.numberFormat !== undefined) {
        out.numberFormat = code(shown.numberFormat)
        out.numberFormatSource = formatSource(own, row, col)
      }
      src.appearance(out, r, c, own)
      if (own?.note !== undefined) out.note = own.note
      if (own?.link !== undefined) out.link = own.link
      if (o.diagnostics && out.value !== null && typeof out.value === 'object') out.issue = src.issue(r, c, origins)!
      if (o.editText) out.editText = editText(out, own, shown.numberFormat, r, c)
      cells[addr] = out
    }
  }
  const crosses = (x: Rect) => hits(w.rows, x.r1, x.r2) && hits(w.cols, x.c1, x.c2)
  return {
    range: w.range,
    truncated: w.truncated,
    cells,
    styles,
    columns: w.cols.map((c) => {
      const col = lines.column(c)
      return col?.hidden === true || col?.width === 0 ? { column: c + 1, width: col.width ?? DEFAULT_WIDTH, hidden: true } : { column: c + 1, width: col?.width ?? DEFAULT_WIDTH }
    }),
    rows: w.rows.map((r) => {
      const row = lines.row(r)
      return row && isHiddenRow(row) ? { row: r + 1, height: row.height ?? DEFAULT_HEIGHT, hidden: true } : { row: r + 1, height: row?.height ?? DEFAULT_HEIGHT }
    }),
    merges: (sheet.merges ?? []).filter((m) => crosses(parseMerge(m)!)),
    tables: tables.list.filter((t) => t.s === s && crosses(t)).map((t) => ({ name: t.name, range: t.input.range, headers: t.headers, totals: t.totals })),
    frozen: { rows: sheet.frozen?.rows ?? 0, columns: sheet.frozen?.columns ?? 0 },
    usedRange: src.usedRange(),
  }
}

/** The window `asked` of a session's sheet `s` (options checked); `display`, the session's: the cells' text and colour; `now`, its clock's. */
export function* viewport(state: EngineState, s: number, asked: Rect, o: ViewportOptions, display: boolean, now: string): Steps<Omit<Viewport, 'revision'>> {
  const w = windowLines(asked, o)
  return yield* windowOf(yield* sessionSource(state, s, w, display, now), w, o)
}

/** A session's sheet `s` as a window of lines `w` reads it. */
function* sessionSource(state: EngineState, s: number, w: WindowLines, display: boolean, now: string): Steps<WindowSource> {
  const look = display ? cellLooks(state) : null
  const spill = new Map<string, string>()
  for (const x of w.rects) for (const [a, from] of yield* spilled(state, s, x, true)) spill.set(a, from)
  const cond = yield* windowLooks(state, s, w.rects)
  return {
    book: state.input,
    s,
    lines: sheetLines(state.input.sheets[s], state),
    tables: state.tables,
    now,
    cell(r, c, addr, own, styled) {
      const from = spill.get(addr)
      if (!own && from === undefined && !styled && !cond.has(r * MAX_COLS + c)) return undefined
      return cellResult(state, s, addr, look, from, false) as ViewportCell
    },
    appearance: appearanceOf(cond, look, s),
    issue: (r, c, origins) => issueDetail(state, s, r, c, origins),
    usedRange: () => sheetExtent(state, s),
  }
}
