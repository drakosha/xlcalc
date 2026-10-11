/**
 * A window over a sheet for a grid (`Session.getViewport`): the cells of a rectangle as a result gives them, each with
 * its look resolved as it shows it (`layout/lines.ts`), styles as objects, the number format as a code and where it
 * comes from, an empty cell with a style of its row or column listed too; the widths and heights of the window's lines,
 * the merged areas that meet it and where the sheet's cells end. It reads the window's cells, never the whole sheet.
 */
import type { CellStyle, Viewport, ViewportCell, ViewportOptions } from '../contract.js'
import type { EngineState } from '../engine.js'
import { paced, type Steps } from '../exec/budget.js'
import { checkBoolean, checkNumber, object, onlyKeys } from '../input.js'
import { issueDetail, type Place } from '../issues.js'
import { isHiddenRow } from '../layout/hidden.js'
import { formatSource, lookOf, sheetLines } from '../layout/lines.js'
import { mergeText, parseMerge } from '../layout/merges.js'
import { toLocal } from '../locale/index.js'
import { DEFAULT_LOCALE, LOCALES } from '../locale/registry.js'
import { builtinIn } from '../locale/rules.js'
import { formatAddress, MAX_COLS } from '../model.js'
import { columnPoints, widthOf } from '../numfmt.js'
import { cellResult, spilled } from '../result.js'
import { cellLooks } from '../shown.js'
import { meets, type Rect } from '../structure/area.js'
import { tableAt } from '../tables/model.js'
import { lookShown, windowLooks } from '../cond/looks.js'
import { sheetExtent } from './extent.js'

/** The cells a window gives by default, as many as the window of a sheet a server sends. */
const MAX_CELLS = 20_000

/** Excel's width of a column without its own (8 digits and the padding) and height of a row without its own, in points (Calibri 11). */
const DEFAULT_WIDTH = widthOf(columnPoints(undefined))
const DEFAULT_HEIGHT = 15

export function checkViewportOptions(v: unknown, path = 'opts'): ViewportOptions {
  if (v === undefined) return {}
  const o = object(v, path, 'options such as { formulaLocal: true, maxCells: 20000 }')
  onlyKeys(o, path, ['formulaLocal', 'diagnostics', 'maxCells'])
  if (o.formulaLocal !== undefined) checkBoolean(o.formulaLocal, `${path}.formulaLocal`)
  if (o.diagnostics !== undefined) checkBoolean(o.diagnostics, `${path}.diagnostics`)
  if (o.maxCells !== undefined) checkNumber(o.maxCells, `${path}.maxCells`, 1, MAX_COLS * 1_048_576, true)
  return o as ViewportOptions
}

/** The part of `x` a window of `max` cells gives: its first rows that fit, or the first cells of its first row. */
function fitted(x: Rect, max: number): Rect {
  const width = x.c2 - x.c1 + 1
  if (width * (x.r2 - x.r1 + 1) <= max) return x
  return width <= max ? { ...x, r2: x.r1 + Math.floor(max / width) - 1 } : { ...x, r2: x.r1, c2: x.c1 + max - 1 }
}

/** The window `asked` of sheet `s` (options checked); `display`, the session's: the cells' text and colour. The steps yield between cells. */
export function* viewport(state: EngineState, s: number, asked: Rect, o: ViewportOptions, display: boolean): Steps<Omit<Viewport, 'revision'>> {
  const x = fitted(asked, o.maxCells ?? MAX_CELLS)
  const sheet = state.input.sheets[s]
  const lines = sheetLines(sheet, state)
  const look = display ? cellLooks(state) : null
  const spill = yield* spilled(state, s, x, true)
  const cond = yield* windowLooks(state, s, x)
  // a style used in the window once, as one object of the answer's own
  const copies = new Map<number, CellStyle>()
  const style = (i: number) => {
    let c = copies.get(i)
    if (!c) copies.set(i, (c = JSON.parse(JSON.stringify(state.input.styles![i])) as CellStyle))
    return c
  }
  const builtin = builtinIn(state.input.locale)
  // formulas in the book's locale, each text once per table it stands in (a table's column is one formula)
  const loc = o.formulaLocal ? LOCALES.get(state.input.locale ?? DEFAULT_LOCALE)! : null
  const locals = new Map<string, string>()
  const local = (formula: string, r: number, c: number) => {
    const table = tableAt(state.tables, s, r, c)?.name
    const k = `${table ?? ''}\n${formula}`
    let out = locals.get(k)
    if (out === undefined) {
      const t = toLocal(formula, loc!, table === undefined ? {} : { table })
      locals.set(k, (out = t.ok ? t.formula : formula))
    }
    return out
  }
  // the cells errors began in, shared by the cells on one way to them
  const origins = new Map<string, Place>()
  const cells: Record<string, ViewportCell> = {}
  for (let r = x.r1; r <= x.r2; r++) {
    const row = lines.row(r)
    for (let c = x.c1; c <= x.c2; c++) {
      if (paced(1)) yield
      const addr = formatAddress(r, c)
      const own = sheet.cells[addr]
      const from = spill.get(addr)
      const cf = cond.get(r * MAX_COLS + c)
      const col = lines.column(c)
      const shown = lookOf(own, row, col)
      // an empty cell is listed when it shows something: a style of its row or column, a look, a note or a link
      if (!own && from === undefined && !cf && shown.style === undefined) continue
      const out = cellResult(state, s, addr, look, from, false) as ViewportCell
      if (loc && out.formula !== undefined) out.formulaLocal = local(out.formula, r, c)
      if (own?.array !== undefined) out.array = own.array
      if (own?.style !== undefined) out.style = style(own.style)
      if (shown.style !== undefined) out.effectiveStyle = style(shown.style)
      if (shown.format !== undefined) {
        out.numberFormat = typeof shown.format === 'number' ? builtin(shown.format) : shown.format
        out.numberFormatSource = formatSource(own, row, col)
      }
      if (cf) {
        out.appearance = cf
        lookShown(out, cf, look && ((format) => look(s, r, c, out.value, { ...own, format })))
      }
      if (own?.note !== undefined) out.note = own.note
      if (own?.link !== undefined) out.link = own.link
      if (o.diagnostics && out.value !== null && typeof out.value === 'object') out.issue = issueDetail(state, s, r, c, origins)!
      cells[addr] = out
    }
  }
  const columns: Viewport['columns'] = []
  for (let c = x.c1; c <= x.c2; c++) {
    const col = lines.column(c)
    columns.push(col?.hidden === true || col?.width === 0 ? { width: col.width ?? DEFAULT_WIDTH, hidden: true } : { width: col?.width ?? DEFAULT_WIDTH })
  }
  const rows: Viewport['rows'] = []
  for (let r = x.r1; r <= x.r2; r++) {
    const row = lines.row(r)
    rows.push(row && isHiddenRow(row) ? { height: row.height ?? DEFAULT_HEIGHT, hidden: true } : { height: row?.height ?? DEFAULT_HEIGHT })
  }
  const single = x.r1 === x.r2 && x.c1 === x.c2
  return {
    range: single ? formatAddress(x.r1, x.c1) : mergeText(x),
    truncated: x !== asked,
    cells,
    columns,
    rows,
    merges: (sheet.merges ?? []).filter((m) => meets(parseMerge(m)!, x)),
    usedRange: sheetExtent(state, s),
  }
}
