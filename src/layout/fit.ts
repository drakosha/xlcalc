/**
 * Column widths as Excel for Mac fits them (`docs/formats.md`, "Fitting columns"). AutoFit gives a column the width of
 * its widest cell: each value by its format (its own, else its row's, else its column's) in the book's locale, laid out
 * with the room Excel gives it (`fitPoints`); cells of hidden rows count, cells merged across columns do not, and a
 * column with nothing to show keeps its width. A number typed into a column whose width nobody set widens the column.
 */
import type { BookInput, CellValue, ColumnInput, ColumnsFit, ComputeResult, EntryWidthOptions, LayoutEdit, LinePatch, NumberFormat, SheetInput } from '../contract.js'
import { epochOf } from '../dates.js'
import { toCellValue, type EngineState } from '../engine.js'
import { checkBook, checkCellValue, checkEntryWidthOptions, checkSheet, InputError } from '../input.js'
import { displayRules, type Rules } from '../locale/rules.js'
import { formatAddress, parseAddress } from '../model.js'
import { columnPoints, widthOf } from '../numfmt.js'
import { outValue } from '../result.js'
import { applyLayout } from '../session/book.js'
import { sheetKey } from '../syntax.js'
import { parseMerge } from './merges.js'
import { lines, parseSpan, spanKey, type Line } from './spans.js'

/** A cell of a sheet and its value: zero-based row and column. */
export type Placed = readonly [r: number, c: number, v: CellValue]

/** The widest column of the model (`ColumnInput.width`); Excel's widest is 255.83, 255 digits and the padding. */
const MAX_WIDTH = 255

/** The entry of line `i` among sorted, disjoint lines, by bisection. */
function lineOf<T>(ls: Line<T>[], i: number): T | undefined {
  let [a, b] = [0, ls.length - 1]
  while (a <= b) {
    const m = (a + b) >> 1
    if (ls[m][1] < i) a = m + 1
    else if (ls[m][0] > i) b = m - 1
    else return ls[m][2]
  }
  return undefined
}

/** The width of a column of `points`: null for the default width, which Excel writes as no width at all. */
const widthAt = (points: number): number | null => (points === columnPoints(undefined) ? null : Math.min(MAX_WIDTH, widthOf(points)))

/**
 * The widths AutoFit gives the columns `lo` to `hi` of a sheet whose cells hold `cells` (in any order), as
 * `ColumnInput.width`, null for the default width; a column with nothing to show is left out, as it keeps its width.
 */
export function fitWidths(rules: Rules, sheet: SheetInput, cells: Iterable<Placed>, lo: number, hi: number): Record<string, number | null> {
  const rows = lines(sheet.rows, 'r')
  const cols = lines(sheet.columns, 'c')
  const across = (sheet.merges ?? []).map((m) => parseMerge(m)!).filter((x) => x.c2 > x.c1 && x.c2 >= lo && x.c1 <= hi)
  const widest = new Map<number, number>()
  for (const [r, c, v] of cells) {
    if (c < lo || c > hi || v === null || across.some((x) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2)) continue
    const format: NumberFormat | undefined = sheet.cells[formatAddress(r, c)]?.format ?? (rows.length ? lineOf(rows, r)?.format : undefined) ?? (cols.length ? lineOf(cols, c)?.format : undefined)
    const points = rules.fitPoints(v, format)
    if (points !== null && points > (widest.get(c) ?? 0)) widest.set(c, points)
  }
  return Object.fromEntries([...widest].sort((a, b) => a[0] - b[0]).map(([c, points]) => [spanKey(c, c, 'c'), widthAt(points)]))
}

/** What AutoFit sets on a column it fits: the width as Excel's own, the column shown; the default width bare. */
const fitted = (width: number | null): LinePatch<ColumnInput> => ({ width, bestFit: width === null ? null : true, hidden: null })

/** The values of columns `lo` to `hi` of a sheet of the state, spilled ones too; the unstored cells of a large spill once per column. */
function* stateCells(s: EngineState, i: number, lo: number, hi: number): Generator<Placed> {
  const data = s.ctx.sheets[i]
  for (let c = lo; c <= Math.min(hi, data.cols.length - 1); c++) {
    const col = data.cols[c]
    if (col) for (let r = 0; r < col.length; r++) if (col[r] !== undefined && col[r] !== null) yield [r, c, outValue(toCellValue(col[r]))]
  }
  for (const f of data.fills) for (let c = Math.max(lo, f.c1); c <= Math.min(hi, f.c2); c++) yield [f.r1, c, outValue(toCellValue(f.v))]
}

/**
 * Layout edits with AutoFit (`ColumnsFit`) turned into the widths it sets, each from the state's values and the layout
 * the edits before it leave.
 */
export function fitEdits(s: EngineState, edits: LayoutEdit[]): Exclude<LayoutEdit, ColumnsFit>[] {
  if (!edits.some((e) => 'fit' in e)) return edits as Exclude<LayoutEdit, ColumnsFit>[]
  const scratch: BookInput = { ...s.input, sheets: s.input.sheets.slice() }
  const index = (name: string) => s.sheetIdx.get(sheetKey(name))!
  const out: Exclude<LayoutEdit, ColumnsFit>[] = []
  for (const e of edits) {
    const set: Exclude<LayoutEdit, ColumnsFit>[] = []
    if ('fit' in e) {
      const i = index(e.sheet)
      const [lo, hi] = parseSpan(e.columns, 'c')!
      for (const [col, width] of Object.entries(fitWidths(s.ctx.rules, scratch.sheets[i], stateCells(s, i, lo, hi), lo, hi))) set.push({ sheet: e.sheet, columns: col, set: fitted(width) })
    } else set.push(e)
    applyLayout(scratch, set, index)
    out.push(...set)
  }
  return out
}

/**
 * The widths AutoFit gives columns `C` or `B:D` of the sheet `sheet` of a computed book, from the values of its result
 * (`compute(book)`), by column letter: null for the default width; a column with nothing to show is left out.
 */
export function fitColumns(book: BookInput, result: ComputeResult, sheet: string, columns: string): Record<string, number | null> {
  checkBook(book)
  const index = (name: string) => {
    const i = book.sheets.findIndex((x) => sheetKey(x.name) === sheetKey(name))
    return i < 0 ? undefined : i
  }
  const i = checkSheet(sheet, index, 'sheet')
  const span = typeof columns === 'string' ? parseSpan(columns, 'c') : null
  if (!span) throw new InputError('columns', `expected columns such as "C" or "B:D", got ${JSON.stringify(columns)}`)
  const values = result.sheets.find((x) => sheetKey(x.name) === sheetKey(book.sheets[i].name))
  if (!values) throw new InputError('result', `the result has no sheet ${JSON.stringify(book.sheets[i].name)}`)
  const cells = Object.entries(values.cells).map(([a, x]): Placed => {
    const { row, col } = parseAddress(a)
    return [row, col, x.value]
  })
  return fitWidths(displayRules(book.locale, epochOf(book.dateSystem)), book.sheets[i], cells, ...span)
}

/**
 * The width Excel gives a column as `value`, shown by `opts.format`, is typed into one of its cells (`test/excel/fit/`):
 * a number AutoFit would give more room widens a column whose width is the default or one Excel fitted (`bestFit`), as
 * AutoFit fits it; undefined when the column keeps its width (text overflows, General drops decimals, a width set by
 * hand stays, a column never narrows), null for the default width.
 */
export function entryWidth(value: CellValue, opts?: EntryWidthOptions): number | null | undefined {
  const o = checkEntryWidthOptions(opts)
  checkCellValue(value, 'value')
  if (o.width !== undefined && !o.bestFit) return undefined
  const points = displayRules(o.locale, epochOf(o.dateSystem)).typedPoints(value, o.format, columnPoints(o.width))
  return points === null ? undefined : widthAt(points)
}
