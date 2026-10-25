/**
 * Column widths as Excel for Mac fits them (`docs/formats.md`, "Fitting columns"). AutoFit gives a column the width of
 * its widest cell: each value by its format (its own, else its row's, else its column's) in the book's locale, laid out
 * with the room Excel gives it (`fitPoints`); cells of hidden rows count, cells merged across columns do not, and a
 * column with nothing to show keeps its width. A number typed into a column whose width nobody set widens the column.
 */
import type { BookInput, CellValue, ColumnInput, ComputeResult, EntryWidthOptions, LayoutEdit, LinePatch, RowInput, SheetInput } from '../contract.js'
import { epochOf } from '../dates.js'
import type { EngineState } from '../engine.js'
import { checkCellValue, checkEntryWidthOptions } from '../input.js'
import { displayRules, type Rules } from '../locale/rules.js'
import { columnPoints, widthOf } from '../numfmt.js'
import { outCell } from '../result.js'
import { applyLayout, type SetLayoutEdit } from '../session/book.js'
import { sheetKey } from '../syntax.js'
import { sheetLines } from './lines.js'
import { parseMerge } from './merges.js'
import { hasCell } from '../structure/area.js'
import { parseSpan, spanKey } from './spans.js'
import { fitHeights, stateRows } from './heights.js'
import { resultSheet, type Placed } from './result-sheet.js'

/** The widest column of the model (`ColumnInput.width`); Excel's widest is 255.83, 255 digits and the padding. */
const MAX_WIDTH = 255

/** The width of a column of `points`: null for the default width, which Excel writes as no width at all. */
const widthAt = (points: number): number | null => (points === columnPoints(undefined) ? null : Math.min(MAX_WIDTH, widthOf(points)))

/**
 * The widths AutoFit gives the columns `lo` to `hi` of a sheet whose cells hold `cells` (in any order), as
 * `ColumnInput.width`, null for the default width; a column with nothing to show is left out, as it keeps its width.
 */
export function fitWidths(rules: Rules, sheet: SheetInput, cells: Iterable<Placed>, lo: number, hi: number): Record<string, number | null> {
  const formats = sheetLines(sheet)
  const across = (sheet.merges ?? []).map((m) => parseMerge(m)!).filter((x) => x.c2 > x.c1 && x.c2 >= lo && x.c1 <= hi)
  const widest = new Map<number, number>()
  for (const [r, c, v] of cells) {
    if (c < lo || c > hi || v === null || across.some((x) => hasCell(x, r, c))) continue
    const points = rules.fitPoints(v, formats.format(r, c))
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
    if (col)
      for (let r = 0; r < col.length; r++) {
        const v = col.get(r)
        if (v !== undefined && v !== null) yield [r, c, outCell(v)]
      }
  }
  for (const f of data.fills) for (let c = Math.max(lo, f.c1); c <= Math.min(hi, f.c2); c++) yield [f.r1, c, outCell(f.v)]
}

/** What AutoFit sets on a row it fits: the height as Excel's own, the row shown; the default height bare. */
const rowFitted = (height: number | null): LinePatch<RowInput> => ({ height, bestFit: height === null ? null : true, hidden: null })

/**
 * Layout edits with AutoFit (`ColumnsFit`, `RowsFit`) turned into the widths and heights it sets, each from the state's
 * values and the layout the edits before it leave.
 */
export function fitEdits(s: EngineState, edits: LayoutEdit[]): SetLayoutEdit[] {
  if (!edits.some((e) => 'fit' in e)) return edits as SetLayoutEdit[]
  const scratch: BookInput = { ...s.input, sheets: s.input.sheets.slice() }
  const index = (name: string) => s.sheetIdx.get(sheetKey(name))!
  const out: SetLayoutEdit[] = []
  for (const e of edits) {
    const set: SetLayoutEdit[] = []
    if ('fit' in e && 'columns' in e) {
      const i = index(e.sheet)
      const [lo, hi] = parseSpan(e.columns, 'c')!
      for (const [col, width] of Object.entries(fitWidths(s.ctx.rules, scratch.sheets[i], stateCells(s, i, lo, hi), lo, hi))) set.push({ sheet: e.sheet, columns: col, set: fitted(width) })
    } else if ('fit' in e) {
      const i = index(e.sheet)
      const [lo, hi] = parseSpan(e.rows, 'r')!
      const { cells, spans } = stateRows(s, i, lo, hi)
      for (const [rows, height] of Object.entries(fitHeights(scratch.styles ?? [], scratch.sheets[i], cells, lo, hi, spans))) set.push({ sheet: e.sheet, rows, set: rowFitted(height) })
    } else set.push(e)
    applyLayout(scratch, set, index)
    out.push(...set)
  }
  return out
}

/**
 * The widths AutoFit gives columns `C` or `B:D` of the sheet `sheet` of a calculated workbook, from the values of its
 * result (`compute(book)`), by column letter: null for the default width; a column with nothing to display is left out.
 */
export function fitColumns(book: BookInput, result: ComputeResult, sheet: string, columns: string): Record<string, number | null> {
  const x = resultSheet(book, result, sheet, columns, 'c')
  return fitWidths(displayRules(x.book.locale, epochOf(x.book.dateSystem)), x.book.sheets[x.i], x.cells, x.lo, x.hi)
}

/**
 * The width Excel gives a column when `value`, displayed in `opts.numberFormat`, is typed into one of its cells
 * (`test/excel/fit/`): a number for which AutoFit would make more room widens a column whose width is the default or
 * was fitted by Excel (`bestFit`), as AutoFit fits it, even to the default width (Excel keeps a fitted width that
 * equals it); undefined when the column keeps its width (text overflows, General drops decimals, a manually set width
 * stays, a column never narrows).
 */
export function entryWidth(value: CellValue, opts?: EntryWidthOptions): number | undefined {
  const o = checkEntryWidthOptions(opts)
  checkCellValue(value, 'value')
  if (o.width !== undefined && !o.bestFit) return undefined
  const points = displayRules(o.locale, epochOf(o.dateSystem)).typedPoints(value, o.numberFormat, columnPoints(o.width))
  return points === null ? undefined : Math.min(MAX_WIDTH, widthOf(points))
}
