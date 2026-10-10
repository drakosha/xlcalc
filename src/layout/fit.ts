/**
 * Column widths as Excel for Mac fits them (`docs/formats.md`, "Fitting columns"). AutoFit gives a column the width of
 * its widest cell: each value by its format (its own, else its row's, else its column's) in the book's locale, laid out
 * with the room Excel gives it (`fitPoints`); cells of hidden rows count, cells merged across columns do not, and a
 * column with nothing to show keeps its width. A number typed into a column whose width nobody set widens the column.
 */
import type { CellValue, NumberFormat, SheetInput } from '../contract.js'
import type { Rules } from '../locale/rules.js'
import { formatAddress } from '../model.js'
import { columnPoints, widthOf } from '../numfmt.js'
import { parseMerge } from './merges.js'
import { lines, spanKey, type Line } from './spans.js'

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
export const widthAt = (points: number): number | null => (points === columnPoints(undefined) ? null : Math.min(MAX_WIDTH, widthOf(points)))

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

/** A column as entry finds it: its width (the default without one) and whether Excel fitted it rather than a hand set it. */
export interface ColumnState {
  width?: number
  fitted?: boolean
}

/**
 * The width Excel gives a column when `value`, shown by `format`, is typed into one of its cells
 * (`test/excel/fit/typed.json`): a number the column shows as `#`, or in General in a scientific form other than its
 * own, widens a column whose width is the default or one Excel fitted, as AutoFit fits the number; undefined when the
 * column keeps its width (text overflows, General drops decimals, a width set by hand stays, a column never narrows).
 */
export function typedWidth(rules: Rules, value: CellValue, format: NumberFormat | undefined, column: ColumnState = {}): number | null | undefined {
  if (column.width !== undefined && !column.fitted) return undefined
  const points = rules.typedPoints(value, format, columnPoints(column.width))
  return points === null ? undefined : widthAt(points)
}
