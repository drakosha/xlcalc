/**
 * Column widths as Excel for Mac fits them (`docs/formats.md`, "Fitting columns"). AutoFit gives a column the width of
 * the widest text its cells show with all the room they need: each value by its format (its own, else its row's, else
 * its column's) in the book's locale; cells of hidden rows and of merged areas are left out. A number typed into a
 * column whose width nobody set widens the column until the number shows whole.
 */
import type { CellValue, NumberFormat, SheetInput } from '../contract.js'
import type { Rules } from '../locale/rules.js'
import { formatAddress } from '../model.js'
import { columnPoints, room, widthFor } from '../numfmt.js'
import { isHiddenRow } from './hidden.js'
import { parseMerge } from './merges.js'
import { lines, spanKey, type Line } from './spans.js'

/** A cell of a sheet and its value: zero-based row and column. */
export type Placed = readonly [r: number, c: number, v: CellValue]

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

/**
 * The widths AutoFit gives columns `lo` to `hi` of a sheet whose cells hold `cells` (in any order), as
 * `ColumnInput.width`: a column with something to show gets the width of its widest text, a column with nothing
 * that has a width of its own gets null (the default width), and a column with neither is left out.
 */
export function fitWidths(rules: Rules, sheet: SheetInput, cells: Iterable<Placed>, lo: number, hi: number): Record<string, number | null> {
  const rows = lines(sheet.rows, 'r')
  const cols = lines(sheet.columns, 'c')
  const merges = (sheet.merges ?? []).map((m) => parseMerge(m)!).filter((x) => x.c2 >= lo && x.c1 <= hi)
  const widest = new Map<number, number>()
  for (const [r, c, v] of cells) {
    if (c < lo || c > hi || v === null) continue
    const row = rows.length ? lineOf(rows, r) : undefined
    if ((row && isHiddenRow(row)) || merges.some((x) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2)) continue
    const format: NumberFormat | undefined = sheet.cells[formatAddress(r, c)]?.format ?? row?.format ?? (cols.length ? lineOf(cols, c)?.format : undefined)
    const need = rules.need(v, format)
    if (need !== null && need > (widest.get(c) ?? -1)) widest.set(c, need)
  }
  const out: Record<string, number | null> = {}
  for (const [c, need] of [...widest].sort((a, b) => a[0] - b[0])) out[spanKey(c, c, 'c')] = widthFor(need)
  for (const [a, b, e] of cols)
    if (e.width !== undefined) for (let c = Math.max(a, lo); c <= Math.min(b, hi); c++) if (!widest.has(c)) out[spanKey(c, c, 'c')] = null
  return out
}

/** A column as entry finds it: its width (the default without one) and whether Excel fitted it rather than a hand set it. */
export interface ColumnState {
  width?: number
  fitted?: boolean
}

/**
 * The width Excel gives a column when `value`, shown by `format`, is typed into one of its cells: a number that does not
 * fit widens a column whose width is the default or one Excel fitted, to the width that shows it whole; undefined when
 * the column keeps its width (text overflows, a width set by hand stays, a column never narrows).
 */
export function typedWidth(rules: Rules, value: CellValue, format: NumberFormat | undefined, column: ColumnState = {}): number | undefined {
  if (typeof value !== 'number' || (column.width !== undefined && !column.fitted)) return undefined
  const need = rules.need(value, format)
  return need === null || need <= room(columnPoints(column.width)) ? undefined : widthFor(need)
}
