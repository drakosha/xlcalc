/**
 * The layout of a sheet under inserted and deleted rows and columns, as Excel changes it (checked in
 * `test/excel/struct/styles.json`): columns and rows keep their entries as they move, new lines take the
 * entry and the cells' style and format of the line before them, merged areas grow, shrink and move as
 * ranges do and are gone when one cell or none is left.
 */
import type { CellInput, SheetInput } from '../contract.js'
import { spend } from '../exec/budget.js'
import { lines, toMap, type Line } from '../layout/spans.js'
import { mergeText, parseMerge } from '../layout/merges.js'
import { formatAddress, parseAddress } from '../model.js'
import { LIMIT, type Axis, type Rect } from './area.js'

/** Index `i` after `n` lines inserted before `at` (n > 0) or deleted from `at` (n < 0); `gone` for a deleted one. */
function moved(i: number, at: number, n: number, gone: number): number {
  if (i < at) return i
  return n > 0 ? i + n : i >= at - n ? i + n : gone
}

/**
 * Columns or rows after the lines move: a span holding the line before `at` takes the inserted lines (they get
 * its entry, shown even when it is hidden: the struct set `hidden`), a span past the sheet's end is cut there;
 * deleted lines leave their spans. A span of columns that
 * reaches XFD keeps reaching it when columns are deleted, unless all of it is, as xlsx keeps it (`<col max=16384>`);
 * where the deleted columns include XFD, the columns that come in at the end get no width (one observation). Rows,
 * kept one by one, do not reach on (the struct set `edge-layout`).
 */
export function shiftLines<T extends object>(map: Record<string, T> | undefined, axis: Axis, at: number, n: number): Record<string, T> | undefined {
  if (!map) return map
  const out: Line<T>[] = []
  const last = LIMIT[axis] - 1
  for (const [lo, hi, v] of lines(map, axis)) {
    if (n < 0 && lo >= at && hi < at - n) continue
    const a = moved(lo, at, n, at)
    if (n < 0 && axis === 'c' && hi === last) {
      const full = at - n > last ? last + n : last
      if (a <= full) out.push([a, full, v])
      const rest = Object.fromEntries(Object.entries(v).filter(([k]) => k !== 'width')) as T
      if (full < last && Object.keys(rest).length) out.push([Math.max(a, full + 1), last, rest])
      continue
    }
    const b = n > 0 && hi === at - 1 ? hi + n : moved(hi, at, n, at - 1)
    if (a > b || a >= LIMIT[axis]) continue
    // the inserted lines take the entry of the line before them, but for `hidden`
    if (n > 0 && 'hidden' in v && lo < at && hi >= at - 1) {
      const { hidden: _, ...shown } = v
      out.push([a, at - 1, v], ...(Object.keys(shown).length ? [[at, Math.min(at + n - 1, last), shown as T] as Line<T>] : []))
      if (at + n <= b) out.push([at + n, Math.min(b, last), v])
    } else out.push([a, Math.min(b, last), v])
  }
  return toMap(out, axis)
}

/** The axis's bounds of a rectangle, and the rectangle with them replaced. */
const bounds = (x: Rect, axis: Axis): [number, number] => (axis === 'r' ? [x.r1, x.r2] : [x.c1, x.c2])
const withBounds = (x: Rect, axis: Axis, lo: number, hi: number): Rect => (axis === 'r' ? { ...x, r1: lo, r2: hi } : { ...x, c1: lo, c2: hi })

/**
 * Merged areas after the lines move, as ranges: one grows with lines inserted inside it and loses deleted
 * ones; it is gone when one cell or none is left; null when an inserted line pushes one past the sheet's end.
 * Deleting the sheet's last lines leaves an area that ends there as it is, unless all of its lines go (the struct
 * set `edge-layout`: deleting row 1,048,576 keeps `C1048575:D1048576`).
 */
export function shiftMerges(merges: string[] | undefined, axis: Axis, at: number, n: number): string[] | undefined | null {
  if (!merges) return merges
  const out: string[] = []
  for (const m of merges) {
    const x = parseMerge(m)!
    const [lo, hi] = bounds(x, axis)
    if (n < 0 && lo >= at && hi < at - n) continue
    const a = moved(lo, at, n, at)
    const b = n < 0 && hi === LIMIT[axis] - 1 && at - n === LIMIT[axis] ? hi : moved(hi, at, n, at - 1)
    if (b >= LIMIT[axis]) return null
    const y = withBounds(x, axis, a, b)
    if (a <= b && (y.r1 < y.r2 || y.c1 < y.c2)) out.push(mergeText(y))
  }
  return out
}

/**
 * The cells that inserted lines get from the line before them, `at - 1`: the style and format of each of its
 * cells that has one, no value; nothing before the first line.
 */
export function formatBefore(sheet: SheetInput, axis: Axis, at: number, n: number): Record<string, CellInput> {
  const out: Record<string, CellInput> = {}
  if (at === 0) return out
  for (const [addr, cell] of Object.entries(sheet.cells)) {
    spend(1)
    const p = parseAddress(addr)
    if ((axis === 'r' ? p.row : p.col) !== at - 1 || (cell.style === undefined && cell.format === undefined)) continue
    const copy: CellInput = { ...(cell.format !== undefined ? { format: cell.format } : {}), ...(cell.style !== undefined ? { style: cell.style } : {}) }
    for (let i = at; i < at + n; i++) out[axis === 'r' ? formatAddress(i, p.col) : formatAddress(p.row, i)] = copy
  }
  return out
}
