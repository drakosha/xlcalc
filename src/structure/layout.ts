/**
 * The layout of a sheet under inserted and deleted rows and columns, as Excel changes it (checked in
 * `test/excel/struct/styles.json`): columns and rows keep their entries as they move, new lines take the
 * entry and the cells' style and format of the line before them (`lookBefore`), merged areas grow, shrink and move as
 * ranges do and are gone when one cell or none is left, frozen panes follow the lines above their split.
 */
import type { FrozenPanes, SheetInput } from '../contract.js'
import { entryAt, lines, patchLines, toMap, type Line } from '../layout/spans.js'
import { parseMerge } from '../layout/merges.js'
import type { LookTable, RowsEdit } from '../looks/layer.js'
import { hasLook } from '../looks/sheet.js'
import { cornersText, hi, LIMIT, LINES, lo, withBounds, type Axis } from './area.js'
import { gridOf } from './grid.js'

/** Index `i` after `n` lines inserted before `at` (n > 0) or deleted from `at` (n < 0); `gone` for a deleted one. */
function moved(i: number, at: number, n: number, gone: number): number {
  if (i < at) return i
  return n > 0 ? i + n : i >= at - n ? i + n : gone
}

/**
 * Columns or rows after the lines move: a span holding the line before `at` takes the inserted lines (they get
 * its entry, shown even when it is hidden: the struct set `hidden`, and not collapsing a group), a span past the sheet's end is cut there;
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
    // the inserted lines take the entry of the line before them (its outline level too), but for `hidden` and `collapsed`
    if (n > 0 && ('hidden' in v || 'collapsed' in v) && lo < at && hi >= at - 1) {
      const { hidden: _, collapsed: __, ...shown } = v as T & { hidden?: boolean; collapsed?: boolean }
      out.push([a, at - 1, v], ...(Object.keys(shown).length ? [[at, Math.min(at + n - 1, last), shown as T] as Line<T>] : []))
      if (at + n <= b) out.push([at + n, Math.min(b, last), v])
    } else out.push([a, Math.min(b, last), v])
  }
  return toMap(out, axis)
}

type Outlined = { outlineLevel?: number; hidden?: boolean; collapsed?: boolean }

/**
 * Lines inserted before `at` into a collapsed group (after a hidden line of it, whose level they take) are shown, so the
 * group is collapsed no more: its summary line, the first line past the group (`after`), else the line before it, loses its
 * mark (the valid set `made`).
 */
export function openedGroup<T extends Outlined>(map: Record<string, T> | undefined, axis: Axis, at: number, n: number, after: boolean): Record<string, T> | undefined {
  const ls = lines(map, axis)
  const before = at > 0 ? entryAt(ls, at - 1) : undefined
  const level = before?.outlineLevel ?? 0
  if (n <= 0 || !level || !before?.hidden) return map
  const deep = (i: number) => (entryAt(ls, i)?.outlineLevel ?? 0) >= level
  let k = after ? at + n : at - 1
  while (k >= 0 && k < LIMIT[axis] && deep(k)) k += after ? 1 : -1
  return k >= 0 && entryAt(ls, k)?.collapsed ? (patchLines(map, axis, k, k, { collapsed: null } as never) as Record<string, T>) : map
}

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
    const [i, j] = [lo(x, axis), hi(x, axis)]
    if (n < 0 && i >= at && j < at - n) continue
    const a = moved(i, at, n, at)
    const b = n < 0 && j === LIMIT[axis] - 1 && at - n === LIMIT[axis] ? j : moved(j, at, n, at - 1)
    if (b >= LIMIT[axis]) return null
    const y = withBounds(x, axis, a, b)
    if (a <= b && (y.r1 < y.r2 || y.c1 < y.c2)) out.push(cornersText(y))
  }
  return out
}

/**
 * Frozen panes after the lines move: their split is the edge between two lines, which lines inserted above it push
 * on and deleted ones above it draw back, but not past the first line: with every frozen line deleted the first one is
 * frozen (the parts sets `struct-held` and `frozen-held`).
 */
export function shiftFrozen(frozen: FrozenPanes | undefined, axis: Axis, at: number, n: number): FrozenPanes | undefined {
  const key = LINES[axis]
  const k = frozen?.[key]
  if (k === undefined || at >= k) return frozen
  const next = n > 0 ? Math.min(k + n, LIMIT[axis] - 1) : Math.max(1, k - (Math.min(at - n, k) - at))
  return { ...frozen, [key]: next }
}

/**
 * The looks that inserted lines get from the entries of the line before them, `at - 1`: the style and format of each of
 * its cells that has one, no value, as edits of the sheet's looks (`looks/layer.ts`) over the new lines, one rectangle per
 * cell however many lines come in; nothing before the first line. The looks of its cells without an entry come with the
 * sheet's looks (`looks/move.ts`).
 */
export function lookBefore(sheet: SheetInput, axis: Axis, at: number, n: number, looks: LookTable): RowsEdit[] {
  if (at === 0) return []
  const g = gridOf(sheet)
  const line = axis === 'r' ? g.row : g.col
  const found: [number, number][] = []
  for (let k = 0; k < line.length; k++) {
    const cell = g.cell[k]
    if (line[k] !== at - 1 || !hasLook(cell)) continue
    found.push([axis === 'r' ? g.col[k] : g.row[k], looks.id(cell.style, cell.numberFormat)])
  }
  found.sort((a, b) => a[0] - b[0])
  // a row's cells as runs of one edit over the new rows, a column's as one edit of each row over the new columns
  if (axis === 'r') return found.length ? [{ lo: at, hi: at + n - 1, top: found.flatMap(([c, id]) => [c, c, id]) }] : []
  return found.map(([r, id]) => ({ lo: r, hi: r, top: [at, at + n - 1, id] }))
}
