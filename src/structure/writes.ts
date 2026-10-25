/**
 * The cells a copy or a move writes with a value or a formula, held to `limits.writes` (`docs/async.md`) before the
 * operation builds anything: a move writes the cells of its source that hold one, a copy those of its source (a filtered
 * sheet's visible rows, as their values) once for each tile it lays, each run of visible rows of a filtered sheet's
 * target taking tiles of its own. The cells an operation clears, moves aside or rewrites are cells the sheets hold already.
 */
import type { BookInput, SheetInput, StructureOp } from '../contract.js'
import { holds } from '../array-areas.js'
import { allotWrites } from '../exec/budget.js'
import { cellCount, contains, hasCell, intersect, parseRange, type Rect } from './area.js'
import { clipped } from './clip.js'
import { tileCount } from './copy.js'
import { pasteRuns } from './filtered.js'
import { arraysOf, cellsIn, gridOf } from './grid.js'
import { sheetIndex } from './rewrite.js'
import type { ValuesIn } from './values.js'

/** A copy's or a move's sheets (-1 for a name the book has not) and ranges, read as given before the operation is checked. */
interface PasteOp {
  copy: boolean
  fs: number
  src: Rect
  ts: number
  /** A copy's target, a move's cell; null when it is no range. */
  at: Rect | null
}

/** The sheets and ranges of a copy or a move, or null for another operation or a source that is no range. */
export function pasteOf(book: BookInput, op: StructureOp): PasteOp | null {
  if (typeof op !== 'object' || op === null || (op.op !== 'copyRange' && op.op !== 'moveRange')) return null
  const o = op as unknown as Record<string, Record<string, unknown> | undefined>
  const index = (name: unknown) => (typeof name === 'string' ? sheetIndex(book, name) : -1)
  const range = (v: unknown) => (typeof v === 'string' ? parseRange(v) : null)
  const copy = op.op === 'copyRange'
  const src = range(o.from?.range)
  return src && { copy, fs: index(o.from?.sheet), src, ts: index(o.to?.sheet), at: range(copy ? o.to?.range : o.to?.cell) }
}

/** The cells of `x` a copy or a move writes with a value or a formula: those that hold one, and each cell of an array formula's area it takes part of (`copiedCells`). */
function contentIn(sheet: SheetInput, x: Rect): number {
  const parts = arraysOf(sheet).flatMap(([, a]) => {
    const part = contains(x, a) ? null : intersect(a, x)
    return part ? [part] : []
  })
  const g = gridOf(sheet)
  let n = parts.reduce((k, p) => k + cellCount(p), 0)
  for (const i of cellsIn(g, x)) if (holds(g.cell[i]) && !parts.some((p) => hasCell(p, g.row[i], g.col[i]))) n++
  return n
}

/** The cells a copy writes: its source's cells with content, from the clip of a filtered sheet's visible rows, times its tiles; a fill handle's copy (`whole`) the source's own over every row. */
function copyWrites(book: BookInput, p: PasteOp, target: Rect, values: ValuesIn, whole: boolean): number {
  const clip = whole ? null : clipped(book, book.sheets[p.fs].name, p.src, values)
  const [sheet, src] = clip ? [clip.book.sheets.at(-1)!, clip.src] : [book.sheets[p.fs], p.src]
  const runs = whole ? null : pasteRuns(book, book.sheets[p.ts].name, target)
  const laid = runs ? runs.reduce((n, [r1, r2]) => n + tileCount(src, { ...target, r1, r2 }), 0) : tileCount(src, target)
  return contentIn(sheet, src) * laid
}

/** A fill handle's copies (`pieces`, each a source and its target on sheet `s`) over `limits.writes` all together stop before the first. */
export function checkFillWrites(book: BookInput, s: number, pieces: [Rect, Rect][]): void {
  allotWrites(() => pieces.reduce((n, [from, at]) => n + contentIn(book.sheets[s], from) * tileCount(from, at), 0), book.locale)
}

/** A copy or a move over `limits.writes` stops before anything is built; the cells are counted only under the limit. */
export function checkWrites(book: BookInput, op: StructureOp, values: ValuesIn, whole = false): void {
  const p = pasteOf(book, op)
  const target = p?.at
  if (!p || !target || p.fs < 0 || p.ts < 0) return
  allotWrites(() => (p.copy ? copyWrites(book, p, target, values, whole) : contentIn(book.sheets[p.fs], p.src)), book.locale)
}
