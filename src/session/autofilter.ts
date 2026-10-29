/**
 * The AutoFilter grows as Excel grows it when cells right under it are filled (docs/structure.md "The AutoFilter"): a
 * value or a formula written into the row under its range, in one of its columns, takes that row into the range, and
 * so on down while the rows filled follow one another. The session's calls grow it, as Excel grows it for what is
 * entered on the sheet; the functions over a book leave it where it is.
 */
import type { BookInput, CellInput, CellReplace, StructureOp } from '../contract.js'
import { formatAddress, MAX_ROWS, parseAddress } from '../model.js'
import { parseRange, rectText, type Rect } from '../structure/area.js'
import type { SetLayoutEdit } from './book.js'

/** Whether a cell's input holds something: a value or a formula. */
const holds = (x: CellInput | null | undefined) => !!x && (x.formula !== undefined || (x.value !== undefined && x.value !== null))

/** The range `x` grown down over the rows right under it that `filled` gives; null when the row under it is not one. */
function grownDown(x: Rect, filled: (r: number) => boolean): Rect | null {
  let r2 = x.r2
  while (r2 + 1 < MAX_ROWS && filled(r2 + 1)) r2++
  return r2 === x.r2 ? null : { ...x, r2 }
}

/** The AutoFilters the edits of cells grow, as the layout edits that set them. */
export function filterGrowth(book: BookInput, edits: CellReplace[], sheetIndex: (name: string) => number): SetLayoutEdit[] {
  if (!book.sheets.some((x) => x.autoFilter)) return []
  // each sheet's AutoFilter and the rows under it filled in its columns
  const under = new Map<number, [Rect, Set<number>]>()
  for (const e of edits) {
    if (!holds(e.input)) continue
    const i = sheetIndex(e.sheet)
    const f = book.sheets[i].autoFilter
    if (!f) continue
    let at = under.get(i)
    if (!at) under.set(i, (at = [parseRange(f.range)!, new Set()]))
    const { row, col } = parseAddress(e.cell)
    if (row > at[0].r2 && col >= at[0].c1 && col <= at[0].c2) at[1].add(row)
  }
  const out: SetLayoutEdit[] = []
  for (const [i, [x, rows]] of under) {
    const y = rows.size ? grownDown(x, (r) => rows.has(r)) : null
    if (y) out.push({ sheet: book.sheets[i].name, autoFilter: { range: rectText(y) } })
  }
  return out
}

/** The cells a copy or a fill writes: a copy's target, at least as large as its source, and a fill's whole range. */
function writtenBy(op: Extract<StructureOp, { op: 'copyRange' | 'fillRange' }>): [string, Rect] {
  if (op.op === 'fillRange') return [op.sheet, parseRange(op.to)!]
  const [src, t] = [parseRange(op.from.range)!, parseRange(op.to.range)!]
  const [h, w] = op.transpose ? [src.c2 - src.c1, src.r2 - src.r1] : [src.r2 - src.r1, src.c2 - src.c1]
  return [op.to.sheet, { ...t, r2: Math.max(t.r2, t.r1 + h), c2: Math.max(t.c2, t.c1 + w) }]
}

/** The book after a copy or a fill (`op`), its target sheet's AutoFilter grown over the rows under it the operation filled. */
export function grownByOp(book: BookInput, op: Extract<StructureOp, { op: 'copyRange' | 'fillRange' }>, sheetIndex: (name: string) => number): BookInput {
  const [sheet, w] = writtenBy(op)
  const i = sheetIndex(sheet)
  const s = book.sheets[i]
  if (!s.autoFilter) return book
  const x = parseRange(s.autoFilter.range)!
  const [c1, c2] = [Math.max(x.c1, w.c1), Math.min(x.c2, w.c2)]
  const filled = (r: number) => {
    if (r < w.r1 || r > w.r2) return false
    for (let c = c1; c <= c2; c++) if (holds(s.cells[formatAddress(r, c)])) return true
    return false
  }
  const y = c1 <= c2 ? grownDown(x, filled) : null
  if (!y) return book
  const sheets = book.sheets.slice()
  sheets[i] = { ...s, autoFilter: { range: rectText(y) } }
  return { ...book, sheets }
}
