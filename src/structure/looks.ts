/**
 * A sheet's looks (`SheetInput.looks`) under a structure operation. Lines inserted and deleted move the layer by its
 * blocks and patterns (`looks/move.ts`), the new lines taking the look of the line before them as the operation gives
 * their entries; the cells an operation copies, moves or writes otherwise (a range, a table's cells) are given to it as
 * entries of a look alone, which the engine takes back into the looks.
 */
import type { BookInput, StructureOp } from '../contract.js'
import { shifted } from '../looks/move.js'
import { expanded, hasLooks, layerOf, withLayer } from '../looks/sheet.js'
import { sheetKey } from '../syntax.js'
import { tableKey, tablesOf } from '../tables/model.js'
import { parseRange, shiftRect, type Rect } from './area.js'
import { tiles } from './copy.js'
import type { Applied } from './geometry.js'
import { pasteOf } from './writes.js'

/** The rectangles of a sheet whose looks an operation reads, moves or writes as entries, by sheet: its tables, a range's source and target. */
function looksRegions(book: BookInput, op: StructureOp): Map<number, Rect[]> {
  const out = new Map<number, Rect[]>()
  if (!book.sheets.some(hasLooks) || typeof op !== 'object' || op === null) return out
  const index = (name: unknown) => (typeof name === 'string' ? book.sheets.findIndex((s) => sheetKey(s.name) === sheetKey(name)) : -1)
  const add = (s: number, x: Rect | null) => {
    if (s >= 0 && x && hasLooks(book.sheets[s])) out.set(s, [...(out.get(s) ?? []), x])
  }
  const tables = (s: number) => {
    for (const t of tablesOf(book).list) if (t.s === s) add(s, t)
  }
  const range = (v: unknown) => (typeof v === 'string' ? parseRange(v) : null)
  const o = op as unknown as Record<string, unknown>
  switch (op.op) {
    case 'insertRows':
    case 'deleteRows':
    case 'insertColumns':
    case 'deleteColumns':
      // a table's cells, whose header and calculated columns the operation writes over their looks
      tables(index(o.sheet))
      break
    case 'moveRange':
    case 'copyRange': {
      const p = pasteOf(book, op)
      if (!p) break
      const { src, at } = p
      add(p.fs, src)
      if (at) add(p.ts, p.copy ? tiles(src, at).area : shiftRect(src, at.r1 - src.r1, at.c1 - src.c1))
      tables(p.fs)
      tables(p.ts)
      break
    }
    case 'addTable':
      add(index(o.sheet), range((o.table as Record<string, unknown> | undefined)?.range))
      break
    case 'deleteTable': {
      const t = tablesOf(book).list.find((x) => typeof o.table === 'string' && tableKey(x.name) === tableKey(o.table))
      if (t) add(t.s, t)
      break
    }
  }
  return out
}

/** The book with the looks an operation copies, moves or writes as entries of a look alone (`looksRegions`). */
export function overLooks(book: BookInput, op: StructureOp): BookInput {
  const regions = looksRegions(book, op)
  return regions.size ? { ...book, sheets: book.sheets.map((s, i) => (regions.has(i) ? expanded(s, regions.get(i)!) : s)) } : book
}

/**
 * The book an operation gave (`r`, from `input`) with its sheets' looks moved by the lines it inserted and deleted: each
 * new line with the look of the line before it, as the operation gives its entries (`formatBefore`).
 */
export function linesMoved(input: BookInput, r: Extract<Applied, { ok: true }>): Applied {
  const lines = r.geometry.flatMap((m) => (m.kind === 'lines' ? [m] : []))
  const sheets = r.book.sheets.slice()
  let moved = false
  for (const s of new Set(lines.map((m) => m.s))) {
    let layer = layerOf(input.sheets[s])
    if (!layer) continue
    for (const m of lines) if (m.s === s) layer = shifted(layer, m.axis, m.at, m.n, m.n > 0)
    sheets[s] = withLayer(sheets[s], sheets[s].cells, layer)
    moved = true
  }
  return moved ? { ...r, book: { ...r.book, sheets } } : r
}
