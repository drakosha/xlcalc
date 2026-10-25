/**
 * A sheet's looks (`SheetInput.looks`) under a structure operation. Lines inserted and deleted move the layer by its
 * blocks and patterns (`looks/move.ts`), the new lines taking the look of the line before them as the operation gives
 * their entries, and a copied or moved range's looks go with it as rectangles; the cells an operation writes otherwise (a
 * table's cells) are given to it as entries of a look alone, which the engine takes back into the looks.
 */
import type { BookInput, CellInput, StructureOp } from '../contract.js'
import { edited, emptyLayer, LookTable, type Layer } from '../looks/layer.js'
import { shifted } from '../looks/move.js'
import { expanded, hasLook, hasLooks, layerOf, withLayer } from '../looks/sheet.js'
import { sheetKey } from '../syntax.js'
import { tableKey, tablesOf } from '../tables/model.js'
import { readAddress } from '../model.js'
import { parseRange, type Axis, type Rect } from './area.js'
import type { Applied } from './geometry.js'
import { lookBefore } from './layout.js'
import { pasteOf } from './writes.js'

/**
 * The rectangles of a sheet whose looks an operation reads, moves or writes as entries, by sheet: its tables. A copy or
 * a move lays the looks of its range by rectangles (`looks/move.ts`), a merge gives each area one look rectangle over the
 * sheet's (`merge.ts`).
 */
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
      // the looks of the range move and are laid by rectangles (`copyRange`, `moveRange`)
      const p = pasteOf(book, op)
      if (!p) break
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
 * new line with the look of the line before it, its cells' with an entry too (`lookBefore`: a rectangle each, however many
 * lines come in). An insert is the one move of lines of its operation, so the line before it is read from `input`.
 */
export function linesMoved(input: BookInput, r: Extract<Applied, { ok: true }>): Applied {
  const lines = r.geometry.flatMap((m) => (m.kind === 'lines' ? [m] : []))
  const sheets = r.book.sheets.slice()
  let moved = false
  for (const s of new Set(lines.map((m) => m.s))) {
    let layer = layerOf(input.sheets[s])
    let cells = sheets[s].cells
    for (const m of lines) {
      if (m.s !== s) continue
      if (layer) layer = shifted(layer, m.axis, m.at, m.n, m.n > 0)
      if (m.n <= 0) continue
      const base = layer ?? emptyLayer()
      const looks = new LookTable(base.looks)
      const edits = lookBefore(input.sheets[s], m.axis, m.at, m.n, looks)
      if (edits.length) layer = edited(base, edits, looks)
      if (layer && sheets[s].tables) cells = lookedContent(cells, layer, m.axis, m.at, m.n)
    }
    if (!layer) continue
    sheets[s] = withLayer(sheets[s], cells, layer)
    moved = true
  }
  return moved ? { ...r, book: { ...r.book, sheets } } : r
}

/**
 * The cells with those of content in new lines [at, at + n) that hold no look of their own given the look of `layer`
 * there, which the line before gave them: a table's calculated columns and headers written in them (`tables/shift.ts`).
 */
function lookedContent(cells: Record<string, CellInput>, layer: Layer, axis: Axis, at: number, n: number): Record<string, CellInput> {
  let out = cells
  const p = { r: 0, c: 0 }
  for (const a in cells) {
    const cell = cells[a]
    if (hasLook(cell)) continue
    readAddress(a, p)
    const i = axis === 'r' ? p.r : p.c
    if (i < at || i >= at + n) continue
    const look = layer.lookAt(p.r, p.c)
    if (!look) continue
    if (out === cells) out = { ...cells }
    out[a] = { ...cell, ...look }
  }
  return out
}
