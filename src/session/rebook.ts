/**
 * The way back of a change that builds a session's book again (a structure operation, new names, a table's header typed
 * over): what another book holds that the current one does not. The cells the operation moved are carried back by its
 * geometry (`carry`), the rest is kept as it was: the cells the geometry does not bring back as they were (deleted ones,
 * formulas it rewrote), each changed sheet's layout, and the book's own fields. Unchanged sheets are named, not kept.
 * The diff is exact whatever the geometry gives: a cell it carries wrong is kept as it was.
 */
import type { BookInput, CellInput, CellRef, SheetInput, StructureOp } from '../contract.js'
import { paced, type Steps } from '../exec/budget.js'
import { colToIndex, formatAddress, MAX_COLS, MAX_ROWS, parseAddress } from '../model.js'
import { parseRange } from '../structure/area.js'
import { sheetKey } from '../syntax.js'

/** The geometry's direction: from the book after the operation to the one before (`back`), or the other way. */
export type Way = 'back' | 'forth'

/** A cell of one book as a cell of the other: its sheet and address; null when nothing of the other book is there. */
type Carry = (sheet: string, addr: string) => [string, string] | null

export interface BookDiff {
  /** The other book's fields but its sheets. */
  head: Omit<BookInput, 'sheets'>
  op: StructureOp | null
  way: Way
  /** The changed sheets of the current book, whose cells the geometry carries. */
  moved: string[]
  /** The other book's sheets in order: one the current book has unchanged, by name; else its fields and its cells the geometry does not give (null: none). */
  sheets: ({ same: string } | { rest: Omit<SheetInput, 'cells'>; put: [string, CellInput | null][] })[]
}

const inside = (r: number, c: number, x: { r1: number; c1: number; r2: number; c2: number }) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2

/** Rows (`r`) or columns inserted (`n` > 0) or deleted (`n` < 0) at index `at` of one sheet, as a cell's index moves. */
function shifted(key: string, axis: 'r' | 'c', at: number, n: number): Carry {
  const max = axis === 'r' ? MAX_ROWS : MAX_COLS
  return (sheet, addr) => {
    if (sheetKey(sheet) !== key) return [sheet, addr]
    const p = parseAddress(addr)
    const i = axis === 'r' ? p.row : p.col
    if (i < at) return [sheet, addr]
    if (n < 0 && i < at - n) return null
    const j = i + n
    if (j >= max) return null
    return [sheet, axis === 'r' ? formatAddress(j, p.col) : formatAddress(p.row, j)]
  }
}

/** A rectangle moved from one sheet to another (cut and paste): the cells it covered where it lands are gone. */
function moved(from: string, src: { r1: number; c1: number; r2: number; c2: number }, to: string, dr: number, dc: number): Carry {
  const [a, b] = [sheetKey(from), sheetKey(to)]
  const dst = { r1: src.r1 + dr, r2: src.r2 + dr, c1: src.c1 + dc, c2: src.c2 + dc }
  return (sheet, addr) => {
    const k = sheetKey(sheet)
    if (k !== a && k !== b) return [sheet, addr]
    const { row, col } = parseAddress(addr)
    if (k === a && inside(row, col, src)) return row + dr < MAX_ROWS && col + dc < MAX_COLS ? [to, formatAddress(row + dr, col + dc)] : null
    return k === b && inside(row, col, dst) ? null : [sheet, addr]
  }
}

const same: Carry = (sheet, addr) => [sheet, addr]

/** The geometry of an operation (null: none, cells stay where they are), in one direction. */
export function carry(op: StructureOp | null, way: Way): Carry {
  if (!op) return same
  const back = way === 'back'
  switch (op.op) {
    case 'insertRows':
    case 'deleteRows':
    case 'insertColumns':
    case 'deleteColumns': {
      const rows = op.op === 'insertRows' || op.op === 'deleteRows'
      const at = rows ? (op as { row: number }).row - 1 : colToIndex((op as { column: string }).column)
      const n = (op.count ?? 1) * (op.op.startsWith('insert') === back ? -1 : 1)
      // the way back from an insert deletes the rows inserted, from a delete inserts them again
      return shifted(sheetKey(op.sheet), rows ? 'r' : 'c', at, n)
    }
    case 'moveRange': {
      const src = parseRange(op.from.range)!
      const at = parseAddress(op.to.cell)
      const [dr, dc] = [at.row - src.r1, at.col - src.c1]
      return back ? moved(op.to.sheet, { r1: src.r1 + dr, r2: src.r2 + dr, c1: src.c1 + dc, c2: src.c2 + dc }, op.from.sheet, -dr, -dc) : moved(op.from.sheet, src, op.to.sheet, dr, dc)
    }
    case 'copyRange': {
      // the pasted cells come from nowhere either way: the copy's source, the paste's target, tiled or not
      const src = parseRange(op.from.range)!
      const t = parseRange(op.to.range)!
      const [h, w] = [src.r2 - src.r1 + 1, src.c2 - src.c1 + 1]
      const tiled = (t.r2 - t.r1 + 1) % h === 0 && (t.c2 - t.c1 + 1) % w === 0
      const area = tiled ? t : { r1: t.r1, c1: t.c1, r2: t.r1 + h - 1, c2: t.c1 + w - 1 }
      const key = sheetKey(op.to.sheet)
      return (sheet, addr) => {
        if (sheetKey(sheet) !== key) return [sheet, addr]
        const { row, col } = parseAddress(addr)
        return inside(row, col, area) ? null : [sheet, addr]
      }
    }
    case 'renameSheet': {
      const [a, b] = back ? [sheetKey(op.name), op.sheet] : [sheetKey(op.sheet), op.name]
      return (sheet, addr) => [sheetKey(sheet) === a ? b : sheet, addr]
    }
    default:
      return same
  }
}

/** What a `mapCell` of an undone or redone structure operation is made of: the operation, the way, the sheets of both books. */
export interface Geometry {
  op: StructureOp | null
  way: Way
  before: string[]
  after: string[]
}

/** The geometry of each `mapCell` an undo or a redo gave, for a pool's thread to send instead of the function. */
export const GEOMETRY = new WeakMap<object, Geometry>()

/** Where a cell of the current book is in the other one, by the geometry; sheets spelled as `other` spells them. */
export function mapCellOf(op: StructureOp | null, way: Way, other: string[]): (ref: CellRef) => CellRef | null {
  const g = carry(op, way)
  const spelled = new Map(other.map((n) => [sheetKey(n), n]))
  return (ref) => {
    const x = g(ref.sheet, ref.cell)
    const sheet = x && spelled.get(sheetKey(x[0]))
    return sheet ? { sheet, cell: x![1] } : null
  }
}

/**
 * The diff that gives `other` from `current` (`current` after a structure operation `op` and `other` before it, `back`,
 * or the other way round); a sheet of `other` that is a sheet object of `current` is unchanged.
 */
export function* bookDiff(current: BookInput, other: BookInput, op: StructureOp | null, way: Way): Steps<BookDiff> {
  const g = carry(op, way)
  const mine = new Set(current.sheets)
  const theirs = new Set(other.sheets)
  // the other book's changed sheets: the cells the geometry brings to each
  const got = new Map<string, Map<string, CellInput>>()
  const clash = new Map<string, Set<string>>()
  for (const o of other.sheets) if (!mine.has(o)) got.set(sheetKey(o.name), new Map())
  const movedSheets: string[] = []
  for (const c of current.sheets) {
    if (theirs.has(c)) continue
    movedSheets.push(c.name)
    for (const addr in c.cells) {
      if (paced(1)) yield
      const to = g(c.name, addr)
      const cells = to && got.get(sheetKey(to[0]))
      if (!cells) continue
      if (cells.has(to![1])) {
        if (!clash.has(sheetKey(to![0]))) clash.set(sheetKey(to![0]), new Set())
        clash.get(sheetKey(to![0]))!.add(to![1])
      }
      cells.set(to![1], c.cells[addr])
    }
  }
  const sheets: BookDiff['sheets'] = []
  for (const o of other.sheets) {
    if (mine.has(o)) {
      sheets.push({ same: o.name })
      continue
    }
    const k = sheetKey(o.name)
    const cells = got.get(k)!
    const twice = clash.get(k)
    const put: [string, CellInput | null][] = []
    for (const addr in o.cells) {
      if (paced(1)) yield
      if (twice?.has(addr) || cells.get(addr) !== o.cells[addr]) put.push([addr, o.cells[addr]])
      cells.delete(addr)
    }
    // cells the geometry brings where the other book has none
    for (const addr of cells.keys()) put.push([addr, null])
    const { cells: _, ...rest } = o
    sheets.push({ rest, put })
  }
  const { sheets: _, ...head } = other
  return { head, op, way, moved: movedSheets, sheets }
}

/** The other book, from the current one (as it was when the diff was taken) and the diff. */
export function* applyDiff(current: BookInput, d: BookDiff): Steps<BookInput> {
  const g = carry(d.op, d.way)
  const byKey = new Map(current.sheets.map((s) => [sheetKey(s.name), s]))
  const got = new Map<string, Record<string, CellInput>>()
  for (const x of d.sheets) if ('rest' in x) got.set(sheetKey(x.rest.name), {})
  for (const name of d.moved) {
    const c = byKey.get(sheetKey(name))!
    for (const addr in c.cells) {
      if (paced(1)) yield
      const to = g(c.name, addr)
      const cells = to && got.get(sheetKey(to[0]))
      if (cells) cells[to![1]] = c.cells[addr]
    }
  }
  const sheets = d.sheets.map((x): SheetInput => {
    if ('same' in x) return byKey.get(sheetKey(x.same))!
    const cells = got.get(sheetKey(x.rest.name))!
    for (const [addr, input] of x.put) {
      if (input) cells[addr] = input
      else delete cells[addr]
    }
    return { ...x.rest, cells }
  })
  return { ...d.head, sheets }
}
