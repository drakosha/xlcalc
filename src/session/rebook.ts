/**
 * The way back of a change that builds a session's book again (a structure operation, new names, a table's header typed
 * over): what another book holds that the current one does not. The cells the operation moved are carried back by its
 * geometry (`carry`), the rest is kept as it was: the cells the geometry does not bring back as they were (deleted ones,
 * formulas it rewrote), each changed sheet's layout, and the book's own fields. Unchanged sheets are named, not kept.
 * The diff is exact whatever the geometry gives, as long as it never brings two cells to one place: a cell it carries
 * wrong is kept as it was.
 */
import type { BookInput, CellInput, CellRef, SheetInput, StructureOp } from '../contract.js'
import { paced, type Steps } from '../exec/budget.js'
import { colToIndex, indexToCol, MAX_COLS, MAX_ROWS } from '../model.js'
import { parseRange } from '../structure/area.js'
import { sheetKey } from '../syntax.js'

/** The geometry's direction: from the book after the operation to the one before (`back`), or the other way. */
export type Way = 'back' | 'forth'

/**
 * How the cells of one sheet go to the other book: `to`, the sheet they are on there (renamed or not), and `at`, where
 * each cell goes (null: it is gone; a sheet and an address, for a range moved to another sheet); without `at` each cell
 * stays at its address.
 */
type SheetCarry = { to: string; at?: (addr: string) => [string, string] | null }
type Carry = (sheet: string) => SheetCarry

export interface BookDiff {
  /** The other book's fields but its sheets. */
  head: Omit<BookInput, 'sheets'>
  op: StructureOp | null
  way: Way
  /** The changed sheets of the current book, whose cells the geometry carries. */
  moved: string[]
  /**
   * The other book's sheets in order: one the current book has unchanged, by name; else its fields and its cells the
   * geometry does not give, by address (null: none there). A sheet the current book has none of keeps its own cells.
   */
  sheets: ({ same: string } | { rest: Omit<SheetInput, 'cells'>; put: Record<string, CellInput | null> })[]
}

type Rect = { r1: number; c1: number; r2: number; c2: number }

/** Where an address's row starts: its column letters are before it. */
const digit = (addr: string) => {
  let k = 0
  while (addr.charCodeAt(k) > 64) k++
  return k
}
const rowOf = (addr: string, k: number) => Number(addr.slice(k)) - 1
const colOf = (addr: string, k: number) => colToIndex(addr.slice(0, k))

const same: Carry = (sheet) => ({ to: sheet })

/** Rows (`r`) or columns inserted (`n` > 0) or deleted (`n` < 0) at index `at` of one sheet, as a cell's index moves. */
function shifted(key: string, rows: boolean, at: number, n: number): Carry {
  const max = rows ? MAX_ROWS : MAX_COLS
  return (sheet) => ({
    to: sheet,
    ...(sheetKey(sheet) === key && {
      at(addr: string): [string, string] | null {
        const k = digit(addr)
        const i = rows ? rowOf(addr, k) : colOf(addr, k)
        if (i < at) return [sheet, addr]
        const j = i + n
        if ((n < 0 && i < at - n) || j >= max) return null
        return [sheet, rows ? addr.slice(0, k) + (j + 1) : indexToCol(j) + addr.slice(k)]
      },
    }),
  })
}

const inside = (r: number, c: number, x: Rect) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2

/** A rectangle moved from one sheet to another (cut and paste): the cells it covered where it lands are gone. */
function moved(from: string, src: Rect, to: string, dr: number, dc: number): Carry {
  const [a, b] = [sheetKey(from), sheetKey(to)]
  const dst = { r1: src.r1 + dr, r2: src.r2 + dr, c1: src.c1 + dc, c2: src.c2 + dc }
  return (sheet) => {
    const key = sheetKey(sheet)
    if (key !== a && key !== b) return { to: sheet }
    return {
      to: sheet,
      at(addr) {
        const k = digit(addr)
        const [row, col] = [rowOf(addr, k), colOf(addr, k)]
        if (key === a && inside(row, col, src)) return row + dr < MAX_ROWS && col + dc < MAX_COLS ? [to, indexToCol(col + dc) + (row + dr + 1)] : null
        return key === b && inside(row, col, dst) ? null : [sheet, addr]
      },
    }
  }
}

/** The cells of a rectangle of one sheet come from nowhere (a paste), the rest stay. */
function pasted(to: string, area: Rect): Carry {
  const key = sheetKey(to)
  return (sheet) => ({
    to: sheet,
    ...(sheetKey(sheet) === key && {
      at(addr: string): [string, string] | null {
        const k = digit(addr)
        return inside(rowOf(addr, k), colOf(addr, k), area) ? null : [sheet, addr]
      },
    }),
  })
}

/**
 * The geometry of an operation (null: none, cells stay where they are), in one direction; it never brings two cells to
 * one place. The way back from an insert deletes what it inserted, from a delete inserts it again.
 */
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
      return shifted(sheetKey(op.sheet), rows, at, n)
    }
    case 'moveRange': {
      const src = parseRange(op.from.range)!
      const at = parseRange(op.to.cell)!
      const [dr, dc] = [at.r1 - src.r1, at.c1 - src.c1]
      return back ? moved(op.to.sheet, { r1: src.r1 + dr, r2: src.r2 + dr, c1: src.c1 + dc, c2: src.c2 + dc }, op.from.sheet, -dr, -dc) : moved(op.from.sheet, src, op.to.sheet, dr, dc)
    }
    case 'copyRange': {
      // the target tiled with copies, or one copy at its top-left; a wrong guess only makes the diff larger
      const src = parseRange(op.from.range)!
      const t = parseRange(op.to.range)!
      const [h, w] = [src.r2 - src.r1 + 1, src.c2 - src.c1 + 1]
      const tiled = (t.r2 - t.r1 + 1) % h === 0 && (t.c2 - t.c1 + 1) % w === 0
      return pasted(op.to.sheet, tiled ? t : { r1: t.r1, c1: t.c1, r2: t.r1 + h - 1, c2: t.c1 + w - 1 })
    }
    case 'renameSheet': {
      const [a, b] = back ? [sheetKey(op.name), op.sheet] : [sheetKey(op.sheet), op.name]
      return (sheet) => ({ to: sheetKey(sheet) === a ? b : sheet })
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
    const k = g(ref.sheet)
    const x = k.at ? k.at(ref.cell) : [k.to, ref.cell]
    const sheet = x && spelled.get(sheetKey(x[0]))
    return sheet ? { sheet, cell: x![1] } : null
  }
}

/**
 * The diff that gives `other` from `current` (`current` after a structure operation `op` and `other` before it, `back`,
 * or the other way round); a sheet of `other` that is a sheet object of `current` is unchanged. The cells of `current`
 * the geometry brings to `other` as they are there cost one lookup each; `other` is read cell by cell only when some of
 * its cells come from nowhere (deleted rows, a range overwritten).
 */
export function* bookDiff(current: BookInput, other: BookInput, op: StructureOp | null, way: Way): Steps<BookDiff> {
  const [g, from] = [carry(op, way), carry(op, way === 'back' ? 'forth' : 'back')]
  const mine = new Set(current.sheets)
  const theirs = new Set(other.sheets)
  const changed = new Map<string, { o: SheetInput; put: Record<string, CellInput | null>; kept: number; held: number }>()
  for (const o of other.sheets) if (!mine.has(o)) changed.set(sheetKey(o.name), { o, put: {}, kept: 0, held: 0 })
  const movedSheets: string[] = []
  for (const c of current.sheets) {
    if (theirs.has(c)) continue
    movedSheets.push(c.name)
    const k = g(c.name)
    const home = changed.get(sheetKey(k.to))
    for (const addr in c.cells) {
      if (paced(1)) yield
      let [d, a] = [home, addr]
      if (k.at) {
        const x = k.at(addr)
        if (!x) continue
        a = x[1]
        if (x[0] !== k.to) d = changed.get(sheetKey(x[0]))
      }
      if (!d) continue
      const was = d.o.cells[a]
      if (was === c.cells[addr]) d.kept++
      else {
        d.put[a] = was ?? null
        if (was) d.held++
      }
    }
  }
  const sheets: BookDiff['sheets'] = []
  const movers = new Map(current.sheets.filter((c) => !theirs.has(c)).map((c) => [sheetKey(c.name), c]))
  for (const o of other.sheets) {
    if (mine.has(o)) {
      sheets.push({ same: o.name })
      continue
    }
    const d = changed.get(sheetKey(o.name))!
    const { cells: _, ...rest } = o
    // a sheet no cell came to (a deleted one) keeps its cells as they are, shared
    if (!d.kept && !d.held && !Object.keys(d.put).length) {
      sheets.push({ rest, put: o.cells })
      continue
    }
    let size = 0
    for (const _ in o.cells) size++
    // cells no cell of `current` came to as they are: found by the geometry the other way, and checked to come back here
    if (d.kept + d.held < size) {
      const k = from(o.name)
      const [home, key] = [movers.get(sheetKey(k.to)), sheetKey(o.name)]
      const carried = new Map<SheetInput, SheetCarry>()
      for (const addr in o.cells) {
        if (paced(1)) yield
        if (addr in d.put) continue
        const x = k.at ? k.at(addr) : [k.to, addr]
        const c = x && (x[0] === k.to ? home : movers.get(sheetKey(x[0])))
        if (c && c.cells[x![1]] === o.cells[addr]) {
          // the cell there comes back here by the geometry
          if (!carried.has(c)) carried.set(c, g(c.name))
          const h = carried.get(c)!
          const y = h.at ? h.at(x![1]) : [h.to, x![1]]
          if (y && y[1] === addr && sheetKey(y[0]) === key) continue
        }
        d.put[addr] = o.cells[addr]
      }
    }
    sheets.push({ rest, put: d.put })
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
    const k = g(c.name)
    const home = got.get(sheetKey(k.to))
    for (const addr in c.cells) {
      if (paced(1)) yield
      if (!k.at) {
        if (home) home[addr] = c.cells[addr]
        continue
      }
      const x = k.at(addr)
      const cells = x && (x[0] === k.to ? home : got.get(sheetKey(x[0])))
      if (cells) cells[x![1]] = c.cells[addr]
    }
  }
  const sheets = d.sheets.map((x): SheetInput => {
    if ('same' in x) return byKey.get(sheetKey(x.same))!
    const cells = got.get(sheetKey(x.rest.name))!
    for (const addr in x.put) {
      const input = x.put[addr]
      if (input) cells[addr] = input
      else delete cells[addr]
    }
    return { ...x.rest, cells }
  })
  return { ...d.head, sheets }
}
