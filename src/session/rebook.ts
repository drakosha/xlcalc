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

/** Where the geometry takes a cell of a sheet: a sheet and an address, null when it is gone. */
const place = (k: SheetCarry, addr: string): [string, string] | null => (k.at ? k.at(addr) : [k.to, addr])

/** Where a cell of the current book is in the other one, by the geometry; sheets spelled as `other` spells them. */
export function mapCellOf(op: StructureOp | null, way: Way, other: string[]): (ref: CellRef) => CellRef | null {
  const g = carry(op, way)
  const spelled = new Map(other.map((n) => [sheetKey(n), n]))
  return (ref) => {
    const x = place(g(ref.sheet), ref.cell)
    const sheet = x && spelled.get(sheetKey(x[0]))
    return sheet ? { sheet, cell: x![1] } : null
  }
}

/**
 * Each cell of `sheet` where the geometry `k` takes it, in a sheet `target` gives by name: `visit` gets that sheet's own,
 * the address there, the cell's input and its address here; `gone` the cells it takes nowhere or to no such sheet. A
 * cell that stays on its sheet costs no lookup of the sheet.
 */
function* carried<T>(sheet: SheetInput, k: SheetCarry, target: (name: string) => T | undefined, visit: (t: T, at: string, input: CellInput, addr: string) => void, gone?: (input: CellInput, addr: string) => void): Steps<void> {
  const home = target(k.to)
  for (const addr in sheet.cells) {
    if (paced(1)) yield
    let [t, at] = [home, addr]
    if (k.at) {
      const x = k.at(addr)
      if (x) {
        at = x[1]
        if (x[0] !== k.to) t = target(x[0])
      } else t = undefined
    }
    if (t) visit(t, at, sheet.cells[addr], addr)
    else gone?.(sheet.cells[addr], addr)
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
  const [mine, theirs] = [new Set(current.sheets), new Set(other.sheets)]
  type Changed = { o: SheetInput; put: Record<string, CellInput | null>; kept: number; held: number }
  const changed = new Map<string, Changed>()
  for (const o of other.sheets) if (!mine.has(o)) changed.set(sheetKey(o.name), { o, put: {}, kept: 0, held: 0 })
  const movers = current.sheets.filter((c) => !theirs.has(c))
  const there = (name: string) => changed.get(sheetKey(name))
  for (const c of movers)
    yield* carried(c, g(c.name), there, (d, at, input) => {
      const was = d.o.cells[at]
      if (was === input) d.kept++
      else {
        d.put[at] = was ?? null
        if (was) d.held++
      }
    })
  const byKey = new Map(movers.map((c) => [sheetKey(c.name), c]))
  const here = (name: string) => byKey.get(sheetKey(name))
  const sheets: BookDiff['sheets'] = []
  for (const o of other.sheets) {
    if (mine.has(o)) {
      sheets.push({ same: o.name })
      continue
    }
    const d = there(o.name)!
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
      const keep = (input: CellInput, addr: string) => {
        if (!(addr in d.put)) d.put[addr] = input
      }
      yield* carried(o, from(o.name), here, (c, at, input, addr) => {
        const y = c.cells[at] === input && place(g(c.name), at)
        if (!y || y[1] !== addr || sheetKey(y[0]) !== sheetKey(o.name)) keep(input, addr)
      }, keep)
    }
    sheets.push({ rest, put: d.put })
  }
  const { sheets: _, ...head } = other
  return { head, op, way, moved: movers.map((c) => c.name), sheets }
}

/** The other book, from the current one (as it was when the diff was taken) and the diff. */
export function* applyDiff(current: BookInput, d: BookDiff): Steps<BookInput> {
  const g = carry(d.op, d.way)
  const byKey = new Map(current.sheets.map((s) => [sheetKey(s.name), s]))
  const got = new Map<string, Record<string, CellInput>>()
  for (const x of d.sheets) if ('rest' in x) got.set(sheetKey(x.rest.name), {})
  for (const name of d.moved) {
    const c = byKey.get(sheetKey(name))!
    yield* carried(c, g(c.name), (n) => got.get(sheetKey(n)), (cells, at, input) => (cells[at] = input))
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
