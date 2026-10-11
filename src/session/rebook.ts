/**
 * The way back of a change that builds a session's book again (a structure operation, new names, a table's header typed
 * over): what another book holds that the current one does not. The cells the operation moved are carried back by its
 * geometry (`structure/geometry.ts`), the rest is kept as it was: the cells the geometry does not bring back as they were
 * (deleted ones, formulas it rewrote), each changed sheet's layout, and the book's own fields. Unchanged sheets are named,
 * not kept. The diff is exact whatever the operation did to the cells, as the geometry never brings two cells to one
 * place and its inverse takes back every cell it brings somewhere: a cell it carries to a cell of another input is kept
 * as it was.
 */
import type { BookInput, CellInput, SheetInput } from '../contract.js'
import { paced, type Steps } from '../exec/budget.js'
import { formatAddress, readAddress } from '../model.js'
import type { Rect } from '../structure/area.js'
import { inverse, placer, vanishing, type Geometry, type Place } from '../structure/geometry.js'
import { sheetKey } from '../syntax.js'

export interface BookDiff {
  /** The other book's fields but its sheets. */
  head: Omit<BookInput, 'sheets'>
  /** Where the cells of the current book are in the other one. */
  geometry: Geometry
  /** The changed sheets of the current book, by index, whose cells the geometry carries. */
  moved: number[]
  /**
   * The other book's sheets in order: one the current book has unchanged, by name; else its fields and its cells the
   * geometry does not give, by address (null: none there). A sheet the current book has none of keeps its own cells.
   */
  sheets: ({ same: string } | { rest: Omit<SheetInput, 'cells'>; put: Record<string, CellInput | null> })[]
}

/**
 * Each cell of sheet `s` of a book where `place` takes it, in a sheet `target` gives by index: `visit` gets that sheet's
 * own, the address there, the cell's input and its address here; `gone` the cells it takes nowhere or to no such sheet.
 * A cell that stays on its sheet costs no lookup of the sheet, one that stays at its address no new address.
 */
function* carried<T>(cells: Record<string, CellInput>, s: number, place: (p: Place) => boolean, target: (i: number) => T | undefined, visit: (t: T, at: string, input: CellInput, addr: string) => void, gone?: (input: CellInput, addr: string) => void): Steps<void> {
  const home = target(s)
  const p: Place = { s, r: 0, c: 0 }
  for (const addr in cells) {
    if (paced(1)) yield
    readAddress(addr, p)
    const r = p.r
    const c = p.c
    p.s = s
    const t = place(p) ? (p.s === s ? home : target(p.s)) : undefined
    if (t) visit(t, p.r === r && p.c === c ? addr : formatAddress(p.r, p.c), cells[addr], addr)
    else gone?.(cells[addr], addr)
  }
}

/** The cells at the places of rectangles, looked up one by one; null when they hold more places than `most`. */
function cellsAt(cells: Record<string, CellInput>, xs: Rect[], most: number): [string, CellInput][] | null {
  if (xs.reduce((n, x) => n + (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1), 0) > most) return null
  const out: [string, CellInput][] = []
  for (const x of xs)
    for (let r = x.r1; r <= x.r2; r++)
      for (let c = x.c1; c <= x.c2; c++) {
        const addr = formatAddress(r, c)
        const input = cells[addr]
        if (input) out.push([addr, input])
      }
  return out
}

/**
 * The diff that gives `other` from `current`, `g` taking the cells of `current` to their places in `other` (the inverse of
 * the operation that made `current` from `other`, or the operation itself); a sheet of `other` that is a sheet object of
 * `current` is unchanged. The cells of `current` the geometry brings to `other` as they are there cost one lookup each;
 * `other` is read cell by cell only when some of its cells come from nowhere (deleted rows, a range overwritten).
 */
export function* bookDiff(current: BookInput, other: BookInput, g: Geometry): Steps<BookDiff> {
  const away = inverse(g)
  const [place, back] = [placer(g), placer(away)]
  const [mine, theirs] = [new Set(current.sheets), new Set(other.sheets)]
  type Changed = { put: Record<string, CellInput | null>; kept: number; held: number }
  const changed = other.sheets.map((o): Changed | undefined => (mine.has(o) ? undefined : { put: {}, kept: 0, held: 0 }))
  const movers = current.sheets.flatMap((c, i) => (theirs.has(c) ? [] : [i]))
  for (const i of movers)
    yield* carried(
      current.sheets[i].cells,
      i,
      place,
      (j) => changed[j] && ([changed[j], other.sheets[j].cells] as const),
      ([d, cells], at, input) => {
        const was = cells[at]
        if (was === input) d.kept++
        else {
          d.put[at] = was ?? null
          if (was) d.held++
        }
      },
    )
  const moving = new Set(movers)
  const sheets: BookDiff['sheets'] = []
  for (let j = 0; j < other.sheets.length; j++) {
    const [o, d] = [other.sheets[j], changed[j]]
    if (!d) {
      sheets.push({ same: o.name })
      continue
    }
    const { cells: _, ...rest } = o
    // a sheet no cell came to (a deleted one) keeps its cells as they are, shared
    if (!d.kept && !d.held && !Object.keys(d.put).length) {
      sheets.push({ rest, put: o.cells })
      continue
    }
    let size = 0
    for (const _ in o.cells) size++
    // cells no cell of `current` came to: where the geometry the other way takes cells nowhere, when they are all there
    // (a few rows deleted), else found cell by cell by that geometry
    const left = size - d.kept - d.held
    if (left > 0) {
      const keep = (input: CellInput, addr: string) => {
        if (!(addr in d.put)) d.put[addr] = input
      }
      const holes = vanishing(away, j)
      const found = holes && cellsAt(o.cells, holes, size)
      if (found?.length === left) for (const [addr, input] of found) keep(input, addr)
      else yield* carried(o.cells, j, back, (i) => (moving.has(i) ? current.sheets[i].cells : undefined), (cells, at, input, addr) => cells[at] !== input && keep(input, addr), keep)
    }
    sheets.push({ rest, put: d.put })
  }
  const { sheets: _, ...head } = other
  return { head, geometry: g, moved: movers, sheets }
}

/** The other book, from the current one (as it was when the diff was taken) and the diff. */
export function* applyDiff(current: BookInput, d: BookDiff): Steps<BookInput> {
  const place = placer(d.geometry)
  const got = d.sheets.map((x) => ('rest' in x ? ({} as Record<string, CellInput>) : undefined))
  for (const i of d.moved) yield* carried(current.sheets[i].cells, i, place, (j) => got[j], (cells, at, input) => (cells[at] = input))
  const byKey = new Map(current.sheets.map((s) => [sheetKey(s.name), s]))
  const sheets = d.sheets.map((x, j): SheetInput => {
    if ('same' in x) return byKey.get(sheetKey(x.same))!
    const cells = got[j]!
    for (const addr in x.put) {
      const input = x.put[addr]
      if (input) cells[addr] = input
      else delete cells[addr]
    }
    return { ...x.rest, cells }
  })
  return { ...d.head, sheets }
}
