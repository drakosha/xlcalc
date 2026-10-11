/**
 * The way back of a change that builds a session's book again (a structure operation, new names, a table's header typed
 * over): what another book holds that the current one does not. The cells the operation moved are carried back by its
 * geometry (`structure/geometry.ts`), the rest is kept as it was: the cells the geometry does not bring back as they were
 * (deleted ones, formulas it rewrote), each changed sheet's layout, and the book's own fields. Unchanged sheets are named,
 * not kept. The diff is exact whatever the operation did to the cells, as the geometry never brings two cells to one
 * place and its inverse takes back every cell it brings somewhere: a cell it carries to a cell of another input is kept
 * as it was. Each cell is met once, in the session's comparison of every cell with its old one (`moved.ts`), which
 * gives the diff its pairs (`DiffPass`).
 */
import type { BookInput, CellInput, SheetInput } from '../contract.js'
import { paced, type Steps } from '../exec/budget.js'
import { formatAddress, MAX_COLS, readAddress } from '../model.js'
import { cellCount, type Rect } from '../structure/area.js'
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
  if (xs.reduce((n, x) => n + cellCount(x), 0) > most) return null
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

/** A changed sheet of the other book while its diff is taken. */
interface Back {
  /** Its cells to put back so far, by address (null: none there). */
  put: Record<string, CellInput | null>
  /** The cells of changed sheets of the current book that come back to one of its cells as it is (`kept`), or as another input (`held`). */
  kept: number
  held: number
  /** Its cells no such cell comes back to, met at their place. */
  left: number
  /**
   * The places (row · MAX_COLS + column) the geometry takes nowhere, which may hold a cell; null past `GONE` of them (a
   * deleted sheet, many lines deleted), when the sheet's cells are counted instead.
   */
  gone: number[] | null
}

/** The places taken nowhere a sheet's diff looks up one by one, at most: more cost more than counting a large sheet's cells. */
const GONE = 1 << 16

/**
 * The diff's pass over the cells of `current` and `other`, made by the session's comparison of each cell with the one at
 * its old place (`movedChanges`), which meets every cell of both books: each cell of `other` with the cell of `current`
 * at its place (`pair`) or none (`gone`), and each cell of `current` that came from a place `other` holds nothing at.
 * A sheet of `other` that is a sheet object of `current` is unchanged and not met.
 */
export class DiffPass {
  /** The changed sheets of `other`, by index. */
  readonly back: (Back | undefined)[]
  /** Whether a sheet of `current`, by index, is a changed one, whose cells the diff carries. */
  readonly moving: boolean[]

  constructor(
    readonly current: BookInput,
    readonly other: BookInput,
  ) {
    const [mine, theirs] = [new Set(current.sheets), new Set(other.sheets)]
    this.back = other.sheets.map((o) => (mine.has(o) ? undefined : { put: {}, kept: 0, held: 0, left: 0, gone: [] }))
    this.moving = current.sheets.map((c) => !theirs.has(c))
  }

  /** Whether the cells of sheet `s` of `other` are diffed. */
  takes(s: number): boolean {
    return this.back[s] !== undefined
  }

  /** The place `addr` of sheet `s` of `other` (its cell `was`) and the cell `now` at the place the geometry takes it to, on sheet `t` of `current`. */
  pair(s: number, addr: string, was: CellInput | undefined, t: number, now: CellInput | undefined) {
    const b = this.back[s]
    if (!b) return
    if (now !== undefined && this.moving[t]) {
      if (was === now) b.kept++
      else {
        b.put[addr] = was ?? null
        if (was) b.held++
      }
    } else if (was !== undefined) b.left++
  }

  /** A place of sheet `s` of `other` the geometry takes nowhere: its cell, if any, is looked up when the sheet is kept. */
  gone(s: number, r: number, c: number) {
    const b = this.back[s]
    if (!b?.gone) return
    if (b.gone.length < GONE) b.gone.push(r * MAX_COLS + c)
    else b.gone = null
  }
}

/**
 * The diff that gives `other` from `current` by their pass (`DiffPass`), `g` taking the cells of `current` to their places
 * in `other` (the inverse of the operation that made `current` from `other`, or the operation itself). `other` is read
 * cell by cell only when some of its cells come from nowhere (deleted rows, a range overwritten) and the geometry does not
 * tell where they are.
 */
export function* bookDiff(pass: DiffPass, g: Geometry): Steps<BookDiff> {
  const { current, other } = pass
  const away = inverse(g)
  const back = placer(away)
  const movers = pass.moving.flatMap((m, i) => (m ? [i] : []))
  const moving = new Set(movers)
  // the cells of `current`, which the session's state built from it takes as its own and changes
  const live = new Set(current.sheets.map((s) => s.cells))
  const sheets: BookDiff['sheets'] = []
  for (let j = 0; j < other.sheets.length; j++) {
    const [o, d] = [other.sheets[j], pass.back[j]]
    if (!d) {
      sheets.push({ same: o.name })
      continue
    }
    const { cells: _, ...rest } = o
    // a sheet no cell came to (a deleted one) keeps its cells as they are, shared; a copy of those `current` has too
    if (!d.kept && !d.held && !Object.keys(d.put).length) {
      sheets.push({ rest, put: live.has(o.cells) ? { ...o.cells } : o.cells })
      continue
    }
    // its cells: those the pass met, and those at the places taken nowhere looked up; past `GONE` of these, counted
    let size = d.kept + d.held + d.left
    if (d.gone) for (const n of d.gone) size += o.cells[formatAddress(Math.floor(n / MAX_COLS), n % MAX_COLS)] ? 1 : 0
    else {
      size = 0
      for (const _ in o.cells) size++
    }
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

/**
 * The other book, from the current one (as it was when the diff was taken) and the diff. A sheet no cell comes back to
 * (a deleted one) takes the diff's cells as they are, not a copy: a session's state built from the book takes them as
 * its own, unless a host was given them, and the step that held them is taken off the history.
 */
export function* applyDiff(current: BookInput, d: BookDiff): Steps<BookInput> {
  const place = placer(d.geometry)
  const got = d.sheets.map((x) => ('rest' in x ? { cells: {} as Record<string, CellInput>, n: 0 } : undefined))
  for (const i of d.moved)
    yield* carried(current.sheets[i].cells, i, place, (j) => got[j], (t, at, input) => {
      t.cells[at] = input
      t.n++
    })
  const byKey = new Map(current.sheets.map((s) => [sheetKey(s.name), s]))
  const sheets = d.sheets.map((x, j): SheetInput => {
    if ('same' in x) return byKey.get(sheetKey(x.same))!
    const { cells, n } = got[j]!
    // the diff keeps such a sheet's cells whole (`bookDiff`), none of them null
    if (!n) return { ...x.rest, cells: x.put as Record<string, CellInput> }
    for (const addr in x.put) {
      const input = x.put[addr]
      if (input) cells[addr] = input
      else delete cells[addr]
    }
    return { ...x.rest, cells }
  })
  return { ...d.head, sheets }
}
