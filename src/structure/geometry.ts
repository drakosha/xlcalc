/**
 * Where a structure operation takes a book's cells, in numbers: each operation gives its moves, and the rest is made
 * from them: the `mapCell` callers get, how addresses moved (`EditResult.moved`), the way of an undo and a redo, the
 * comparison of each cell with the one at its old place. A sheet is its index in the book of its move, rows and columns
 * are zero-based. A geometry never brings two cells to one place, and its inverse takes back every cell it brings
 * somewhere, so a cell of the book after it that no cell came to is new.
 */
import type { AddressMove, BookInput, CellRef, StructureResult } from '../contract.js'
import { formatAddress, parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { areaText, hasCell, LINES, LIMIT, minus, shiftRect, type Axis, type Rect } from './area.js'

export type Move =
  /** `n` lines inserted before index `at` (`n` > 0), or `-n` lines from `at` deleted; a cell pushed past the sheet's end is gone. */
  | { kind: 'lines'; s: number; axis: Axis; at: number; n: number }
  /** The cells of `src` on sheet `from` cut and pasted `dr` rows and `dc` columns away on sheet `to`; those they land on are gone. */
  | { kind: 'range'; from: number; src: Rect; to: number; dr: number; dc: number }
  /** The cells of `area` gone (a paste over them, a table deleted with its cells); the way back takes the cells there away too. */
  | { kind: 'clear'; s: number; area: Rect }
  | { kind: 'insertSheet'; at: number }
  | { kind: 'deleteSheet'; s: number }
  /** The sheet `s` moved to index `at` of the sheets after the move. */
  | { kind: 'moveSheet'; s: number; at: number }
  | { kind: 'renameSheet'; s: number }

export type Geometry = Move[]

/** An operation's result inside the library: the new book and where its cells went, or Excel's refusal. */
/**
 * An operation applied: the new book and where its cells went; `looked`, the area of sheet `s` on which a copy laid the
 * looks of its source, whose rows a session fits as after a patch of styles.
 */
export type Applied = { ok: true; book: BookInput; geometry: Geometry; looked?: { s: number; area: Rect } } | Exclude<StructureResult, { ok: true }>

/** Applies the steps in turn, each to the book the last one gave; the geometry is theirs one after another. */
export function chain(book: BookInput, steps: ((b: BookInput) => Applied)[]): Applied {
  let out = book
  const geometry: Geometry = []
  for (const step of steps) {
    const r = step(out)
    if (!r.ok) return r
    out = r.book
    geometry.push(...r.geometry)
  }
  return { ok: true, book: out, geometry }
}

/** A cell's place: its sheet's index, its row and its column. */
export interface Place {
  s: number
  r: number
  c: number
}

/** A sheet index after a sheet is taken out of `s` and put in at `at`. */
const reordered = (i: number, s: number, at: number) => {
  const j = i > s ? i - 1 : i
  return j >= at ? j + 1 : j
}

/** One move as a function that carries a place to where the move takes it, in place; false when the cell is gone. */
function mover(m: Move): (p: Place) => boolean {
  switch (m.kind) {
    case 'lines': {
      const { s, at, n } = m
      const [max, end] = [LIMIT[m.axis], at - n]
      const along = (i: number) => (i < at ? i : n > 0 ? (i + n < max ? i + n : -1) : i >= end ? i + n : -1)
      return m.axis === 'r'
        ? (p) => p.s !== s || (p.r = along(p.r)) >= 0
        : (p) => p.s !== s || (p.c = along(p.c)) >= 0
    }
    case 'range': {
      const { from, src, to, dr, dc } = m
      const dst = shiftRect(src, dr, dc)
      return (p) => {
        if (p.s === from && hasCell(src, p.r, p.c)) {
          p.s = to
          p.r += dr
          p.c += dc
          return p.r < LIMIT.r && p.c < LIMIT.c
        }
        return p.s !== to || !hasCell(dst, p.r, p.c)
      }
    }
    case 'clear':
      return (p) => p.s !== m.s || !hasCell(m.area, p.r, p.c)
    case 'insertSheet':
      return (p) => {
        if (p.s >= m.at) p.s++
        return true
      }
    case 'deleteSheet':
      return (p) => {
        if (p.s === m.s) return false
        if (p.s > m.s) p.s--
        return true
      }
    case 'moveSheet':
      return (p) => {
        p.s = p.s === m.s ? m.at : reordered(p.s, m.s, m.at)
        return true
      }
    case 'renameSheet':
      return () => true
  }
}

/** The geometry as one function that carries a place, in place; false when the cell is gone. */
export function placer(g: Geometry): (p: Place) => boolean {
  const fs = g.map(mover)
  if (fs.length === 1) return fs[0]
  return (p) => {
    for (const f of fs) if (!f(p)) return false
    return true
  }
}

function invert(m: Move): Move {
  switch (m.kind) {
    case 'lines':
      return { ...m, n: -m.n }
    case 'range':
      return { kind: 'range', from: m.to, src: shiftRect(m.src, m.dr, m.dc), to: m.from, dr: -m.dr, dc: -m.dc }
    case 'insertSheet':
      return { kind: 'deleteSheet', s: m.at }
    case 'deleteSheet':
      return { kind: 'insertSheet', at: m.s }
    case 'moveSheet':
      return { kind: 'moveSheet', s: m.at, at: m.s }
    default:
      return m
  }
}

/** The way back: from the book after the geometry to the one before it. */
export const inverse = (g: Geometry): Geometry => g.map(invert).reverse()

/**
 * Whether no cell of the book before the geometry comes to sheet `s` of the book after it, a sheet it inserted (or put
 * back): the way back deletes the sheet before a range takes any of its cells off it.
 */
export function newSheet(g: Geometry, s: number): boolean {
  const p: Place = { s, r: 0, c: 0 }
  for (const m of inverse(g)) {
    if (m.kind === 'range' && m.from === p.s) return false
    if ((m.kind === 'insertSheet' || m.kind === 'deleteSheet' || m.kind === 'moveSheet') && !mover(m)(p)) return true
  }
  return false
}

/**
 * The rectangles of sheet `s` whose cells the geometry takes nowhere (rows deleted, cells landed on or pasted over),
 * for a geometry of one move of cells; null for another (no move, sheets, several moves one after another).
 */
export function vanishing(g: Geometry, s: number): Rect[] | null {
  if (g.length !== 1) return null
  const m = g[0]
  switch (m.kind) {
    case 'lines': {
      if (m.s !== s) return []
      const [a, b] = m.n < 0 ? [m.at, m.at - m.n - 1] : [LIMIT[m.axis] - m.n, LIMIT[m.axis] - 1]
      return [m.axis === 'r' ? { r1: a, r2: b, c1: 0, c2: LIMIT.c - 1 } : { r1: 0, r2: LIMIT.r - 1, c1: a, c2: b }]
    }
    case 'range':
      if (m.to !== s) return []
      return m.from === s ? minus(shiftRect(m.src, m.dr, m.dc), m.src) : [shiftRect(m.src, m.dr, m.dc)]
    case 'clear':
      return m.s === s ? [m.area] : []
    default:
      return null
  }
}

/** Where a cell of the book whose sheets are `before` is in the one whose sheets are `after`: a sheet as `after` spells it. */
export function mapCellOf(g: Geometry, before: string[], after: string[]): (ref: CellRef) => CellRef | null {
  const place = placer(g)
  const index = new Map(before.map((name, i) => [sheetKey(name), i]))
  return (ref) => {
    const s = index.get(sheetKey(ref.sheet))
    if (s === undefined) return null
    const { row, col } = parseAddress(ref.cell)
    const p = { s, r: row, c: col }
    if (!place(p)) return null
    return { sheet: after[p.s], cell: p.r === row && p.c === col ? ref.cell : formatAddress(p.r, p.c) }
  }
}

/** What a pool's thread sends for a `mapCell` (`GEOMETRY`): the calling side makes the same function from it. */
export interface SentGeometry {
  geometry: Geometry
  before: string[]
  after: string[]
}

/** The geometry of each `mapCell` a session gave, for a pool's thread to send instead of the function. */
export const GEOMETRY = new WeakMap<object, SentGeometry>()

/**
 * How the geometry moved addresses, from the book whose sheets are `before` to the one whose sheets are `after`, in
 * order: lines inserted and deleted, a range cut and pasted, a sheet renamed, deleted, inserted or moved. A paste moves
 * no address.
 */
export function movesOf(g: Geometry, before: string[], after: string[]): AddressMove[] {
  return g.flatMap((m): AddressMove[] => {
    switch (m.kind) {
      case 'lines':
        return [{ kind: 'lines', sheet: after[m.s], axis: LINES[m.axis], at: m.at + 1, count: m.n }]
      case 'range':
        return [{ kind: 'range', from: { sheet: before[m.from], range: areaText(m.src) }, to: { sheet: after[m.to], cell: formatAddress(m.src.r1 + m.dr, m.src.c1 + m.dc) } }]
      case 'renameSheet':
        return [{ kind: 'renameSheet', sheet: before[m.s], name: after[m.s] }]
      case 'deleteSheet':
        return [{ kind: 'deleteSheet', sheet: before[m.s] }]
      case 'insertSheet':
        return [{ kind: 'insertSheet', sheet: after[m.at], index: m.at }]
      case 'moveSheet':
        return [{ kind: 'moveSheet', sheet: after[m.at], index: m.at }]
      case 'clear':
        return []
    }
  })
}
