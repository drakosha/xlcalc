/**
 * Array formulas of the old kind (`CellInput.array`) under structure operations, as Excel treats them (the struct set
 * `arrays`): an area moves whole with its formula's cell; inserting lines inside it, deleting part of it, moving part
 * of it and pasting over part of it are refused; a copy of part of it pastes each copied cell as an array formula of
 * its own.
 */
import type { BookInput, CellInput, SheetInput } from '../contract.js'
import { spend } from '../exec/budget.js'
import { formatAddress, parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { contains, meets, parseRange, type Axis, type Rect } from './area.js'
import type { Paste } from './paste.js'

/** The array formulas of a sheet: their cell and area. */
export function areasOf(sheet: SheetInput): [string, Rect][] {
  const out: [string, Rect][] = []
  for (const addr of Object.keys(sheet.cells)) {
    const c = sheet.cells[addr]
    spend(1)
    if (c.array !== undefined && c.formula !== undefined) out.push([addr, parseRange(c.array)!])
  }
  return out
}

const text = (x: Rect) => (x.r1 === x.r2 && x.c1 === x.c2 ? formatAddress(x.r1, x.c1) : `${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2, x.c2)}`)

/** A cell moved from `from` to `to`: an array formula's area goes along. */
export function relocated(cell: CellInput, from: string, to: string): CellInput {
  if (cell.array === undefined || from === to) return cell
  const [a, b] = [parseAddress(from), parseAddress(to)]
  const x = parseRange(cell.array)!
  const [dr, dc] = [b.row - a.row, b.col - a.col]
  return { ...cell, array: text({ r1: x.r1 + dr, r2: x.r2 + dr, c1: x.c1 + dc, c2: x.c2 + dc }) }
}

/** Why Excel refuses to insert (`n` > 0) or delete lines of `sheet` at `at`, or null: lines inside an area, part of an area deleted. */
export function shiftRefusal(sheet: SheetInput, axis: Axis, at: number, n: number): string | null {
  const [lo, hi] = axis === 'r' ? (['r1', 'r2'] as const) : (['c1', 'c2'] as const)
  for (const [addr, x] of areasOf(sheet)) {
    const cut = n > 0 ? at > x[lo] && at <= x[hi] : at <= x[hi] && at - n - 1 >= x[lo] && !(at <= x[lo] && at - n - 1 >= x[hi])
    if (cut) return `the operation changes part of the array formula ${sheet.name}!${addr} (${text(x)})`
  }
  return null
}

const sheetOf = (book: BookInput, name: string) => book.sheets.find((s) => sheetKey(s.name) === sheetKey(name))!

/** Why Excel refuses a move or a paste, or null: a move taking part of an area, cells landing on part of one that stays. */
export function arrayRefusal(book: BookInput, p: Paste): string | null {
  const refuse = (sheet: string, addr: string, x: Rect) => `the operation changes part of the array formula ${sheet}!${addr} (${text(x)})`
  if (p.move) for (const [addr, x] of areasOf(sheetOf(book, p.from))) if (meets(x, p.src) && !contains(p.src, x)) return refuse(p.from, addr, x)
  const moving = (x: Rect) => p.move && sheetKey(p.from) === sheetKey(p.to) && contains(p.src, x)
  for (const [addr, x] of areasOf(sheetOf(book, p.to))) if (meets(x, p.area) && !contains(p.area, x) && !moving(x)) return refuse(p.to, addr, x)
  return null
}

/**
 * The cells of a copy's source that lie in an area without all of it: each as its own array formula over the cell, its
 * area's formula as copied from the area's cell to it (`shift`); whole areas copy as they are.
 */
export function copiedCells(sheet: SheetInput, src: Rect, shift: (formula: string, dr: number, dc: number) => string): [string, CellInput][] {
  const out: [string, CellInput][] = []
  for (const [addr, x] of areasOf(sheet)) {
    if (!meets(x, src) || contains(src, x)) continue
    const top = sheet.cells[addr]
    for (let r = Math.max(x.r1, src.r1); r <= Math.min(x.r2, src.r2); r++)
      for (let c = Math.max(x.c1, src.c1); c <= Math.min(x.c2, src.c2); c++) out.push([formatAddress(r, c), { ...top, formula: shift(top.formula!, r - x.r1, c - x.c1), array: formatAddress(r, c) }])
  }
  return out
}
