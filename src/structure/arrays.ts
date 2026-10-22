/**
 * Array formulas of the old kind (`CellInput.array`) under structure operations, as Excel treats them (the struct set
 * `arrays`): an area moves whole with its formula's cell; inserting lines inside it, deleting part of it, moving part
 * of it and pasting over part of it are refused; a copy of part of it pastes each copied cell as an array formula of
 * its own.
 */
import type { BookInput, CellInput, SheetInput } from '../contract.js'
import type { EngineMessage } from '../locale/messages.js'
import { formatAddress, parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { contains, hi, intersect, lo, meets, parseRange, rectText, shiftRect, type Axis, type Rect } from './area.js'
import { arraysOf } from './grid.js'
import type { Paste } from './paste.js'
import { sheetOf } from './rewrite.js'

/**
 * Excel's refusal of an array formula of several cells that a paste left over a table of sheet `sheet` (checkBook
 * refuses the book), or null.
 */
export function arrayInTable(sheet: SheetInput): EngineMessage | null {
  if (!sheet.tables?.length) return null
  for (const [, x] of arraysOf(sheet)) {
    if (x.r1 === x.r2 && x.c1 === x.c2) continue
    const t = sheet.tables.find((t) => meets(parseRange(t.range)!, x))
    if (t) return { key: 'arrayInTable', area: rectText(x), table: t.name }
  }
  return null
}

/** A cell moved from `from` to `to`: an array formula's area goes along. */
export function relocated(cell: CellInput, from: string, to: string): CellInput {
  if (cell.array === undefined || from === to) return cell
  const [a, b] = [parseAddress(from), parseAddress(to)]
  return { ...cell, array: rectText(shiftRect(parseRange(cell.array)!, b.row - a.row, b.col - a.col)) }
}

/** Why Excel refuses to insert (`n` > 0) or delete lines of `sheet` at `at`, or null: lines inside an area, part of an area deleted. */
export function shiftRefusal(sheet: SheetInput, axis: Axis, at: number, n: number): EngineMessage | null {
  for (const [addr, x] of arraysOf(sheet)) {
    const [a, b] = [lo(x, axis), hi(x, axis)]
    const cut = n > 0 ? at > a && at <= b : at <= b && at - n - 1 >= a && !(at <= a && at - n - 1 >= b)
    if (cut) return refuse(sheet.name, addr, x)
  }
  return null
}

const refuse = (sheet: string, cell: string, x: Rect): EngineMessage => ({ key: 'arrayPart', sheet, cell, area: rectText(x) })

/** Why Excel refuses a move or a paste, or null: a move taking part of an area, cells landing on part of one that stays. */
export function arrayRefusal(book: BookInput, p: Paste): EngineMessage | null {
  if (p.move) for (const [addr, x] of arraysOf(sheetOf(book, p.from))) if (meets(x, p.src) && !contains(p.src, x)) return refuse(p.from, addr, x)
  const moving = (x: Rect) => p.move && sheetKey(p.from) === sheetKey(p.to) && contains(p.src, x)
  for (const [addr, x] of arraysOf(sheetOf(book, p.to))) if (meets(x, p.area) && !contains(p.area, x) && !moving(x)) return refuse(p.to, addr, x)
  return null
}

/**
 * The cells of a copy's source that lie in an area without all of it: each as its own array formula over the cell, its
 * area's formula as copied from the area's cell to it (`shift`), with the cell's own looks, note, link and meta; whole
 * areas copy as they are.
 */
export function copiedCells(sheet: SheetInput, src: Rect, shift: (formula: string, dr: number, dc: number) => string): [string, CellInput][] {
  const out: [string, CellInput][] = []
  for (const [addr, x] of arraysOf(sheet)) {
    const part = contains(src, x) ? null : intersect(x, src)
    if (!part) continue
    const formula = sheet.cells[addr].formula!
    for (let r = part.r1; r <= part.r2; r++)
      for (let c = part.c1; c <= part.c2; c++) {
        // the area's other cells hold no formula or value (checkArrays), only their own looks, note, link and meta
        const at = formatAddress(r, c)
        out.push([at, { ...sheet.cells[at], formula: shift(formula, r - x.r1, c - x.c1), array: at }])
      }
  }
  return out
}
