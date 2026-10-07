/**
 * The areas of array formulas of the old kind (`CellInput.array`) in a book's input: each on its formula's cell, at
 * the area's top-left, nothing else in it. Apart from the computation, so the xlsx subpath checks books without it.
 */
import type { BookInput, CellInput } from './contract.js'
import { InputError } from './input.js'
import { formatAddress, parseAddress } from './model.js'
import { parseRange, type Rect } from './structure/area.js'

/** The area of an array formula on cell `addr`, which is its top-left cell; null when the text is not one. */
export function arrayArea(addr: string, text: string): Rect | null {
  const x = parseRange(text)
  const at = parseAddress(addr)
  return x && x.r1 === at.row && x.c1 === at.col ? x : null
}

/** A cell that holds a value or a formula. */
export const holds = (c: CellInput | undefined) => c !== undefined && (c.formula !== undefined || (c.value !== undefined && c.value !== null))

/** The areas of a sheet's array formulas, by their top-left cell; the other cells of an area hold no value or formula. */
function sheetAreas(cells: Record<string, CellInput>, path: string): Map<string, Rect> {
  const out = new Map<string, Rect>()
  for (const [addr, c] of Object.entries(cells)) {
    if (c.array === undefined) continue
    const x = c.formula === undefined ? null : arrayArea(addr, c.array)
    if (!x) throw new InputError(`${path}.cells[${JSON.stringify(addr)}].array`, `expected the area of the cell's array formula with the cell at its top-left, such as "${addr}:${addr.replace(/\d+$/, (n) => String(Number(n) + 2))}" (and a formula), got ${JSON.stringify(c.array)}`)
    out.set(addr, x)
  }
  if (!out.size) return out
  for (const [addr, c] of Object.entries(cells)) {
    if (!holds(c)) continue
    const { row, col } = parseAddress(addr)
    for (const [top, x] of out)
      if (top !== addr && row >= x.r1 && row <= x.r2 && col >= x.c1 && col <= x.c2)
        throw new InputError(`${path}.cells[${JSON.stringify(addr)}]`, `the cell is in the area of the array formula at ${top} (${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2, x.c2)}), which holds nothing else`)
  }
  return out
}

/** A book's array formulas: each area on its formula's cell, nothing else in it. */
export function checkArrays(book: BookInput, path: string) {
  book.sheets.forEach((s, i) => sheetAreas(s.cells, `${path}.sheets[${i}]`))
}
