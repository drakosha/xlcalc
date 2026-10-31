/**
 * The areas of array formulas of the old kind (`CellInput.array`) in a book's input: each on its formula's cell, at
 * the area's top-left, nothing else in it. Apart from the computation, so the xlsx subpath checks books without it.
 */
import type { CellInput } from './contract.js'
import { InputError } from './input-error.js'
import { parseAddress } from './model.js'
import { cornersText, hasCell, meets, parseRange, type Rect } from './structure/area.js'

/** The area of an array formula on cell `addr`, which is its top-left cell; null when the text is not one. */
export function arrayArea(addr: string, text: string): Rect | null {
  const x = parseRange(text)
  const at = parseAddress(addr)
  return x && x.r1 === at.row && x.c1 === at.col ? x : null
}

/** A cell that holds a value or a formula. */
export const holds = (c: CellInput | undefined) => c !== undefined && (c.formula !== undefined || (c.value !== undefined && c.value !== null))

/** The fields of a cell's content: a formula's value held while a circular reference stops it goes with the formula. */
export const CONTENT = ['value', 'formula', 'array', 'held'] as const

/** The two areas cross, as Excel never lets array formulas do: the text of the refusal; null when they do not meet. */
export const crossing = (x: Rect, top: string, y: Rect, other: string): string | null =>
  meets(x, y) ? `the area of the array formula at ${top} (${cornersText(x)}) overlaps the area of the array formula at ${other} (${cornersText(y)})` : null

/** A sheet's array formulas, on the cells `tops`: each area on its formula's cell, nothing else in it, no two crossing. */
export function checkArrays(cells: Record<string, CellInput>, tops: string[], path: string) {
  const areas = new Map<string, Rect>()
  for (const addr of tops) {
    const c = cells[addr]
    const x = c.formula === undefined ? null : arrayArea(addr, c.array!)
    if (!x) throw new InputError(`${path}.cells[${JSON.stringify(addr)}].array`, `expected the area of the cell's array formula with the cell at its top-left, such as "${addr}:${addr.replace(/\d+$/, (n) => String(Number(n) + 2))}" (and a formula), got ${JSON.stringify(c.array)}`)
    areas.set(addr, x)
  }
  for (const addr in cells) {
    if (!holds(cells[addr])) continue
    const { row, col } = parseAddress(addr)
    for (const [top, x] of areas)
      if (top !== addr && hasCell(x, row, col))
        throw new InputError(`${path}.cells[${JSON.stringify(addr)}]`, `the cell is inside the area of the array formula at ${top} (${cornersText(x)}), so it cannot hold anything else`)
  }
  // two areas that cross with neither top in the other: a cell of both holds no input
  const list = [...areas]
  for (let i = 1; i < list.length; i++)
    for (let j = 0; j < i; j++) {
      const why = crossing(list[i][1], list[i][0], list[j][1], list[j][0])
      if (why) throw new InputError(`${path}.cells[${JSON.stringify(list[i][0])}].array`, why)
    }
}
