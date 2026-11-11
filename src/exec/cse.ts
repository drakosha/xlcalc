/** Coordinate selection and evaluation of the parts of a CSE formula that vary over its output area. */
import type { Arg } from '../args.js'
import { first } from '../arrays/lift.js'
import { Arr, broadcastAt } from '../arrays/value.js'
import { E, Rng, type V } from '../values.js'
import { allot, spend } from './budget.js'

type Cell = (s: number, r: number, c: number) => V

/** A value at one CSE coordinate: singleton axes repeat, and positions past a larger axis are #N/A. */
export function cseAt(x: Arg, r: number, c: number, cell: Cell): Arg {
  if (x instanceof Rng) {
    const row = x.r1 === x.r2 ? x.r1 : x.r1 + r
    const col = x.c1 === x.c2 ? x.c1 : x.c1 + c
    return row > x.r2 || col > x.c2 ? E.NA : cell(x.s, row, col)
  }
  if (!(x instanceof Arr)) return x
  const v = broadcastAt(x, r, c)
  return v === undefined ? E.NA : v
}

/** One evaluation per output coordinate, keeping ordinary arrays inside each evaluation until its result is read. */
export function cseMap(rows: number, cols: number, run: (r: number, c: number) => Arg, cell: Cell): Arr {
  allot(rows * cols)
  const values: V[] = new Array(rows * cols)
  const read = (g: Rng) => cell(g.s, g.r1, g.c1)
  for (let r = 0; r < rows; r++) {
    spend(cols)
    for (let c = 0; c < cols; c++) values[r * cols + c] = first(cseAt(run(r, c), r, c, cell), read)
  }
  return new Arr(rows, cols, rows, cols, values)
}
