/**
 * The cells a query reads (`readCells`, `summarizeRanges`), in `readCells`' order: range by range, each range's cells
 * by rows, a cell once, in the first range that has it. Three kinds merged by their places: input cells, from the places
 * of a sheet's cells sorted once and kept with the state (`search.ts`, which Find reads too; a place whose cell is gone
 * is passed over); the cells arrays spill into, from the spills that meet the range; and the cells of a look alone, from
 * the sheet's layer of looks. A query reads only the cells of its ranges, and a page only up to its last cell.
 */
import type { EngineState } from '../engine.js'
import { paced } from '../exec/budget.js'
import { layerOf } from '../looks/sheet.js'
import { colOf, formatAddress, MAX_COLS, rowOf } from '../model.js'
import { lowerBound } from '../sorted.js'
import { hasCell, meets, type Rect } from '../structure/area.js'
import type { Ref } from './find.js'
import { textsOf } from './search.js'

/** A cell of a query: its sheet, row and column, the range it comes in, and its address. */
export type Read = [s: number, r: number, c: number, n: number, addr: string]

/** The rows of `rows` (ascending) from row `r` on, as the index of the first; all rows without `rows`. */
const nextRow = (rows: number[] | null, r: number) => (rows ? (rows[lowerBound(rows, r)] ?? Infinity) : r)

/** The index of the first place of `at` (ascending) from index `j` on in rectangle `x` and rows `rows`, not before place `start`, each row's run found by a search; `at.length` when none. */
export function nextIn(at: Float64Array, x: Rect, rows: number[] | null, start: number, j: number): number {
  while (j < at.length) {
    const m = at[j]
    const r = rowOf(m)
    const c = colOf(m)
    const row = nextRow(rows, r)
    if (row > x.r2) return at.length
    if (row > r || c < x.c1) j = lowerBound(at, Math.max(start, row * MAX_COLS + x.c1))
    else if (c > x.c2) j = lowerBound(at, Math.max(start, nextRow(rows, r + 1) * MAX_COLS + x.c1))
    else return j
  }
  return at.length
}

/** The index of the first place of `at` in rectangle `x` and rows `rows` from place `start` on; `at.length` when none. */
export const firstIn = (at: Float64Array, x: Rect, rows: number[] | null, start: number) => nextIn(at, x, rows, start, lowerBound(at, Math.max(start, x.r1 * MAX_COLS + x.c1)))

/** The places of `at` in rectangle `x` and rows `rows`, by rows, from place `start` on. */
function* indexed(at: Float64Array, x: Rect, rows: number[] | null, start: number): Generator<number> {
  for (let j = firstIn(at, x, rows, start); j < at.length; j = nextIn(at, x, rows, start, j + 1)) yield at[j]
}

/** The places of the cells of rectangles that do not overlap, in rows `rows`, by rows, from place `start` on. */
function* covered(rects: Rect[], rows: number[] | null, start: number): Generator<number> {
  rects.sort((a, b) => a.r1 - b.r1)
  let active: Rect[] = []
  let k = 0
  for (let r = nextRow(rows, Math.floor(start / MAX_COLS)); r !== Infinity; r = nextRow(rows, r + 1)) {
    while (k < rects.length && rects[k].r1 <= r) active.push(rects[k++])
    active = active.filter((x) => x.r2 >= r)
    if (!active.length) {
      if (k === rects.length) return
      // past the rows no rectangle covers
      r = nextRow(rows, rects[k].r1) - 1
      continue
    }
    active.sort((a, b) => a.c1 - b.c1)
    for (const x of active) for (let c = x.c1; c <= x.c2; c++) if (r * MAX_COLS + c >= start) yield r * MAX_COLS + c
  }
}

/**
 * The cells of `refs` in `readCells`' order, after the position `from` (a range and a place, as a page's cursor keeps
 * it); with `rows` (zero-based, ascending), only the cells of those rows; without `looks`, the cells of a look alone left
 * out. Undefined between cells, where the steps may yield.
 */
export function* readOrder(state: EngineState, refs: Ref[], rows: number[] | null, from: readonly number[] | null, looks = true): Generator<Read | void> {
  for (let n = from?.[0] ?? 0; n < refs.length; n++) {
    const x = refs[n]
    const start = from && n === from[0] ? from[1] + 1 : 0
    const sheet = state.input.sheets[x.s]
    const { cells } = sheet
    const spills = [...state.spills.on.values()].filter((y) => y.s === x.s && meets(y, x)).map((y) => ({ r1: Math.max(y.r1, x.r1), r2: Math.min(y.r2, x.r2), c1: Math.max(y.c1, x.c1), c2: Math.min(y.c2, x.c2) }))
    const earlier = refs.slice(0, n).filter((y) => y.s === x.s && meets(y, x))
    const layer = looks ? layerOf(sheet) : undefined
    const pieces = layer ? [...layer.pieces(x)] : []
    const { at } = yield* textsOf(state, x.s)
    // no spill and no look in the range: its cells are the index's alone, but places whose cells are gone
    if (!spills.length && !pieces.length) {
      for (let j = firstIn(at, x, rows, start); j < at.length; j = nextIn(at, x, rows, start, j + 1)) {
        if (paced(1)) yield
        const r = Math.floor(at[j] / MAX_COLS)
        const c = at[j] - r * MAX_COLS
        const addr = formatAddress(r, c)
        if (addr in cells && !(earlier.length && earlier.some((y) => hasCell(y, r, c)))) yield [x.s, r, c, n, addr]
      }
      continue
    }
    const streams = [indexed(at, x, rows, start), covered(spills, rows, start), covered(pieces, rows, start)]
    const heads = streams.map((g) => g.next())
    for (;;) {
      let k = -1
      for (let i = 0; i < heads.length; i++) if (!heads[i].done && (k < 0 || heads[i].value! < heads[k].value!)) k = i
      if (k < 0) break
      const m = heads[k].value as number
      heads[k] = streams[k].next()
      if (paced(1)) yield
      const r = rowOf(m)
      const c = colOf(m)
      const addr = formatAddress(r, c)
      // an input cell from the index alone, a spilled cell from its spill alone, a look's cell where neither is
      if ((k === 0) !== (addr in cells) || (k === 2 && spills.some((y) => hasCell(y, r, c))) || (earlier.length && earlier.some((y) => hasCell(y, r, c)))) continue
      yield [x.s, r, c, n, addr]
    }
  }
}
