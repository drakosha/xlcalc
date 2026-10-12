/**
 * What rows and columns give a cell that has no style or number format of its own: its row's, else its column's,
 * the style and the format each separately. The cells' looks (`shown.ts`), AutoFit, copy and paste, data validation's
 * lists and the xlsx writer and reader take a cell's look from here.
 */
import type { CellInput, ColumnInput, NumberFormat, RowInput, SheetInput } from '../contract.js'
import type { EngineState } from '../engine.js'
import { inputIn } from '../looks/sheet.js'
import { formatAddress } from '../model.js'
import { entryAt, lines, type Line } from './spans.js'

/** A style (by index) and a number format, of a cell or a line. */
export interface Look {
  style?: number
  numberFormat?: NumberFormat
}

/** The format a cell shows by: its own, else its row's, else its column's. */
export const formatOf = (own: Look | undefined, row: Look | undefined, col: Look | undefined): NumberFormat | undefined => own?.numberFormat ?? row?.numberFormat ?? col?.numberFormat

/** Where the format a cell shows by comes from (`formatOf`); undefined for none (General). */
export const formatSource = (own: Look | undefined, row: Look | undefined, col: Look | undefined): 'cell' | 'row' | 'column' | undefined =>
  own?.numberFormat !== undefined ? 'cell' : row?.numberFormat !== undefined ? 'row' : col?.numberFormat !== undefined ? 'column' : undefined

/** The look a cell shows: its own style and format, each else its row's, else its column's (`formatOf`). */
export const lookOf = (own: Look | undefined, row: Look | undefined, col: Look | undefined): Look => ({ style: own?.style ?? row?.style ?? col?.style, numberFormat: formatOf(own, row, col) })

/** A sheet's rows and columns, each line looked up once. */
export interface SheetLines {
  row(r: number): RowInput | undefined
  column(c: number): ColumnInput | undefined
  /** The look a cell shows (`lookOf`): `own`, else the sheet's input at the cell or its look of the sheet's looks. */
  look(r: number, c: number, own?: CellInput): Look
  /** The format a cell shows by (`formatOf`). */
  format(r: number, c: number, own?: CellInput): NumberFormat | undefined
}

/**
 * The lines of a sheet's rows or columns, read once for a state (a session's): a change of the layout, a structure
 * operation and an undo give the sheet new maps, never change one, so a map read once stays as it was read.
 */
const read = new WeakMap<EngineState, WeakMap<object, Line<unknown>[]>>()
function linesOf<T>(state: EngineState, map: Record<string, T> | undefined, axis: 'r' | 'c'): Line<T>[] {
  if (!map) return []
  let maps = read.get(state)
  if (!maps) read.set(state, (maps = new WeakMap()))
  let ls = maps.get(map) as Line<T>[] | undefined
  if (!ls) maps.set(map, (ls = lines(map, axis)))
  return ls
}

/** A sheet's lines; with the session's `state`, read once for it. */
export function sheetLines(sh: Pick<SheetInput, 'cells' | 'looks' | 'rows' | 'columns'>, state?: EngineState): SheetLines {
  const rows = state ? linesOf(state, sh.rows, 'r') : lines(sh.rows, 'r')
  const cols = state ? linesOf(state, sh.columns, 'c') : lines(sh.columns, 'c')
  const lineOf = <T>(cache: Map<number, T | undefined>, ls: Line<T>[], i: number): T | undefined => {
    if (!ls.length) return undefined
    if (!cache.has(i)) cache.set(i, entryAt(ls, i))
    return cache.get(i)
  }
  const [rowCache, colCache] = [new Map<number, RowInput | undefined>(), new Map<number, ColumnInput | undefined>()]
  const row = (r: number) => lineOf(rowCache, rows, r)
  const column = (c: number) => lineOf(colCache, cols, c)
  const at = (r: number, c: number) => inputIn(sh, formatAddress(r, c), r, c)
  return { row, column, look: (r, c, own = at(r, c)) => lookOf(own, row(r), column(c)), format: (r, c, own = at(r, c)) => formatOf(own, row(r), column(c)) }
}
