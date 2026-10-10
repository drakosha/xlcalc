/**
 * What rows and columns give a cell that has no style or number format of its own: its row's, else its column's,
 * the style and the format each separately. The cells' looks (`shown.ts`), AutoFit, copy and paste, data validation's
 * lists and the xlsx writer and reader take a cell's look from here.
 */
import type { CellInput, ColumnInput, NumberFormat, RowInput, SheetInput } from '../contract.js'
import { formatAddress } from '../model.js'
import { entryAt, lines, type Line } from './spans.js'

/** A style (by index) and a number format, of a cell or a line. */
export interface Look {
  style?: number
  format?: NumberFormat
}

/** The format a cell shows by: its own, else its row's, else its column's. */
export const formatOf = (own: Look | undefined, row: Look | undefined, col: Look | undefined): NumberFormat | undefined => own?.format ?? row?.format ?? col?.format

/** The look a cell shows: its own style and format, each else its row's, else its column's (`formatOf`). */
export const lookOf = (own: Look | undefined, row: Look | undefined, col: Look | undefined): Look => ({ style: own?.style ?? row?.style ?? col?.style, format: formatOf(own, row, col) })

/** A sheet's rows and columns, each line looked up once. */
export interface SheetLines {
  row(r: number): RowInput | undefined
  column(c: number): ColumnInput | undefined
  /** The look a cell shows (`lookOf`): `own`, else the sheet's input at the cell. */
  look(r: number, c: number, own?: CellInput): Look
  /** The format a cell shows by (`formatOf`). */
  format(r: number, c: number, own?: CellInput): NumberFormat | undefined
}

export function sheetLines(sh: Pick<SheetInput, 'cells' | 'rows' | 'columns'>): SheetLines {
  const rows = lines(sh.rows, 'r')
  const cols = lines(sh.columns, 'c')
  const lineOf = <T>(cache: Map<number, T | undefined>, ls: Line<T>[], i: number): T | undefined => {
    if (!ls.length) return undefined
    if (!cache.has(i)) cache.set(i, entryAt(ls, i))
    return cache.get(i)
  }
  const [rowCache, colCache] = [new Map<number, RowInput | undefined>(), new Map<number, ColumnInput | undefined>()]
  const row = (r: number) => lineOf(rowCache, rows, r)
  const column = (c: number) => lineOf(colCache, cols, c)
  const at = (r: number, c: number) => sh.cells[formatAddress(r, c)]
  return { row, column, look: (r, c, own = at(r, c)) => lookOf(own, row(r), column(c)), format: (r, c, own = at(r, c)) => formatOf(own, row(r), column(c)) }
}
