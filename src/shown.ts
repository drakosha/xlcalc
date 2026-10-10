/**
 * What the cells of a book show (`CellResult.text`, `color`): a cell's value by its format (its own, else its
 * row's, else its column's), in the book's locale, in the room of its column (of a merged area's columns from its
 * top-left cell), as `numfmt.ts` shows it.
 */
import type { CellDisplay, CellInput, CellValue, ColumnInput, NumberFormat, RowInput, SheetInput } from './contract.js'
import type { EngineState } from './engine.js'
import { parseMerge } from './layout/merges.js'
import { entryAt, lines, type Line } from './layout/spans.js'
import { formatAddress } from './model.js'
import { columnPoints, room } from './numfmt.js'

/** A sheet's rows and columns, each line looked up once. */
export interface SheetLines {
  row(r: number): RowInput | undefined
  column(c: number): ColumnInput | undefined
  /** The format a cell shows by: its own (`own`, else the sheet's input at the cell), else its row's, else its column's. */
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
  return { row, column, format: (r, c, own = sh.cells[formatAddress(r, c)]) => own?.format ?? row(r)?.format ?? column(c)?.format }
}

/** What the cell at sheet `s`, row `r`, column `c` (zero-based) shows when it holds `v`; `own`, its input when the caller has it. */
export type Looks = (s: number, r: number, c: number, v: CellValue, own?: CellInput) => CellDisplay

/** The looks of cells of the state's book as it is now; a new one after a change of formats or layout. */
export function cellLooks(state: EngineState): Looks {
  // each sheet's lines, and its merged areas by `r:c` of their top-left cell: their last column
  const sheets = new Map<number, { lines: SheetLines; merges: Map<string, number> }>()
  const sheetLook = (s: number) => {
    let x = sheets.get(s)
    if (!x) {
      const sh = state.input.sheets[s]
      const merges = new Map((sh.merges ?? []).map((m) => parseMerge(m)!).map((m): [string, number] => [`${m.r1}:${m.c1}`, m.c2]))
      sheets.set(s, (x = { lines: sheetLines(sh), merges }))
    }
    return x
  }
  return (s, r, c, v, own = state.input.sheets[s].cells[formatAddress(r, c)]) => {
    const x = sheetLook(s)
    const format = x.lines.format(r, c, own)
    const last = x.merges.size ? (x.merges.get(`${r}:${c}`) ?? c) : c
    let points = 0
    for (let k = c; k <= last; k++) points += columnPoints(x.lines.column(k)?.width)
    return state.ctx.rules.display(v, format, room(points))
  }
}
