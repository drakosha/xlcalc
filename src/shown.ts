/**
 * What the cells of a book show (`CellResult.text`, `color`): a cell's value by its format (its own, else its
 * row's, else its column's), in the book's locale, in the room of its column (of a merged area's columns from its
 * top-left cell), as `numfmt.ts` shows it.
 */
import type { CellDisplay, CellInput, CellValue, ColumnInput, RowInput } from './contract.js'
import type { EngineState } from './engine.js'
import { parseMerge } from './layout/merges.js'
import { entryAt, lines, type Line } from './layout/spans.js'
import { formatAddress } from './model.js'
import { columnPoints, room } from './numfmt.js'

/** A sheet's layout, with the lines found so far. */
interface SheetLook {
  rows: Line<RowInput>[]
  cols: Line<ColumnInput>[]
  /** Merged areas by `r:c` of their top-left cell: their last column. */
  merges: Map<string, number>
  row: Map<number, RowInput | undefined>
  col: Map<number, ColumnInput | undefined>
}

/** What the cell at sheet `s`, row `r`, column `c` (zero-based) shows when it holds `v`; `own`, its input when the caller has it. */
export type Looks = (s: number, r: number, c: number, v: CellValue, own?: CellInput) => CellDisplay

/** The looks of cells of the state's book as it is now; a new one after a change of formats or layout. */
export function cellLooks(state: EngineState): Looks {
  const sheets = new Map<number, SheetLook>()
  const sheetLook = (s: number): SheetLook => {
    let x = sheets.get(s)
    if (!x) {
      const sh = state.input.sheets[s]
      const merges = new Map((sh.merges ?? []).map((m) => parseMerge(m)!).map((m): [string, number] => [`${m.r1}:${m.c1}`, m.c2]))
      sheets.set(s, (x = { rows: lines(sh.rows, 'r'), cols: lines(sh.columns, 'c'), merges, row: new Map(), col: new Map() }))
    }
    return x
  }
  const lineOf = <T>(cache: Map<number, T | undefined>, ls: Line<T>[], i: number): T | undefined => {
    if (!cache.has(i)) cache.set(i, entryAt(ls, i))
    return cache.get(i)
  }
  return (s, r, c, v, own = state.input.sheets[s].cells[formatAddress(r, c)]) => {
    const x = sheetLook(s)
    const format = own?.format ?? (x.rows.length ? lineOf(x.row, x.rows, r)?.format : undefined) ?? (x.cols.length ? lineOf(x.col, x.cols, c)?.format : undefined)
    const last = x.merges.size ? (x.merges.get(`${r}:${c}`) ?? c) : c
    let points = 0
    for (let k = c; k <= last; k++) points += columnPoints(x.cols.length ? lineOf(x.col, x.cols, k)?.width : undefined)
    return state.ctx.rules.display(v, format, room(points))
  }
}
