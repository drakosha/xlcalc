/**
 * What the cells of a book show (`CellResult.text`, `color`): a cell's value by its format (its own, else its
 * row's, else its column's), in the book's locale, in the room of its column (of a merged area's columns from its
 * top-left cell), as `numfmt.ts` shows it.
 */
import type { BookInput, CellDisplay, CellInput, CellValue } from './contract.js'
import type { EngineState } from './engine.js'
import { sheetLines, type SheetLines } from './layout/lines.js'
import type { Rules } from './locale/rules.js'
import { parseMerge } from './layout/merges.js'
import { columnPoints, room } from './numfmt.js'

/** What the cell at sheet `s`, row `r`, column `c` (zero-based) shows when it holds `v`; `own`, its input when the caller has it. */
export type Looks = (s: number, r: number, c: number, v: CellValue, own?: CellInput) => CellDisplay

/** The looks of cells of the state's book as it is now; a new one after a change of formats or layout. */
export const cellLooks = (state: EngineState): Looks => bookLooks(state.input, state.ctx.rules, state)

/** The looks of cells of `book` by `rules` (a saved book's, without a session); with a session's `state`, its lines read once for it. */
export function bookLooks(book: BookInput, rules: Rules, state?: EngineState): Looks {
  // each sheet's lines, and its merged areas by `r:c` of their top-left cell: their last column
  const sheets = new Map<number, { lines: SheetLines; merges: Map<string, number> }>()
  const sheetLook = (s: number) => {
    let x = sheets.get(s)
    if (!x) {
      const sh = book.sheets[s]
      const merges = new Map((sh.merges ?? []).map((m) => parseMerge(m)!).map((m): [string, number] => [`${m.r1}:${m.c1}`, m.c2]))
      sheets.set(s, (x = { lines: sheetLines(sh, state), merges }))
    }
    return x
  }
  // a cell without `own` shows by the sheet's input there, or its look of the sheet's looks (`sheetLines`)
  return (s, r, c, v, own) => {
    const x = sheetLook(s)
    const format = x.lines.format(r, c, own)
    const last = x.merges.size ? (x.merges.get(`${r}:${c}`) ?? c) : c
    let points = 0
    for (let k = c; k <= last; k++) points += columnPoints(x.lines.column(k)?.width)
    return rules.display(v, format, room(points))
  }
}
