/**
 * What the cells of a book show (`CellResult.text`, `color`): a cell's value by its format (its own, else its
 * row's, else its column's), in the book's locale, in the room of its column (of a merged area's columns from its
 * top-left cell), as `numfmt.ts` shows it.
 */
import type { BookInput, CellDisplay, CellInput, CellValue, NumberFormat } from './contract.js'
import type { EngineState } from './engine.js'
import { sheetLines, type SheetLines } from './layout/lines.js'
import type { Rules } from './locale/rules.js'
import { parseMerge } from './layout/merges.js'
import { columnPoints, room } from './numfmt.js'

/** What the cell at sheet `s`, row `r`, column `c` (zero-based) shows when it holds `v`; `own`, its input when the caller has it. */
export type Looks = (s: number, r: number, c: number, v: CellValue, own?: CellInput) => CellDisplay

/** The looks of cells of the state's book as it is now; a new one after a change of formats or layout. */
export const cellLooks = (state: EngineState): Looks => bookLooks(state.input, state.ctx.rules, state)

/** A sheet's lines, and its merged areas by `r:c` of their top-left cell: their last column. */
interface SheetLook {
  lines: SheetLines
  merges: Map<string, number>
}

/** Each sheet's lines and merged areas, read when the sheet is first looked at; with a session's `state`, its lines read once for it. */
function sheetLooks(book: BookInput, state?: EngineState): (s: number) => SheetLook {
  const sheets = new Map<number, SheetLook>()
  return (s) => {
    let x = sheets.get(s)
    if (!x) {
      const sh = book.sheets[s]
      const merges = new Map((sh.merges ?? []).map((m) => parseMerge(m)!).map((m): [string, number] => [`${m.r1}:${m.c1}`, m.c2]))
      sheets.set(s, (x = { lines: sheetLines(sh, state), merges }))
    }
    return x
  }
}

/** A cell's room in points: its column's, a merged area's columns' from its top-left cell. */
function pointsAt(x: SheetLook, r: number, c: number): number {
  const last = x.merges.size ? (x.merges.get(`${r}:${c}`) ?? c) : c
  let points = 0
  for (let k = c; k <= last; k++) points += columnPoints(x.lines.column(k)?.width)
  return points
}

/** The looks of cells of `book` by `rules` (a saved book's, without a session); with a session's `state`, its lines read once for it. */
export function bookLooks(book: BookInput, rules: Rules, state?: EngineState): Looks {
  const sheetLook = sheetLooks(book, state)
  // a cell without `own` shows by the sheet's input there, or its look of the sheet's looks (`sheetLines`)
  return (s, r, c, v, own) => {
    const x = sheetLook(s)
    const format = x.lines.format(r, c, own)
    return rules.display(v, format, room(pointsAt(x, r, c)))
  }
}

/**
 * What the cells of the state's book show by beside their value: the format and the room in points. Two cells of one
 * value with equal parts show the same (`bookLooks`), so their texts need not be made to be compared.
 */
export function lookParts(state: EngineState): (s: number, r: number, c: number, own?: CellInput) => [NumberFormat | undefined, number] {
  const sheetLook = sheetLooks(state.input, state)
  return (s, r, c, own) => {
    const x = sheetLook(s)
    return [x.lines.format(r, c, own), pointsAt(x, r, c)]
  }
}
