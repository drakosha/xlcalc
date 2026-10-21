/**
 * Text with a line break typed into cells (`setCellText`, `setCellTexts`, `fillCellText`) or pasted (`pasteText`), as
 * Excel takes it (`test/excel/rows/`): the cell wraps its text, its style taking `alignment.wrap` (a style of the book's
 * table found or added); a formula typed over lines does not. Typed, the row of a cell that wraps fits it as AutoFit
 * does, unless its height was set by hand or it is hidden; a paste fits no row.
 */
import type { CellInput, CellStyle } from '../contract.js'
import type { EngineState } from '../engine.js'
import { fitHeights, stateRows } from '../layout/heights.js'
import { sheetLines } from '../layout/lines.js'
import { parseSpan } from '../layout/spans.js'
import { StyleTable } from '../style-table.js'
import type { SetLayoutEdit } from './book.js'
import { patchStyle } from './styles.js'
import type { Typed } from './texts.js'

/** Whether typed text breaks a line, which makes Excel wrap the cell: text that is no formula (`input`, what it gives the cell). */
export const breaksLine = (text: string, input: CellInput | null): boolean => input?.formula === undefined && /[\r\n]/.test(text)

/**
 * The typed cells marked `wrap` given the style that wraps (their own, else their row's or column's, with `alignment.wrap`)
 * in their inputs; the book's style table when it grew.
 */
export function wrapTyped(s: EngineState, cells: Typed[]): CellStyle[] | undefined {
  if (!cells.some((x) => x.wrap)) return undefined
  const table = new StyleTable(s.input.styles)
  for (const x of cells) {
    if (!x.wrap || !x.edit.input) continue
    const lines = sheetLines(s.input.sheets[x.i], s)
    const shown = x.edit.input.style ?? lines.row(x.r)?.style ?? lines.column(x.c)?.style
    if (table.at(shown).alignment?.wrap) continue
    x.edit.input = { ...x.edit.input, style: table.indexOf(patchStyle(table.at(shown), { alignment: { wrap: true } })) }
  }
  return table.table
}

/** The heights the rows of the typed cells that wrap take now, as AutoFit gives them, where a row's height is the default or one Excel fitted and it shows. */
export function typedHeights(s: EngineState, cells: Typed[]): SetLayoutEdit[] {
  const styles = s.input.styles ?? []
  const rows = new Map<number, Set<number>>()
  const linesOf = new Map<number, ReturnType<typeof sheetLines>>()
  for (const { i, r, c } of cells) {
    let lines = linesOf.get(i)
    if (!lines) linesOf.set(i, (lines = sheetLines(s.input.sheets[i], s)))
    const style = lines.look(r, c).style
    const row = lines.row(r)
    if (style === undefined || !styles[style]?.alignment?.wrap || row?.hidden || (row?.height !== undefined && !row.bestFit)) continue
    let set = rows.get(i)
    if (!set) rows.set(i, (set = new Set()))
    set.add(r)
  }
  const out: SetLayoutEdit[] = []
  for (const [i, set] of rows) {
    // one pass, not a spread: a call can wrap more rows than a function call takes arguments
    let [lo, hi] = [Infinity, -Infinity]
    for (const r of set) [lo, hi] = [Math.min(lo, r), Math.max(hi, r)]
    const { cells: values, spans } = stateRows(s, i, lo, hi)
    const lines = sheetLines(s.input.sheets[i], s)
    for (const [key, height] of Object.entries(fitHeights(styles, s.input.sheets[i], values, lo, hi, spans))) {
      const [a, b] = parseSpan(key, 'r')!
      for (let r = a; r <= b; r++)
        if (set.has(r) && (lines.row(r)?.height ?? null) !== height) out.push({ sheet: s.sheetNames[i], rows: String(r + 1), set: { height, bestFit: height === null ? null : true } })
    }
  }
  return out
}

