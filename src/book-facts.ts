/**
 * What CELL, INFO, SHEET, SHEETS, FORMULATEXT, ISFORMULA and GETPIVOTDATA read of a book beside its values (`Cx.book`):
 * its sheets, the formula a cell holds as Excel's formula bar shows it, a cell's look and its column's width, the file it
 * was opened from (`EngineOptions.file`), the active sheet's window and its PivotTables.
 */
import type { BookFacts, CellFacts } from './args.js'
import { charsOfPoints, columnPoints } from './column-width.js'
import type { BookFile } from './contract.js'
import { sheetIndexOf, type EngineState } from './engine.js'
import { sheetLines, type SheetLines } from './layout/lines.js'
import { quoteSheetIfNeeded } from './a1.js'
import { closingQuote } from './syntax/scan.js'
import { formatAddress, parseAddress, placeOf } from './model.js'
import { enCode } from './numfmt.js'
import { parseRange } from './structure/area.js'
import { pivotArea } from './pivot/area.js'

/**
 * A formula's canonical text with its sheets quoted only where Excel quotes them (`'Sales'!A1` is `Sales!A1`), outside
 * text in quotes; a linked workbook's sheet keeps its quotes. FORMULATEXT shows en-US's text in every locale.
 */
function shownSheets(text: string): string {
  let out = ''
  for (let i = 0; i < text.length; ) {
    const q = text[i]
    const j = q === '"' || q === "'" ? closingQuote(text, i) : -1
    if (j < 0) {
      out += text[i++]
      continue
    }
    const name = text.slice(i + 1, j).replaceAll("''", "'")
    const plain = q === "'" && text[j + 1] === '!' && !name.includes('[') && name.split(':').every((x) => x && quoteSheetIfNeeded(x) === x)
    out += plain ? name : text.slice(i, j + 1)
    i = j + 1
  }
  return out
}

/** The book's facts over a state, read as they are when asked: a session changes them in place. */
export function bookFacts(state: EngineState, file: BookFile | undefined): BookFacts {
  // a sheet's lines are read once for the state's input as it is (a change of looks or layout gives the sheet new ones)
  const lines = new WeakMap<object, SheetLines>()
  const linesOf = (s: number) => {
    const sh = state.input.sheets[s]
    let x = lines.get(sh)
    if (!x) lines.set(sh, (x = sheetLines(sh, state)))
    return x
  }
  // the old kind of array formula whose area holds the cell: each of its cells shows its formula
  const arrayAt = (s: number, r: number, c: number) => {
    const id = state.at.get(placeOf(s, r, c))
    if (id !== undefined || !state.arrays) return id
    return state.formulas.find((f) => f?.area && f.s === s && r >= f.area.r1 && r <= f.area.r2 && c >= f.area.c1 && c <= f.area.c2)?.id
  }
  const textAt = (s: number, r: number, c: number): string | undefined => {
    const id = arrayAt(s, r, c)
    if (id !== undefined) return state.formulas[id]!.text
    return s < state.sheetNames.length ? state.input.sheets[s].cells[formatAddress(r, c)]?.formula : undefined
  }
  const active = () => {
    const name = state.input.activeSheet
    const s = name === undefined ? state.input.sheets.findIndex((x) => !x.hidden) : sheetIndexOf(state, name)!
    return { s, view: state.input.sheets[s].view }
  }
  return {
    get sheets() {
      return state.sheetNames.length
    },
    sheetOf: (name) => sheetIndexOf(state, name),
    sheetName: (s) => state.sheetNames[s],
    hasFormula: (s, r, c) => textAt(s, r, c) !== undefined,
    formula: (s, r, c) => {
      const text = textAt(s, r, c)
      if (text === undefined) return undefined
      const shown = `=${shownSheets(text)}`
      const id = arrayAt(s, r, c)
      return id !== undefined && state.formulas[id]!.area ? `{${shown}}` : shown
    },
    cell: (s, r, c): CellFacts => {
      const x = linesOf(s)
      const look = x.look(r, c)
      const style = look.style === undefined ? undefined : state.input.styles?.[look.style]
      const width = x.column(c)?.width
      return {
        format: look.numberFormat === undefined ? 'General' : enCode(look.numberFormat),
        horizontal: style?.alignment?.horizontal,
        locked: style?.protection?.locked !== false,
        width: charsOfPoints(columnPoints(width)),
        defaultWidth: width === undefined,
      }
    },
    file,
    manual: state.input.calculation === 'manual',
    // the active cell where the selection holds it, else the selection's first cell, as Excel opens a file (`<selection>`'s
    // `sqref` is A1 when not given): the set `indirect-held-2`
    activeCell: () => {
      const { s, view } = active()
      const sel = (view?.selection ?? []).flatMap((x) => parseRange(x) ?? [])
      if (!sel.length) sel.push({ r1: 0, c1: 0, r2: 0, c2: 0 })
      const at = view?.activeCell ? parseAddress(view.activeCell) : undefined
      const inside = at && sel.some((x) => at.row >= x.r1 && at.row <= x.r2 && at.col >= x.c1 && at.col <= x.c2)
      return inside ? { s, r: at.row, c: at.col } : { s, r: sel[0].r1, c: sel[0].c1 }
    },
    pivotAt: (s, r, c) => {
      for (const table of state.input.sheets[s]?.pivotTables ?? []) {
        const area = pivotArea(table)
        // the cell above the top left of a PivotTable without a Filters area is the PivotTable's (the set gpd-held-2)
        const above = area && !table.location.pageRows && r === area.r1 - 1 && c === area.c1
        if (area && (above || (r >= area.r1 && r <= area.r2 && c >= area.c1 && c <= area.c2))) return { table, cache: state.input.pivotCaches?.find((x) => x.id === table.cache), area }
      }
      return undefined
    },
    topLeft: () => {
      const at = active().view?.topLeft
      const p = at ? parseAddress(at) : { row: 0, col: 0 }
      return { r: p.row, c: p.col }
    },
  }
}
