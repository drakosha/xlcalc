/**
 * Formats of rows and columns and the cells under them, for Excel (`run.ts lines <set>`, `bench/windows/run.ts lines <set>`).
 * A case's book is written by `writeXlsx`; Excel opens it, takes the case's steps through its object model (a number
 * format, bold or a fill set on a column, a row or a cell, a value typed into a cell), reads what the listed cells show
 * (`Range.Text`, the number format, the display format's fill and font) and saves the book. The answer is what Excel
 * showed; the file it saved shows how Excel stores the cells of formatted lines, and the reader reads it back. A bare
 * case's file is left without the cells the writer writes where a row with a look crosses a column with one, as other
 * writers leave it: the engine's book of it is the one the reader reads.
 */
import type { BookInput, CellStyle, ColumnInput, RowInput, SheetInput } from '../../../src/contract.js'
import { openSession } from '../../../src/index.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { sheetLines } from '../../../src/layout/lines.js'
import { lineAt, patchLines, parseSpan } from '../../../src/layout/spans.js'
import { parseAddress } from '../../../src/model.js'
import { BUILTIN } from '../../../src/xlsx/styles/numfmt.js'
import { cellsOf, viewLook, type Look } from '../cond/cases.js'
import { editParts } from '../xlsx/sent.js'

/** A step Excel takes through its object model on a column (`B:B`), a row (`4:4`) or a cell: formats set, or a value typed. */
export interface LineStep {
  sheet: string
  range: string
  /** A number format code, en-US. */
  format?: string
  bold?: true
  italic?: true
  /** A solid fill, `#RRGGBB`. */
  fill?: string
  /** A value typed into a cell. */
  value?: number | string
}

export interface LinesCase {
  id: string
  /** The key of the case's book in the set's `books`. */
  book: string
  /** The file without the empty cells the writer writes where lines cross. */
  bare?: true
  steps?: LineStep[]
  /** The cells read: a sheet and a range of it. */
  read: [sheet: string, range: string][]
}

export interface LinesCases {
  books: Record<string, BookInput>
  cases: LinesCase[]
}

export interface LinesSet {
  buildLinesCases(): LinesCases
}

/** What a cell shows (`Look`: the text, the fill, the font's colour, bold, italic) and its number format as Excel names it. */
export interface LineLook extends Look {
  format: string
}

export interface LinesAnswer {
  /** Excel's error when it refused a step. */
  error?: string
  /** The looks of the read cells, by sheet and address. */
  looks: Record<string, Record<string, LineLook>>
}

type LineInput = RowInput & ColumnInput

/** A style with what a step sets added (bold, italic, a fill): the book's styles grow by it unless an equal one is there. */
function withStep(book: BookInput, style: number | undefined, step: LineStep): number | undefined {
  if (!step.bold && !step.italic && !step.fill) return style
  const styles = (book.styles ??= [])
  const base: CellStyle = style === undefined ? {} : styles[style]
  const font = { ...base.font, ...(step.bold ? { bold: true } : {}), ...(step.italic ? { italic: true } : {}) }
  const next: CellStyle = { ...base, ...(step.bold || step.italic ? { font } : {}), ...(step.fill ? { fill: { color: step.fill } } : {}) }
  const key = JSON.stringify(next)
  const i = styles.findIndex((s) => JSON.stringify(s) === key)
  return i >= 0 ? i : styles.push(next) - 1
}

/**
 * The book after the steps, as the model takes them: a line's style and format set (its style with the step's bold or
 * fill added), a cell's likewise, a value typed into a cell. Excel also formats the cells a line's step covers; the
 * sets format lines before any cell under them has a look of its own.
 */
export function applySteps(book: BookInput, steps: LineStep[] = []): BookInput {
  const out: BookInput = structuredClone(book)
  for (const step of steps) {
    const sh = out.sheets.find((s) => s.name === step.sheet)!
    const line = /^([A-Z]+):([A-Z]+)$/.exec(step.range) ?? /^(\d+):(\d+)$/.exec(step.range)
    if (line) {
      const axis = /\d/.test(line[1]) ? 'r' : 'c'
      const key = axis === 'r' ? 'rows' : 'columns'
      const [lo, hi] = parseSpan(line[1] === line[2] ? line[1] : step.range, axis)!
      const map = sh[key] as Record<string, LineInput> | undefined
      const style = withStep(out, lineAt(map, axis, lo)?.style, step)
      Object.assign(sh, { [key]: patchLines(map, axis, lo, hi, { ...(style !== undefined ? { style } : {}), ...(step.format !== undefined ? { format: step.format } : {}) }) })
      continue
    }
    const cell = (sh.cells[step.range] ??= {})
    const style = withStep(out, cell.style, step)
    if (style !== undefined) cell.style = style
    if (step.format !== undefined) cell.format = step.format
    if (step.value !== undefined) cell.value = step.value
  }
  return out
}

/** A number format as Excel names it: General for none, a built-in one by its code. */
export const formatName = (f: string | number | undefined): string => (f === undefined || f === 0 ? 'General' : typeof f === 'number' ? BUILTIN[f] : f)

/** What the cells of `range` on a sheet show in a computed session's book, each by the whole look its row or column gives it. */
export function sheetLooks(book: BookInput, sheet: string, range: string): Record<string, LineLook> {
  const sh: SheetInput = book.sheets.find((s) => s.name === sheet)!
  const s = openSession(book)
  const views = s.cells(sheet, range)
  s.close()
  const at = sheetLines(sh)
  return Object.fromEntries(
    cellsOf(range).map((a) => {
      const { row, col } = parseAddress(a)
      const look = at.look(row, col)
      return [a, { ...viewLook({ ...(views[a] ?? { value: null }), style: look.style }, book.styles), format: formatName(look.format) }]
    }),
  )
}

/** The file of a case: its book as `writeXlsx` writes it; for a bare case without the empty cells no cell of the book is at. */
export async function caseFile(set: LinesCases, c: LinesCase): Promise<Uint8Array> {
  const book = set.books[c.book]
  const data = (await writeXlsx(book)).data
  if (!c.bare) return data
  return editParts(data, (name, xml) => {
    const sh = /^xl\/worksheets\/sheet(\d+)\.xml$/.exec(name)
    const cells = sh && book.sheets[Number(sh[1]) - 1].cells
    return cells ? xml.replace(/<c r="([A-Z]+\d+)" s="\d+"\/>/g, (all, addr: string) => (cells[addr] ? all : '')) : undefined
  })
}

/** The book the engine computes for a case: the set's, or for a bare case the one the reader reads from its file. */
export const caseBook = async (set: LinesCases, c: LinesCase): Promise<BookInput> => (c.bare ? (await readXlsx(await caseFile(set, c))).book : set.books[c.book])

/** The engine's answer to a case on its book (`caseBook`): what the read cells show after its steps. */
export function engineCase(c: LinesCase, book: BookInput): LinesAnswer['looks'] {
  const after = applySteps(book, c.steps)
  return Object.fromEntries(c.read.map(([sheet, range]) => [sheet, sheetLooks(after, sheet, range)]))
}
