/**
 * The engine's answer to a rows case, the same for the predictions written before Excel runs a set and for the tests:
 * the case's book computed, its texts typed in turn (each read by `parseEntry`), its row fitted (`fitRows`) when the
 * case fits it, then the row's height before and after and whether its cells wrap. The engine's functions come as
 * arguments, so a test passes the ones it tests (`src` or `dist`).
 */
import type { BookInput, CellInput, ComputeResult, EntryOptions, EntryResult } from '../../../src/contract.js'
import { caseCols, caseRow, rowsBook, type RowCase } from './cases.js'

export interface RowsEngine {
  compute(book: BookInput): ComputeResult
  parseEntry(text: string, opts: EntryOptions): EntryResult
  fitRows(book: BookInput, result: ComputeResult, sheet: string, rows: string): Record<string, number | null>
}

/** The engine's answer: the row's height when the book opens and after (the default 15 for none, 0 for a hidden row) and whether each cell wraps. */
export interface RowsPrediction {
  before: number
  height: number
  wrap: Record<number, boolean>
}

const DEFAULT = 15
const COLS = 'BCD'

/** The case's answer. */
export function engineRows(e: RowsEngine, c: RowCase, tag: string, now: string): RowsPrediction {
  const book = rowsBook([c])
  book.locale = tag
  const sheet = book.sheets[0]
  const row = String(caseRow(0))
  const shown = () => {
    const r = sheet.rows?.[row]
    return r?.hidden ? 0 : (r?.height ?? DEFAULT)
  }
  const before = shown()
  for (const [col, text] of c.typed ?? []) {
    const addr = `${COLS[col]}${row}`
    const was = sheet.cells[addr]
    const r = e.parseEntry(text, { locale: tag, now, ...(was ? { cell: was } : {}) })
    if (!r.ok) continue
    if (r.input) sheet.cells[addr] = r.input as CellInput
    else delete sheet.cells[addr]
  }
  if (!c.typed || c.fitAfter) {
    const height = e.fitRows(book, e.compute(book), 'S', row)[row]
    const { height: _h, hidden: _x, ...rest } = sheet.rows?.[row] ?? {}
    sheet.rows = { ...sheet.rows, [row]: { ...rest, ...(height === null ? {} : { height }) } }
  }
  const styles = book.styles ?? []
  const wrap = Object.fromEntries(caseCols(c).map((col) => [col, !!styles[sheet.cells[`${COLS[col]}${row}`]?.style ?? -1]?.alignment?.wrap]))
  return { before, height: shown(), wrap }
}
