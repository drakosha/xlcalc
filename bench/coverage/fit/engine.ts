/**
 * The engine's answer to a fit case, the same for the predictions written before Excel runs a set and for the tests:
 * the case's book computed, its column fitted (`fitColumns`) or the texts typed in turn (each read by `parseEntry`, the
 * book computed with it, the column widened by `entryWidth` as a host does), then the column's width and what its cells
 * show at that width. The engine's functions come as arguments, so a test passes the ones it tests (`src` or `dist`).
 */
import type { BookInput, CellInput, CellValue, ColumnInput, ComputeResult, EntryOptions, EntryResult, EntryWidthOptions } from '../../../src/contract.js'
import { caseRows, fitBook, type FitCase } from './cases.js'
import { numberFormatOf } from '../stored.js'

export interface FitEngine {
  compute(book: BookInput): ComputeResult
  parseEntry(text: string, opts: EntryOptions): EntryResult
  fitColumns(book: BookInput, result: ComputeResult, sheet: string, columns: string): Record<string, number | null>
  entryWidth(value: CellValue, opts: EntryWidthOptions): number | null | undefined
}

/** The engine's answer: the column's width (null for the default), whether AutoFit or typing set it, and what each of the case's rows shows. */
export interface FitPrediction {
  width: number | null
  fitted?: boolean
  texts: Record<number, string>
}

const COL = 'B'

/** The case's answer; `widen` false for texts typed through an object model, which on the Mac widens no column. */
export function engineFit(e: FitEngine, c: FitCase, tag: string, now: string, widen = true): FitPrediction {
  const book = fitBook([c])
  const sheet = book.sheets[0]
  book.locale = tag
  const columns = sheet.columns!
  let fitted = false
  const setWidth = (width: number | null | undefined, more: ColumnInput = {}) => {
    if (width === undefined) return
    fitted = true
    const { width: _w, bestFit: _b, hidden: _h, ...rest } = columns[COL] ?? {}
    const col: ColumnInput = { ...rest, ...(width === null ? {} : { width, bestFit: true }), ...(more.hidden ? { hidden: true } : {}) }
    if (Object.keys(col).length) columns[COL] = col
    else delete columns[COL]
  }
  if (!c.typed || c.fitFirst) setWidth(e.fitColumns(book, e.compute(book), 'S', COL)[COL])
  for (const [row, text] of c.typed ?? []) {
    const addr = `${COL}${row}`
    const before = sheet.cells[addr]
    const r = e.parseEntry(text, { locale: tag, now, ...(before ? { cell: before } : {}) })
    if (!r.ok) continue
    if (r.input) sheet.cells[addr] = r.input as CellInput
    else delete sheet.cells[addr]
    // a host widens no column for a turned cell (`entryWidth`)
    const style = sheet.cells[addr]?.style
    if (!widen || (style !== undefined && book.styles?.[style]?.alignment?.rotation)) continue
    const value = e.compute(book).sheets[0].cells[addr]?.value ?? null
    const col = columns[COL] ?? {}
    const numberFormat = numberFormatOf(sheet.cells[addr]) ?? numberFormatOf(col)
    setWidth(e.entryWidth(value, { locale: tag, ...(numberFormat !== undefined ? { numberFormat } : {}), ...(col.width !== undefined ? { width: col.width } : {}), ...(col.bestFit ? { bestFit: true } : {}) }), col)
  }
  const cells = e.compute(book).sheets[0].cells
  return { width: columns[COL]?.width ?? null, ...(fitted ? { fitted } : {}), texts: Object.fromEntries(caseRows(c).map((r) => [r, cells[`${COL}${r}`]?.text ?? ''])) }
}
