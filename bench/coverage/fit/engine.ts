/**
 * The engine's answer to a fit case, the same for the predictions written before Excel runs a set and for the tests:
 * the case's book computed, its column fitted or the texts typed in turn (each read by `parseEntry`, the book computed
 * with it, the column widened as Excel widens it), then the column's width and what its cells show at that width. The
 * engine's functions come as arguments, so a test passes the ones it tests (`src` or the built `dist`).
 */
import type { BookInput, CellInput, ComputeResult, EntryOptions, EntryResult, NumberFormat } from '../../../src/contract.js'
import { caseRows, fitBook, type FitCase } from './cases.js'

export interface FitEngine {
  compute(book: BookInput): ComputeResult
  parseEntry(text: string, opts: EntryOptions): EntryResult
  /** AutoFit of column `col` of the book's sheet `S`: its width, null for the default, undefined when it keeps its own. */
  fit(book: BookInput, result: ComputeResult, col: string): number | null | undefined
  /** The width a column of `width` (`fitted` by Excel) gets as `value` shown by `format` is typed into it; undefined when it keeps it. */
  widen(value: ComputeResult["sheets"][0]["cells"][string]["value"], format: NumberFormat | undefined, tag: string, column: { width?: number; fitted?: boolean }): number | null | undefined
}

/** The engine's answer: the column's width (null for the default) and what each of the case's rows shows. */
export interface FitPrediction {
  width: number | null
  texts: Record<number, string>
}

const COL = 'B'

export function engineFit(e: FitEngine, c: FitCase, tag: string, now: string): FitPrediction {
  const book = fitBook([c])
  const sheet = book.sheets[0]
  book.locale = tag
  const state: { width?: number; fitted?: boolean } = c.column?.width !== undefined ? { width: c.column.width } : {}
  const setWidth = (w: number | null | undefined) => {
    if (w === undefined) return
    const { width: _, ...rest } = sheet.columns![COL] ?? {}
    sheet.columns![COL] = w === null ? rest : { ...rest, width: w }
    if (!Object.keys(sheet.columns![COL]).length) delete sheet.columns![COL]
    if (w === null) delete state.width
    else state.width = w
    state.fitted = w !== null
  }
  if (!c.typed || c.fitFirst) setWidth(e.fit(book, e.compute(book), COL))
  for (const [row, text] of c.typed ?? []) {
    const addr = `${COL}${row}`
    const before = sheet.cells[addr]
    const r = e.parseEntry(text, { locale: tag, now, ...(before ? { cell: before } : {}) })
    if (!r.ok) continue
    if (r.input) sheet.cells[addr] = r.input as CellInput
    else delete sheet.cells[addr]
    const value = e.compute(book).sheets[0].cells[addr]?.value ?? null
    setWidth(e.widen(value, sheet.cells[addr]?.format ?? sheet.columns![COL]?.format, tag, { ...state }))
  }
  const cells = e.compute(book).sheets[0].cells
  return { width: state.width ?? null, texts: Object.fromEntries(caseRows(c).map((r) => [r, cells[`${COL}${r}`]?.text ?? ''])) }
}
