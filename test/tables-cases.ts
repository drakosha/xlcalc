/**
 * A table case of `bench/coverage/tables/` taken by the engine as Excel took it: the book opened in a session, each
 * step applied (a cell typed through `setCells`, an operation through `restructure`, a refusal noted), and the answer
 * read as the harness reads Excel's: the rectangles' values, their formulas as stored text, the tables.
 */
import type { Val } from '../bench/coverage/formulajs.js'
import { cellsOf } from '../bench/coverage/case-book.js'
import type { TableCase, TableShape } from '../bench/coverage/tables/cases.js'
import { InputError, openSession, type BookInput, type ComputeResult } from '../src/index.js'
import { parseAddress } from '../src/model.js'
import { sheetKey, validate } from '../src/syntax.js'
import { tableAt, tablesOf } from '../src/tables/model.js'
import { shownValue } from './spill-shown.js'

export const CLOCK = { now: '2026-10-07T12:00:00Z', seed: 1 }

export interface EngineAnswer {
  values: Record<string, Val>
  /** Each formula of the read rectangles, as stored (canonical) text. */
  formulas: Record<string, string>
  tables: TableShape[]
  /** The steps the engine refused, by index: its message. */
  refused: Record<number, string>
  book: BookInput
  result: ComputeResult
}

/** The tables of a book as Excel lists them. */
export const shapes = (book: BookInput): TableShape[] =>
  book.sheets.flatMap((s) => (s.tables ?? []).map((t) => ({ sheet: s.name, name: t.name, range: t.range.includes(':') ? t.range : `${t.range}:${t.range}`, headers: t.headers !== false, totals: t.totals === true, columns: t.columns.map((c) => c.name) })))

export function engineAnswer(c: TableCase): EngineAnswer {
  const s = openSession(c.book, { clock: CLOCK })
  const refused: Record<number, string> = {}
  ;(c.steps ?? []).forEach((step, i) => {
    if ('op' in step) {
      const r = s.restructure(step.op)
      if (!r.ok) refused[i] = `${r.code}: ${r.message}`
      return
    }
    const { sheet, cell, value: v } = step.edit
    const input = v === null ? null : typeof v === 'object' ? { formula: v.formula } : { value: v }
    try {
      s.setCells(cellsOf(cell).addrs.map((a) => ({ sheet, cell: a, input })))
    } catch (e) {
      if (!(e instanceof InputError)) throw e
      refused[i] = e.message
    }
  })
  const book = s.book()
  const result = s.result()
  const values: Record<string, Val> = {}
  const formulas: Record<string, string> = {}
  for (const g of c.read) {
    const sheet = result.sheets.find((x) => sheetKey(x.name) === sheetKey(g.sheet))
    if (!sheet) continue
    for (const a of cellsOf(g.range).addrs) {
      const v = shownValue(result, sheet.name, a)
      if (v) values[`${g.sheet}!${a}`] = v
      const f = sheet.cells[a]?.formula
      if (f !== undefined) formulas[`${g.sheet}!${a}`] = f
    }
  }
  s.close()
  return { values, formulas, tables: shapes(book), refused, book, result }
}

/** Excel's formula text (`=[@Qty]*2`, as Excel shows it) at a cell of the book as stored text, read with the book's tables. */
export function storedText(book: BookInput, sheet: string, addr: string, shown: string): string {
  const tables = tablesOf(book)
  const s = book.sheets.findIndex((x) => sheetKey(x.name) === sheetKey(sheet))
  const { row, col } = parseAddress(addr)
  const enclosing = tableAt(tables, s, row, col)?.name
  const r = validate(shown.replace(/^=/, ''), { sheets: book.sheets.map((x) => x.name), tables: tables.list.map((t) => ({ name: t.name, columns: t.cols })), ...(enclosing ? { enclosing } : {}) })
  return r.ok ? r.canonical : `(${r.code}) ${shown}`
}
