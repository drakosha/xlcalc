/**
 * Held out: the cases of `indirect-held` where Excel for Mac answered `#NUM!` for INDIRECT and CELL, in books of their
 * own, and CELL without a reference in a book saved with its active cell selected (`selection` as well as
 * `activeCell`), predicted by the engine before Excel computed them (`indirect-held-2.predicted.json`).
 */
import { buildTableCases as held } from './indirect-held.js'
import type { TableCase } from './cases.js'

const AGAIN = ['names-typed', 'table', 'typed', 'typed-formula', 'insert-rows', 'delete-rows']

export function buildTableCases(): TableCase[] {
  const cases = held()
  const out = cases.filter((c) => AGAIN.includes(c.id)).map((c) => ({ ...c, id: `${c.id}-again` }))
  const active = cases.find((c) => c.id === 'active')!
  const book = { ...active.book, sheets: active.book.sheets.map((s) => (s.name === 'T' ? { ...s, view: { activeCell: 'D4', selection: ['D4'] } } : s)) }
  out.push({ ...active, id: 'active-selected', about: 'CELL without a reference, the book saved with T!D4 active and selected', book })
  return out
}
