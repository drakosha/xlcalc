/**
 * Held out, in Excel ru-RU (`run.ts tables headers-ru-held ru`): numbers typed into a table's header cell, predicted by
 * the engine before Excel took them (`headers-ru-held.predicted.json`, from `predict.ts`). The column's name and the
 * formulas over it, in English and in Russian: a decimal (`1,5`), a whole, a negative, a large and a small number, one
 * typed over a header cell with a number format of its own (`0.00`, the short date), text that Excel ru-RU reads as a
 * number, a date or a boolean, and formulas typed in Russian over a column named `1,5`.
 */
import type { NumberFormat } from '../../../src/contract.js'
import { type CellSpec, on, sales, type TableCase, type TableStep, typedLocal } from './cases.js'
import type { StoredBook, StoredCell } from '../stored.js'

export const local = true

const OVER: Record<string, CellSpec> = { J1: '=SUM(Sales[Qty])', J2: '=SUM(Sales[[Region]:[Price]])', J3: '=SUM(Sales[Price])', J4: '=Sales[@Price]', J5: '=Sales[[#Headers],[Price]]', J6: '=ROWS(Sales[#All])' }
const READ = on('S', 'A1:K6')

/** Sales in a ru-RU book with the formulas over it; `head` gives the header cell C1 a number format of its own. */
function ru(head?: NumberFormat): StoredBook {
  const t = sales()
  if (head !== undefined) t.cells.C1 = { ...t.cells.C1, format: head }
  const cells: Record<string, StoredCell> = { ...t.cells, ...Object.fromEntries(Object.entries(OVER).map(([a, f]) => [a, { formula: String(f).slice(1) }])) }
  return { locale: 'ru-RU', sheets: [{ name: 'S', cells, tables: [t.table] }] }
}

const type = (cell: string, value: number | string): TableStep => ({ edit: { sheet: 'S', cell, value } })

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  const add = (id: string, about: string, steps: TableStep[], b = ru()) => cases.push({ id, about, book: b, steps, read: READ })
  add('decimal', 'a decimal typed into the header cell of Price: the column is named 1,5', [type('C1', 1.5)])
  add('whole', 'a whole number typed into a header cell', [type('C1', 2)])
  add('negative', 'a negative decimal typed into a header cell', [type('B1', -1.5)])
  add('large', 'a large number with decimals typed into a header cell', [type('C1', 1234567.25)])
  add('small', 'a small number typed into a header cell', [type('C1', 0.0001)])
  add('tiny', 'a number General shows with an exponent typed into a header cell', [type('C1', 1e-9)])
  add('decimal-format', 'a decimal typed into a header cell with the format 0.00', [type('C1', 1.5)], ru('0.00'))
  add('date-format', 'a date as a number typed into a header cell with the short date format', [type('C1', 45688)], ru(14))
  add('text-date', 'a date typed as text, 31.01.2025', [type('C1', '31.01.2025')])
  add('text-iso', 'a date typed as text, 2025-01-31', [type('C1', '2025-01-31')])
  add('text-decimal', 'a decimal typed as text, 1,5', [type('C1', '1,5')])
  add('text-bool', 'ИСТИНА typed as text', [type('C1', 'ИСТИНА')])
  add('typed-formulas', 'formulas typed in Russian over the column named 1,5', [type('C1', 1.5), typedLocal('K1', '=СУММ(Sales[[1,5]])'), typedLocal('K2', '=Sales[@[1,5]]*2'), typedLocal('K3', '=СУММ(Sales[[Qty]:[1,5]])')])
  return cases
}
