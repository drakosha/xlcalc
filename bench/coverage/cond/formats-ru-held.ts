/**
 * Held out: a rule's number format in a book of ru-RU, Excel in ru-RU, predicted before Excel ran it
 * (`formats-ru-held.predicted.json`). The writer stores a built-in format given by its number in a differential format
 * with that number and its code (`<numFmt numFmtId="14" formatCode="m/d/yyyy"/>`), and the reader takes such a pair back
 * as the built-in format, which ru-RU shows by its own regional data (`15.03.2026`); the same code given as text is a
 * code of the book's own (`numFmtId="164"`), shown as written (`3/15/2026`). Dates, a date with a time, money and percent.
 */
import type { BookInput, CellInput, ConditionalFormat, NumberFormat } from '../../../src/contract.js'
import { BUILTIN } from '../../../src/xlsx/styles/numfmt.js'
import type { CondCases } from './cases.js'

/** Each rule's cell, its value and format: a built-in format by its number, then its code as text. */
const FORMATS: [cell: string, value: number, format: NumberFormat][] = [
  ['A1', 46096.75, 14],
  ['A2', 46096.75, BUILTIN[14]],
  ['A3', 46096.75, 22],
  ['A4', 46096.75, BUILTIN[22]],
  ['A5', 46096.75, 15],
  ['A6', 46096.75, BUILTIN[15]],
  ['A7', 1234.5, 5],
  ['A8', 1234.5, BUILTIN[5]],
  ['A9', 0.5, 10],
]

function book(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (const [cell, value] of FORMATS) cells[cell] = { value }
  const conditionalFormats: ConditionalFormat[] = FORMATS.map(([cell, , format]) => ({ type: 'cellIs', ranges: [cell], operator: 'greaterThan', formulas: ['0'], style: { format } }))
  return { locale: 'ru-RU', sheets: [{ name: 'Лист', cells, columns: { A: { width: 30 } }, conditionalFormats }] }
}

export function buildCondCases(): CondCases {
  return { locale: 'ru-RU', books: { formats: book() }, cases: [{ id: 'formats', book: 'formats', read: [['Лист', 'A1:A9']] }] }
}
