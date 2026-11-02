/**
 * The samples of the Number Format list on Excel's Home tab: Excel for Windows shows under each item of the list the
 * active cell's value in that item's format (`bench/windows/samples.ts` reads them). A case is one cell, in a column of
 * its own (`width`, the default without one), with its value and number format, in a book of the case's date system;
 * a set lists its cases (`bench/coverage/samples/<set>.ts`).
 */
import type { BookInput, CellValue, DateSystem, NumberFormat } from '../../../src/contract.js'
import { colLetters } from '../../../src/a1.js'

export interface SampleCase {
  id: string
  value: CellValue
  numberFormat?: NumberFormat
  /** The column's width (`ColumnInput.width`); Excel's default without it. */
  width?: number
  dateSystem?: DateSystem
}

export interface SampleSet {
  buildSampleCases(): SampleCase[]
}

/** The cases of one date system, each in row 2 of its own column of the sheet S from B on (A stays empty), by the order of `cases`. */
export const sampleAddress = (i: number): string => `${colLetters(i + 2)}2`

/** The book of cases of one date system: each case's cell at `sampleAddress` of its index, its column's width set. */
export function sampleBook(cases: readonly SampleCase[]): BookInput {
  const cells: BookInput['sheets'][number]['cells'] = {}
  const columns: NonNullable<BookInput['sheets'][number]['columns']> = {}
  cases.forEach((c, i) => {
    if (c.value !== null || c.numberFormat !== undefined) cells[sampleAddress(i)] = { ...(c.value !== null ? { value: c.value } : {}), ...(c.numberFormat !== undefined ? { numberFormat: c.numberFormat } : {}) }
    if (c.width !== undefined) columns[colLetters(i + 2)] = { width: c.width }
  })
  return { ...(cases[0]?.dateSystem ? { dateSystem: cases[0].dateSystem } : {}), sheets: [{ name: 'S', cells, columns }] }
}

/** The cases by date system, in the order of their first case: one book each. */
export function sampleBooks(cases: readonly SampleCase[]): SampleCase[][] {
  const by = new Map<DateSystem, SampleCase[]>()
  for (const c of cases) {
    const k = c.dateSystem ?? 1900
    if (!by.has(k)) by.set(k, [])
    by.get(k)!.push(c)
  }
  return [...by.values()]
}
