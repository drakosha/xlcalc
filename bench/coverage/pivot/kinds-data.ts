/**
 * The rows of `kinds`: a column per kind of field Excel's PivotTable Fields list may place differently, twelve rows:
 * numbers alone, with an empty cell, with text, a boolean, an error or a date among them, empty cells alone, text,
 * booleans, errors, numbers kept as text, times of day; dates over three years, within a year, within a month, within a
 * day (with times), with an empty cell, with text among them.
 */
import type { CellInput } from '../../../src/contract.js'
import type { SourceColumn } from './data.js'

const serial = (y: number, m: number, d: number, h = 0) => Date.UTC(y, m - 1, d, h) / 86_400_000 + 25_569
const date = (v: number): CellInput => ({ value: v, numberFormat: 'yyyy-mm-dd' })

export function kindColumns(): SourceColumn[] {
  const rows = Array.from({ length: 12 }, (_, k) => k)
  const col = (name: string, f: (k: number) => CellInput | undefined): SourceColumn => ({ name, values: rows.map(f) })
  return [
    col('Num', (k) => ({ value: k * 3 + 1 })),
    col('NumBlank', (k) => (k % 5 === 2 ? undefined : { value: k + 1 })),
    col('NumText', (k) => ({ value: k === 4 ? 'x' : k + 1 })),
    col('NumBool', (k) => ({ value: k === 4 ? true : k + 1 })),
    col('NumErr', (k) => (k === 4 ? { formula: 'NA()' } : { value: k + 1 })),
    col('NumDate', (k) => (k === 4 ? date(serial(2025, 3, 1)) : { value: k + 1 })),
    col('Blank', () => undefined),
    col('Text', (k) => ({ value: ['a', 'b', 'c'][k % 3] })),
    col('Bool', (k) => ({ value: k % 2 === 0 })),
    col('Err', (k) => ({ formula: k % 2 ? 'NA()' : '1/0' })),
    col('NumAsText', (k) => ({ value: String(k + 1) })),
    col('Times', (k) => ({ value: k / 24, numberFormat: 'hh:mm' })),
    col('Years', (k) => date(serial(2023 + (k % 3), 1 + k, 1 + k))),
    col('OneYear', (k) => date(serial(2025, 1 + k, 5))),
    col('OneMonth', (k) => date(serial(2025, 3, 1 + k * 2))),
    col('OneDay', (k) => ({ value: serial(2025, 3, 4, k * 2), numberFormat: 'yyyy-mm-dd hh:mm' })),
    col('DateBlank', (k) => (k % 4 === 1 ? undefined : date(serial(2024 + (k % 2), 1 + k, 3)))),
    col('DateText', (k) => (k === 3 ? { value: 'none' } : date(serial(2024 + (k % 2), 1 + k, 3)))),
  ]
}
