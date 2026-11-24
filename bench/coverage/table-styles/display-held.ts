import { readFileSync } from 'node:fs'
import type { BookInput } from '../../../src/contract.js'
import { writeXlsx } from '../../../src/xlsx.js'

export const DISPLAY_CASES = [
  { id: 'banded-table', sheet: 'S', range: 'A1:C5' },
  { id: 'saved-pivot', sheet: 'P', range: 'A3:B8' },
] as const

const TABLE: BookInput = {
  styles: [
    { fill: { color: '#FFFF00' }, font: { bold: false, color: '#112233' } },
    { fill: { color: '#FF00FF' } },
  ],
  sheets: [{
    name: 'S',
    cells: {
      A1: { value: 'Name' }, B1: { value: 'Amount' }, C1: { value: 'Memo' },
      A2: { value: 'One' }, B2: { value: 10, style: 0 },
      A3: { value: 'Two', style: 1 }, B3: { value: 20 },
      A4: { value: 'Three' }, B4: { value: 30 },
      A5: { value: 'Total' }, B5: { formula: 'SUM(B2:B4)' },
    },
    tables: [{
      name: 'Amounts', range: 'A1:C5', totals: true,
      columns: [{ name: 'Name', totals: { label: 'Total' } }, { name: 'Amount', totals: { function: 'sum' } }, { name: 'Memo' }],
      style: { name: 'TableStyleMedium2', rowStripes: true, firstColumn: true },
    }],
    conditionalFormats: [{ type: 'expression', ranges: ['B2'], formula: 'TRUE', style: { fill: { color: '#00FF00' }, font: { bold: true } } }],
  }],
}

export async function displayFile(id: (typeof DISPLAY_CASES)[number]['id']): Promise<Uint8Array> {
  if (id === 'banded-table') return (await writeXlsx({ book: TABLE })).data
  return new Uint8Array(readFileSync(new URL('../../../test/excel/win/pivot/ru-RU/basic/columns.xlsx', import.meta.url)))
}
