import { readFileSync } from 'node:fs'
import type { BookInput } from '../../../src/contract.js'
import { registerTableStyles } from '../../../src/table-styles.js'
import { writeXlsx } from '../../../src/xlsx.js'

export const REGION_CASES = [
  { id: 'custom-bands', sheet: 'S', range: 'B2:D8' },
  { id: 'theme-table', sheet: 'S', range: 'B10:D12' },
  { id: 'pivot-subheads', sheet: 'P', range: 'A3:B8' },
  { id: 'pivot-report-filter', sheet: 'P', range: 'A1:B4' },
] as const

const BOOK: BookInput = {
  tableStyles: {
    theme: ['#FFFFFF', '#000000', '#EEEEEE', '#222222', '#126599', '#AABBCC', '#991122', '#117733', '#223344', '#AACC11', '#122113', '#765432'],
    custom: {
      Forecast: {
        table: true, pivot: false, elements: [
          { type: 'wholeTable', style: { font: { color: '#122113' } } },
          { type: 'firstRowStripe', size: 2, style: { fill: { color: '#AAFF99' } } },
          { type: 'secondRowStripe', style: { fill: { color: '#FFDD44' } } },
          { type: 'firstColumnStripe', size: 2, style: { font: { italic: true } } },
          { type: 'secondColumnStripe', style: { font: { color: '#991122' } } },
          { type: 'firstColumn', style: { font: { bold: true } } },
          { type: 'headerRow', style: { fill: { color: '#126599' } } },
          { type: 'lastHeaderCell', style: { font: { color: '#FFFFFF', bold: true } } },
        ],
      },
    },
  },
  sheets: [{
    name: 'S',
    cells: {
      B2: { value: 'Name' }, C2: { value: 'Amount' }, D2: { value: 'Memo' },
      B3: { value: 'One' }, C3: { value: 10 },
      B10: { value: 'Name' }, C10: { value: 'Amount' }, D10: { value: 'Memo' },
    },
    tables: [
      { name: 'Forecast', range: 'B2:D8', columns: [{ name: 'Name' }, { name: 'Amount' }, { name: 'Memo' }], style: { name: 'Forecast', rowStripes: true, columnStripes: true, firstColumn: true } },
      { name: 'Baseline', range: 'B10:D12', columns: [{ name: 'Name' }, { name: 'Amount' }, { name: 'Memo' }] },
    ],
  }],
}

export async function regionFile(id: (typeof REGION_CASES)[number]['id']): Promise<Uint8Array> {
  if (id === 'custom-bands' || id === 'theme-table') {
    registerTableStyles()
    return (await writeXlsx({ book: BOOK })).data
  }
  const file = id === 'pivot-subheads' ? 'two-rows' : 'page'
  return new Uint8Array(readFileSync(new URL(`../../../test/excel/win/pivot/ru-RU/basic/${file}.xlsx`, import.meta.url)))
}
