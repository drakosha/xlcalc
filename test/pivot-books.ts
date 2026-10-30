/** Workbooks with PivotTables for the tests: amounts by region and month on Data, a PivotTable of their sums on Pivot, a slicer and a timeline over it. */
import type { BookInput, PivotCache, PivotTable } from '../src/index.js'

const REGIONS = ['North', 'South', 'East', 'West']

/** Data!A1:C9: Region, Amount, Month; North 10, South 20 … in two months. */
export function dataSheet(): BookInput['sheets'][number] {
  const cells: BookInput['sheets'][number]['cells'] = { A1: { value: 'Region' }, B1: { value: 'Amount' }, C1: { value: 'Month' } }
  for (let i = 0; i < 8; i++) {
    cells[`A${i + 2}`] = { value: REGIONS[i % 4] }
    cells[`B${i + 2}`] = { value: (i + 1) * 10 }
    cells[`C${i + 2}`] = { value: i < 4 ? 'Jan' : 'Feb' }
  }
  return { name: 'Data', cells }
}

export const CACHE: PivotCache = {
  id: 1,
  source: { type: 'range', sheet: 'Data', range: 'A1:C9' },
  fields: [
    { name: 'Region', items: REGIONS, facts: { types: ['string'] } },
    { name: 'Amount', facts: { types: ['number'], integer: true, min: 10, max: 80 } },
    { name: 'Month', items: ['Jan', 'Feb'], facts: { types: ['string'] } },
  ],
  records: Array.from({ length: 8 }, (_, i) => [i % 4, (i + 1) * 10, i < 4 ? 0 : 1]),
}

/** Sum of Amount by Region at Pivot!A3:B8, as Excel lays it out: the regions, then the grand total. */
export const PIVOT: PivotTable = {
  name: 'PivotTable1',
  cache: 1,
  location: { range: 'A3:B8', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 },
  fields: [{ items: [0, 1, 2, 3].map((item) => ({ item })) }, {}, {}],
  rows: [0],
  values: [{ field: 1, name: 'Sum of Amount' }],
  rowLines: [...[0, 1, 2, 3].map((k) => ({ items: [k] })), { type: 'grand', items: [0] }],
  columnLines: [{ items: [] }],
}

/** The PivotTable's cells as Excel shows them. */
export function pivotSheet(): BookInput['sheets'][number] {
  const cells: BookInput['sheets'][number]['cells'] = { A3: { value: 'Row Labels' }, B3: { value: 'Sum of Amount' } }
  REGIONS.forEach((r, k) => {
    cells[`A${k + 4}`] = { value: r }
    cells[`B${k + 4}`] = { value: (k + 1) * 10 + (k + 5) * 10 }
  })
  cells.A8 = { value: 'Grand Total' }
  cells.B8 = { value: 360 }
  return { name: 'Pivot', cells, pivotTables: [PIVOT], slicers: [{ name: 'Region', cache: 'Slicer_Region' }] }
}

export const pivotBook = (): BookInput => ({
  sheets: [dataSheet(), pivotSheet()],
  pivotCaches: [CACHE],
  slicerCaches: [{ name: 'Slicer_Region', field: 'Region', pivots: [{ sheet: 'Pivot', name: 'PivotTable1' }], cache: 1 }],
})
