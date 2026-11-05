/**
 * The last round of `file-refusals`: what the writer now makes of what the other rounds showed Excel refuses, each case the file
 * `writeXlsx` writes itself (`written`), and what its choices rest on. Whether a code only in conditional styles counts on top of 206
 * in `numFmts`, which errors a PivotTable's cache takes, and whether two PivotTables may share a name, cut or not.
 *
 * `refused` and `predicted` were committed before Excel ran this set. The model: the writer's files all open, with what they cannot
 * hold left out. Errors beyond the eight [MS-OI29500] lists cost the file; two PivotTables of a name cost it on one sheet, without
 * case, and not across sheets.
 */
import type { BookInput, CellStyle, PivotTable, SheetInput } from '../../../src/contract.js'
import { CACHE as PIVOT_CACHE, PIVOT, pivotBook } from '../../../test/pivot-books.js'
import type { LinksCase, LinksCases } from './cases.js'
import { CACHE, cell, DXF, inSheet, list, one, pivoted, REFUSED, rep, STYLES, subs } from './file-refusals.js'

const fmts = (n: number) => `<numFmts count="${n}">${list(n, (i) => `<numFmt numFmtId="${164 + i}" formatCode="0.00&quot;i${i + 1}&quot;"/>`)}</numFmts>`
const RULE = inSheet('<conditionalFormatting sqref="D1"><cfRule type="expression" dxfId="0" priority="1"><formula>B1&gt;0</formula></cfRule></conditionalFormatting>')
const onDxf = (id: number, code: string) => subs(RULE, [STYLES, '<fonts ', `${fmts(206)}<fonts `], [DXF[0], DXF[1], `<dxfs count="1"><dxf><numFmt numFmtId="${id}" formatCode="0.00&quot;${code}&quot;"/></dxf></dxfs></styleSheet>`])

const asked = (): LinksCase[] => [
  one('numfmt-206-dxf-new', '206 number formats and one more code in a conditional style alone', {}, { edit: onDxf(400, 'new'), ...REFUSED }),
  one('numfmt-206-dxf-same', '206 number formats and a conditional style of one of them', {}, { edit: onDxf(164, 'i1'), holds: 'cfRule' }),
  cell('fn-mode-mult', 'COUNTIF over MODE.MULT itself', 'COUNTIF(_xlfn.MODE.MULT(B1:B3),1)', REFUSED),
  cell('fn-munit', 'COUNTIF over MUNIT itself', 'COUNTIF(_xlfn.MUNIT(3),1)', REFUSED),
]

const ERRORS: [string, string, boolean][] = [
  ['calc', '#CALC!', true],
  ['busy', '#BUSY!', true],
  ['cycle', '#CYCLE!', true],
  ['connect', '#CONNECT!', true],
  ['blocked', '#BLOCKED!', true],
  ['error', '#ERROR!', true],
  ['getting', '#GETTING_DATA', false],
]

const errors = (): LinksCase[] => ERRORS.map(([k, text, refused]) => pivoted(`pivot-error-${k}`, `${text} among a cache's items`, { slicer: false, edit: subs([CACHE, '<s v="East"/>', `<e v="${text}"/>`]), ...(refused ? REFUSED : {}) }))

/** The book with a second PivotTable of `name`, on the sheet of the first or on one of its own. */
function twoTables(second: string, own: boolean, first = 'PivotTable1') {
  return (b: BookInput) => {
    b.slicerCaches = []
    const sheet = b.sheets[1]
    sheet.slicers = []
    sheet.pivotTables = [{ ...PIVOT, name: first }]
    const other: PivotTable = { ...PIVOT, name: second, location: { ...PIVOT.location, range: own ? 'A3:B8' : 'D3:E8' } }
    if (own) b.sheets.push({ name: 'Pivot2', cells: Object.fromEntries(Object.entries(sheet.cells)), pivotTables: [other] })
    else {
      sheet.pivotTables.push(other)
      for (const [a, c] of Object.entries(sheet.cells)) sheet.cells[a.replace('A', 'D').replace('B', 'E')] = c
    }
  }
}
const PART2 = 'xl/pivotTables/pivotTable2.xml'

const names = (): LinksCase[] => [
  pivoted('pivot-twin-sheet', 'two PivotTables of one name on a sheet', { change: twoTables('Second', false), edit: subs([PART2, 'name="Second"', 'name="PivotTable1"']), ...REFUSED }),
  pivoted('pivot-twin-case', 'two PivotTables on a sheet whose names differ in case', { change: twoTables('Second', false), edit: subs([PART2, 'name="Second"', 'name="pivottable1"']), ...REFUSED }),
  pivoted('pivot-twin-sheets', 'two PivotTables of one name on two sheets', { change: twoTables('Second', true), edit: subs([PART2, 'name="Second"', 'name="PivotTable1"']) }),
  pivoted('pivot-two-control', 'two PivotTables of two names on a sheet', { change: twoTables('Second', false) }),
]

// --- what the writer writes

const written = (id: string, about: string, make: () => Partial<BookInput>, extra: Partial<LinksCase> = {}): LinksCase => ({ id, about, sheets: [{ name: 'S', cells: { A1: { value: 5 } } }], linked: {}, links: [], read: [{ sheet: 'S', range: 'A1' }], predicted: { 'S!A1': 5 }, written: true, make, ...extra })
const S = (more: Partial<SheetInput>): SheetInput => ({ name: 'S', cells: { A1: { value: 5 } }, ...more })
const one300 = Object.fromEntries(Array.from({ length: 300 }, (_, i) => [`A${i + 1}`, { value: 5, numberFormat: `0.00"${i}"` }]))
const styles70 = (): CellStyle[] => Array.from({ length: 70000 }, (_, i) => ({ fill: { color: `#${i.toString(16).padStart(6, '0')}` } }))
const TABLE_CELLS = { E1: { value: 'Q' }, E2: { value: 1 } }
const table = (columns: object): Partial<SheetInput> => ({ tables: [{ name: 'Tbl', range: 'E1:E3', totals: true, columns: [{ name: 'Q', ...columns }] }] })
const inCell = (id: string, about: string, formula: string) => written(id, about, () => ({ sheets: [S({ cells: { A1: { value: 5 }, B1: { formula } } })] }))
const onSheet = (id: string, about: string, more: Partial<SheetInput>) => written(id, about, () => ({ sheets: [S(more)] }))

const writer = (): LinksCase[] => [
  inCell('w-countif-array', 'COUNTIF over an array constant', 'COUNTIF({1,2},1)'),
  inCell('w-countif-unique', 'COUNTIF over UNIQUE', 'COUNTIF(UNIQUE(A1:A3),1)'),
  inCell('w-abs-2', 'ABS of two', 'ABS(1,2)'),
  inCell('w-literal-300', 'a literal of 300', `"${rep('a', 300)}"&"b"`),
  inCell('w-long-file-text', 'a formula that grows past 8,192 with its prefixes', list(350, () => 'XLOOKUP(1,A2:A3,B2:B3)', '+')),
  onSheet('w-cf-array', 'a rule over an array constant', { conditionalFormats: [{ type: 'expression', ranges: ['B1'], formula: 'OR(B1={1,2,3})', style: { font: { bold: true } } }] }),
  onSheet('w-dv-union', 'validation over a union', { validations: [{ range: 'B1', type: 'custom', formula1: 'SUM((A1,A2))>0' }] }),
  onSheet('w-dv-list-300', 'a list of 300', { validations: [{ range: 'B1', type: 'list', formula1: `"${rep('abc,', 300)}"` }] }),
  written('w-names', 'names without their sheets', () => ({ names: [{ name: 'rel', ref: 'A1' }, { name: 'abs', ref: '$B$2' }, { name: 'ok', ref: 'S!$A$1' }] })),
  onSheet('w-formats-300', '300 number formats', { cells: one300 }),
  onSheet('w-formats-dxf', '206 number formats and another in a conditional style', { cells: Object.fromEntries(Object.entries(one300).slice(0, 206)), conditionalFormats: [{ type: 'expression', ranges: ['B1'], formula: 'A1>0', style: { numberFormat: '0.00"new"' } }] }),
  written('w-cell-formats-70000', '70,000 cell formats', () => ({ styles: styles70(), sheets: [S({ cells: Object.fromEntries(Array.from({ length: 70000 }, (_, i) => [`A${i + 1}`, { value: 5, style: i }])) })] })),
  written('w-font-long', 'a cell’s font of 40', () => ({ styles: [{ font: { name: rep('F', 40) } }], sheets: [S({ cells: { A1: { value: 5, style: 0 } } })] })),
  written('w-font-empty', 'a cell’s font with an empty name', () => ({ styles: [{ font: { name: '' } }], sheets: [S({ cells: { A1: { value: 5, style: 0 } } })] })),
  written('w-book-font-long', 'the workbook’s font of 40', () => ({ font: { name: rep('F', 40), size: 11 } })),
  onSheet('w-link-location', 'a location of 3,000', { cells: { A1: { value: 5, link: { location: `S!A1${rep('x', 3000)}` } } } }),
  onSheet('w-link-url', 'an address of 3,000', { cells: { A1: { value: 5, link: { url: `https://e.com/?${rep('q', 3000)}` } } } }),
  onSheet('w-dv-66000', '66,000 validation rules', { validations: Array.from({ length: 66000 }, (_, i) => ({ range: `D${i + 1}`, type: 'custom' as const, formula1: 'TRUE' })) }),
  onSheet('w-totals-number', 'a number in the totals row, no function', { cells: { A1: { value: 5 }, ...TABLE_CELLS, E3: { value: 42 } }, ...table({}) }),
  onSheet('w-totals-text', 'a text in the totals row, no label', { cells: { A1: { value: 5 }, ...TABLE_CELLS, E3: { value: 'T' } }, ...table({}) }),
  onSheet('w-totals-label', 'a label other than the cell’s text', { cells: { A1: { value: 5 }, ...TABLE_CELLS, E3: { value: 'T' } }, ...table({ totals: { label: 'Total' } }) }),
  onSheet('w-totals-sum-other', 'sum over another formula', { cells: { A1: { value: 5 }, ...TABLE_CELLS, E3: { formula: 'E2*2' } }, ...table({ totals: { function: 'sum' } }) }),
  onSheet('w-totals-sum-own', 'sum over its own formula', { cells: { A1: { value: 5 }, ...TABLE_CELLS, E3: { formula: 'SUBTOTAL(109,Tbl[Q])' } }, ...table({ totals: { function: 'sum' } }) }),
  onSheet('w-totals-sum-empty', 'sum over an empty cell', { cells: { A1: { value: 5 }, ...TABLE_CELLS }, ...table({ totals: { function: 'sum' } }) }),
]

const pivotWritten = (id: string, about: string, change: (b: BookInput) => void) =>
  written(id, about, () => {
    const b = pivotBook()
    change(b)
    return b
  }, { sheets: [{ name: 'Data', cells: {} }], read: [{ sheet: 'Data', range: 'B2' }], predicted: { 'Data!B2': 10 } })

const noSlicer = (b: BookInput) => {
  b.slicerCaches = []
  b.sheets[1].slicers = []
}

const pivots = (): LinksCase[] => [
  pivotWritten('w-slicer', 'a slicer of no items', () => {}),
  pivotWritten('w-slicer-space', 'a slicer cache whose name is no name', (b) => {
    b.slicerCaches![0].name = 'Slicer Region'
    b.sheets[1].slicers![0].cache = 'Slicer Region'
  }),
  pivotWritten('w-pivot-spill', '#SPILL! among a cache’s items', (b) => {
    noSlicer(b)
    b.pivotCaches![0].fields[0].items = ['North', 'South', 'East', { error: 'SPILL' }]
  }),
  pivotWritten('w-pivot-field-twice', 'two fields of a name', (b) => {
    noSlicer(b)
    b.pivotCaches![0].fields[2].name = 'region'
  }),
  pivotWritten('w-pivot-field-empty', 'a field of no name', (b) => {
    noSlicer(b)
    b.pivotCaches![0].fields[2].name = ''
  }),
  pivotWritten('w-pivot-name-300', 'a PivotTable’s name of 300', (b) => {
    noSlicer(b)
    b.sheets[1].pivotTables![0].name = rep('p', 300)
  }),
  pivotWritten('w-pivot-names-cut-alike', 'two names of 300 alike in their first 255', (b) => {
    twoTables(`${rep('p', 299)}b`, false, `${rep('p', 299)}a`)(b)
  }),
  pivotWritten('w-pivot-names-case', 'a name of 300 cut into another’s in another case', (b) => {
    twoTables(rep('P', 255), false, rep('p', 300))(b)
  }),
  pivotWritten('w-pivot-caption-300', 'a grand total’s caption of 300', (b) => {
    noSlicer(b)
    b.sheets[1].pivotTables![0].captions = { grandTotal: rep('g', 300) }
  }),
  pivotWritten('w-pivot-source-long', 'a source sheet of 40 the book does not have', (b) => {
    noSlicer(b)
    b.pivotCaches![0].source = { ...PIVOT_CACHE.source, sheet: rep('s', 40) } as typeof PIVOT_CACHE.source
  }),
]

export function buildLinksCases(): LinksCases {

  return { cases: [...asked(), ...errors(), ...names(), ...writer(), ...pivots()] }
}
