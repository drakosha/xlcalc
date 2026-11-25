/** Excel saved the paired workbooks unless their names start with lo-. The .xls keeps its original caches. */
import { readFile } from 'node:fs/promises'
import { describe, expect, it } from 'vitest'
import { libUrl } from './fixtures/target.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { readXlsxWithPrint, writeXlsxWithPrint } from '../src/xlsx-print.js'

const { readXlsWithPrint } = await import(libUrl('xls-print.js')) as typeof import('../src/xls-print.js')
const { readXls } = await import(libUrl('xls.js')) as typeof import('../src/xls.js')
const fixture = (name: string, extension = 'xls') => readFile(new URL(`./excel/xls-read/${name}.${extension}`, import.meta.url))
const pair = async (name: string) => Promise.all([fixture(name).then(readXls), fixture(name, 'xlsx').then(readXlsx)])

function discardRecreatedTotalsLabel(read: Awaited<ReturnType<typeof readXlsx>>): void {
  // Excel creates a localized default label when it opens an XLS with a hidden totals row.
  const first = read.book.sheets[0].tables![0].columns[0]
  expect(first.totals).toEqual({ label: 'Итог' })
  delete first.totals
}

describe('readXlsWithPrint', () => {
  it('keeps Excel’s page setup, margins, header and footer variants and manual breaks on conversion to XLSX', async () => {
    const binary = await fixture('print').then(readXlsWithPrint)
    const xlsx = await fixture('print', 'xlsx').then(readXlsxWithPrint)
    expect(binary.printIssues).toStrictEqual(['Landscape', 'Scale', 'Fit'].map((sheet) => ({ sheet, part: 'pageSetup', reason: 'printerSettings' })))
    const find = (name: string) => xlsx.print.find((sheet) => sheet.sheet === name)!
    expect(binary.print).toEqual(expect.arrayContaining([
      expect.objectContaining({ sheet: 'Default', margins: find('Default').margins }),
      expect.objectContaining({ sheet: 'HeaderFooter', headerFooter: find('HeaderFooter').headerFooter }),
      expect.objectContaining({ sheet: 'Landscape', setup: expect.objectContaining(find('Landscape').setup!) }),
      expect.objectContaining({ sheet: 'Scale', setup: expect.objectContaining(find('Scale').setup!) }),
      expect.objectContaining({ sheet: 'Fit', setup: expect.objectContaining(find('Fit').setup!) }),
      expect.objectContaining({ sheet: 'Breaks', rowBreaks: find('Breaks').rowBreaks }),
    ]))
    const base = await fixture('print').then(readXls)
    expect({ book: binary.book, cached: binary.cached, issues: binary.issues }).toStrictEqual(base)
    expect(binary.sheetDefaults).toContainEqual({ sheet: 'Default', defaultRowHeight: 15 })
    const written = await writeXlsxWithPrint({ book: binary.book, result: binary.cached }, binary.print, { values: 'preserve' })
    expect((await readXlsxWithPrint(written.data)).print).toStrictEqual(binary.print)
  })

  it('reads LibreOffice’s parenthesized print areas and titles, margins, print flags and manual column break', async () => {
    const binary = await fixture('lo-print-titles').then(readXlsWithPrint)
    const xml = await fixture('lo-print-titles', 'xlsx').then(readXlsxWithPrint)
    const [expected] = xml.print
    expect(expected.areas).toEqual(['B2:D4', 'F7:H9'])
    expect(binary.printIssues).toStrictEqual([])
    expect(binary.print).toEqual([expect.objectContaining({
      sheet: "O'Brien Print", areas: expected.areas, repeatRows: expected.repeatRows, repeatColumns: expected.repeatColumns,
      options: expected.options, columnBreaks: [{ before: 3, manual: true }],
      headerFooter: { oddHeader: '&CReport &P' },
      setup: expect.objectContaining({ firstPageNumber: 7, useFirstPageNumber: true, pageOrder: 'overThenDown' }),
    })])
    expect(binary.print[0].margins?.left).toBeCloseTo(0.35)
    expect(binary.print[0].margins?.bottom).toBeCloseTo(0.85)
    const written = await writeXlsxWithPrint({ book: binary.book, result: binary.cached }, binary.print, { values: 'preserve' })
    expect((await readXlsxWithPrint(written.data)).print).toStrictEqual(binary.print)
  })

  it('reads the same stored print area Excel writes when saving the PivotTable fixture as XLSX', async () => {
    const binary = await fixture('held-print-area-source').then(readXlsWithPrint)
    const xml = await fixture('held-print-area-source', 'xlsx').then(readXlsxWithPrint)
    const areas = (result: typeof xml) => result.print.filter((sheet) => sheet.areas).map(({ sheet, areas }) => ({ sheet, areas }))
    expect(areas(xml)).toHaveLength(1)
    expect(areas(binary)).toStrictEqual(areas(xml))
    const written = await writeXlsxWithPrint({ book: binary.book, result: binary.cached }, binary.print, { values: 'preserve' })
    const saved = await readXlsxWithPrint(written.data)
    expect(areas(saved)).toStrictEqual(areas(binary))
    expect(saved.book.pivotCaches?.map(({ source }) => source)).toStrictEqual(binary.book.pivotCaches?.map(({ source }) => source))
  })
})

describe('readXls', () => {
  it.each(['probe', 'notes', 'conditional', 'validation-multi-range', 'filter-and', 'filter-non-blanks', 'filter-long-value', 'table-totals-formula', 'sheet-view', 'what-if', 'styles', 'print'])(
    'reads both of Excel’s %s files identically',
    async (name) => {
      const [binary, xml] = await pair(name)
      expect(binary).toStrictEqual(xml)
    },
  )

  it.each(['looks-text', 'looks-period', 'looks-top', 'looks-scale'])('reads Excel’s %s formatting rules', async (name) => {
    const [binary, xml] = await pair('conditional-' + name)
    expect(binary).toStrictEqual(xml)
  })

  it('reads native conditional templates and the BIFF data bar settings', async () => {
    const [binary, xml] = await pair('conditional-made-made')
    const rules = (read: typeof binary) => read.book.sheets[0].conditionalFormats ?? []
    expect(rules(binary).filter((rule) => rule.type !== 'dataBar'))
      .toStrictEqual(rules(xml).filter((rule) => rule.type !== 'dataBar'))
    // On reopening, Excel upgrades the original bar to an x14 bar and changes its display flags.
    expect(rules(binary).find((rule) => rule.type === 'dataBar'))
      .toMatchObject({ excel2007: true, minLength: 0, maxLength: 100, showValue: false })
    expect(rules(xml).find((rule) => rule.type === 'dataBar')).toMatchObject({ axis: 'none' })
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it.each([
    ['theme-absent', '#C0E6F5'],
    ['theme-marker', '#DCE6F1'],
  ])('reads the implicit colors and differential font name in %s', async (name, accent) => {
    const [binary, xml] = await pair(name)
    expect(xml.book.sheets[0].conditionalFormats?.[3]).toMatchObject({ style: { fill: { color: accent } } })
    expect(xml.issues).toContainEqual(expect.objectContaining({ code: 'conditionalFormat', message: "a conditional format's font name is not in the model" }))
    expect(binary).toStrictEqual(xml)
  })

  it('keeps an explicit Normal font color on cells, including direct font overrides', async () => {
    const { book } = await fixture('layout-normal-style').then(readXls)
    expect(book.font).toEqual({ name: 'Arial', size: 10 })
    const cells = book.sheets[0].cells
    const font = (at: string) => book.styles?.[cells[at].style ?? -1]?.font
    expect(font('A1')).toEqual({ bold: true, color: '#154488' })
    expect(font('A2')).toEqual({ bold: true, color: '#AA6622' })
    expect(font('A3')).toEqual({ color: '#154488' })
    expect(font('A4')).toEqual({ bold: true })
    expect(font('A5')).toEqual({ name: 'Times New Roman', bold: true, color: '#154488' })
  })

  it.each(['sentinel-baseline', 'sentinel-custom', 'sentinel-hidden', 'sentinel-style', 'sentinel-with-standard'])(
    'reads Excel’s default column formatting in %s',
    async (name) => {
      const [binary, xml] = await pair(name)
      expect(binary).toStrictEqual(xml)
    },
  )

  it('reads a default column record outside the binary grid', async () => {
    const [binary, xml] = await pair('sentinel-only')
    const actual = binary.book.sheets[0]
    const expected = xml.book.sheets[0]
    expect(actual.defaultColumnWidth).toBe(expected.defaultColumnWidth)
    // The file gives only a coarse base width here. Excel for Windows and the model use different font metrics.
    expect(expected.columns?.['G:IV']?.width).toBe(8.88671875)
    expect(actual.columns?.['G:IV']?.width).toBe(8.83203125)
    expect(actual.columns?.F).toEqual(expected.columns?.F)
    expect(actual.cells).toEqual(expected.cells)
  })

  it('keeps an explicit width ending at IV off the sheet default', async () => {
    const [binary, xml] = await pair('regular-end-iv')
    const sheet = binary.book.sheets[0]
    expect(sheet.defaultColumnWidth).toBeUndefined()
    expect(sheet.columns?.['G:IV']).toEqual(xml.book.sheets[0].columns?.['G:IV'])
  })

  it.each([
    'sheet-above-average', 'sheet-today', 'sheet-fill', 'sheet-icon-third', 'sheet-numbers', 'sheet-sort-two',
    'sheet-top', 'sheet-values-blank', 'table-sort-two', 'table-values', 'sheet-hidden-button-filter',
  ])('reads the filter and sort in Excel’s %s workbook', async (name) => {
    const [binary, xml] = await pair('filter-' + name)
    expect(binary).toStrictEqual(xml)
  })

  it('reads AutoFilter date groups split across ContinueFrt12 records', async () => {
    const [binary, xml] = await pair('filter-date-groups')
    const filter = binary.book.sheets[0].autoFilter
    expect(filter).toStrictEqual(xml.book.sheets[0].autoFilter)
    expect(filter?.criteria?.columns?.[0].filter).toEqual({
      type: 'values', dates: [{ level: 'month', year: 2025, month: 3 }, { level: 'year', year: 2024 }],
    })
  })

  it.each(['in-place', 'copy', 'combined'])('reads the cells and filter mode Excel saved after an advanced filter (%s)', async (mode) => {
    const [binary, xml] = await pair('advanced-' + mode)
    expect(binary).toStrictEqual(xml)
    if (mode === 'in-place') {
      expect(binary.book.sheets[0]).toMatchObject({ filtered: true, rows: { 2: { hidden: true } } })
      expect(binary.issues.some(({ code }) => code === 'autoFilter')).toBe(true)
    }
  })

  it.each(['tables-refs-no-headers', 'tables-refs-totals', 'table-totals-offset', 'tables-refs-inside', 'tables-refs-names', 'tables-refs-this-row'])(
    'reads the table layout and columns in Excel’s %s workbook',
    async (name) => {
      const [binary, xml] = await pair(name)
      const source = binary.book.sheets.map((sheet) => sheet.tables)
      if (['tables-refs-inside', 'tables-refs-names', 'tables-refs-this-row'].includes(name)) discardRecreatedTotalsLabel(xml)
      expect(source).toStrictEqual(xml.book.sheets.map((sheet) => sheet.tables))
      expect(binary.issues).toStrictEqual(xml.issues)
    },
  )

  it('reads a wide table spanning two ContinueFrt11 records', async () => {
    const [binary, xml] = await pair('table-wide')
    expect(binary.book.sheets[0].tables?.[0].columns).toHaveLength(100)
    discardRecreatedTotalsLabel(xml)
    expect(binary).toStrictEqual(xml)
  })

  it.each(['table-style-empty', 'table-style-stripe-only'])('keeps a table with its style cleared in %s', async (name) => {
    const [binary, xml] = await pair(name)
    expect(binary).toStrictEqual(xml)
  })

  it('reconstructs calculated table columns from Excel’s saved formulas', async () => {
    const [binary, xml] = await pair('table-calculated')
    expect(binary.book.sheets.map(({ tables }) => tables?.[0].columns[1].formula))
      .toStrictEqual(xml.book.sheets.map(({ tables }) => tables?.[0].columns[1].formula))
    expect(binary.cached).toStrictEqual(xml.cached)
  })

  it.each(['formula-small-numbers', 'formula-whole-sheet', 'functions-066', 'range-start-name'])('reads the saved formulas of Excel’s %s workbook', async (name) => {
    const [binary, xml] = await pair(name)
    expect(binary.book).toStrictEqual(xml.book)
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it('keeps the binary grid bounds in ranges over another workbook', async () => {
    const [binary, xml] = await pair('external-areas')
    expect(xml.book.sheets[0].cells.A1.formula).toBe('SUM([1]Ledger!A$1:A$65536)')
    expect(xml.book.sheets[0].cells.A5.formula).toBe('SUM(Local!$2:$3)')
    expect(binary).toStrictEqual(xml)
  })

  it.each([6, 7, 8])('keeps an external workbook whose BIFF path has base type %i', async (pathType) => {
    const [bytes, xlsx] = await Promise.all([fixture('external-areas'), fixture('external-areas', 'xlsx')])
    const offset = bytes.indexOf('\x01outside.xlsx')
    expect(offset).toBeGreaterThan(0)
    const target = 'Ledger.xlsx'
    // The base byte replaces one character, keeping the SupBook and sheet offsets intact.
    bytes.write('\x01' + String.fromCharCode(pathType) + target, offset, 'latin1')
    const [binary, expected] = await Promise.all([readXls(bytes), readXlsx(xlsx)])
    expected.book.externalBooks![0].target = target
    expect(binary).toStrictEqual(expected)
  })

  it('keeps external sheets and their references past the SupBook record boundary', async () => {
    const [binary, xml] = await pair('many-link-sheets')
    const sheets = binary.book.externalBooks?.[0].sheets
    expect(sheets).toHaveLength(330)
    expect(sheets?.[0]).toEqual({ name: 'LedgerSheetNameNameName0000', cells: { A1: 99 } })
    expect(sheets?.[329]).toEqual({ name: 'LedgerSheetNameNameName0329', cells: { A1: 428 } })
    const cells = binary.book.sheets[0].cells
    expect(cells.A1).toEqual(xml.book.sheets[0].cells.A1)
    expect(cells.A2.formula).toBe('[1]LedgerSheetNameNameName0329!A1')
    // Excel's XLSX export loses the sheets after Continue and keeps the saved value of their references.
    expect(xml.book.sheets[0].cells.A2.formula).toBe('#REF!')
    expect(binary.cached).toStrictEqual(xml.cached)
    expect(binary.issues).toEqual([])
  })

  it('resolves formulas on both sides of a continued ExternSheet index', async () => {
    const { book, cached, issues } = await readXls(await fixture('continued-external-index'))
    const cells = book.sheets[0].cells
    expect(book.externalBooks?.[0].sheets).toHaveLength(1490)
    expect(cells.A1368.formula).toBe('[1]LedgerM1367!A1')
    expect(cells.A1369.formula).toBe('[1]LedgerM1368!A1')
    expect(cells.A1490.formula).toBe('[1]LedgerM1489!A1')
    expect(cached.sheets[0].cells.A1490).toEqual({ value: 1588 })
    expect(issues).toEqual([])
  })

  it('uses the worksheet version for historical Formula record identifiers', async () => {
    const [binary, xml] = await pair('legacy-formula-records')
    expect(xml.cached.sheets[0].cells).toMatchObject({ B1: { value: 6 }, C1: { value: 'OLD' }, D1: { value: true } })
    expect(binary).toStrictEqual(xml)
  })

  it.each([
    ['column-label', 0, 'C9', 225], ['column-relative', 0, 'C9', 225], ['column-sum', 0, 'C9', 225],
    ['column-value', 0, 'C9', 225], ['row-relative', 0, 'C9', 225], ['row-value', 0, 'C9', 225],
    ['range-label', 1, 'A2', 0], ['radical-relative', 1, 'A2', 0], ['radical-ref-error', 1, 'A2', 0],
    ['deleted-label', 0, 'H1', 15], ['deleted-label-range', 0, 'H1', 15],
    ['stacked-column', 0, 'H1', 15], ['stacked-value', 0, 'H1', 15], ['stacked-sum', 0, 'H2', 21],
    ['stacked-absolute-range', 0, 'H3', 28],
  ] as const)('reads the legacy label formula in %s', async (name, sheet, cell, saved) => {
    const [binary, xml] = await pair(name)
    const expected = xml.book.sheets[sheet].cells[cell].formula
    expect(expected).toBeDefined()
    expect(binary.book.sheets[sheet].cells[cell].formula).toBe(expected)
    expect(binary.cached.sheets[sheet].cells[cell].value).toBe(saved)
    expect(binary.issues.some(({ code }) => code === 'formula')).toBe(false)
  })

  it('keeps saved values and errors when their BIFF formulas cannot be translated', async () => {
    const { book, cached, issues } = await readXls(await fixture('formula-fallback'))
    expect(book.sheets[0].cells).toMatchObject({ A1: { value: 4 }, B1: { value: 6 }, C1: { value: 'OLD' } })
    expect(cached.sheets[0].cells).toEqual({
      B1: { value: 6 }, C1: { value: 'OLD' }, D1: { value: { error: 'GETTING_DATA' } },
    })
    expect(issues).toContainEqual(expect.objectContaining({ code: 'formula', cells: expect.arrayContaining(['B1', 'C1', 'D1']) }))
    expect(issues.some(({ code }) => code === 'errorConstant')).toBe(false)
  })

  it('keeps the cached result of a formula that ends at an incomplete label token', async () => {
    const bytes = await fixture('probe')
    const formula = Buffer.from('0c0024000000c01e020042020400', 'hex')
    const at = bytes.indexOf(formula)
    expect(at).toBeGreaterThan(0)
    bytes.writeUInt16LE(1, at)
    bytes[at + 2] = 0x18
    const read = await readXls(bytes)
    expect(read.book.sheets[0].cells.A2).toEqual({ value: 5 })
    expect(read.cached.sheets[0].cells.A2).toEqual({ value: 5 })
    expect(read.issues).toContainEqual(expect.objectContaining({ code: 'formula', cells: ['A2'] }))
  })

  it.each(['param-names', 'param-names-held'])('separates stored parameters from user names in %s', async (name) => {
    const [binary, xml] = await pair(name)
    expect(binary.book).toStrictEqual(xml.book)
    expect(binary.issues).toStrictEqual(xml.issues)
    expect(binary.cached.sheets[0].cells).toMatchObject({ A1: { value: 21 }, A2: { value: 11 } })
  })

  it('keeps names from the old grid that would be cell addresses in the expanded grid', async () => {
    const [binary, xml] = await pair('grid-names')
    const names = (read: typeof binary) => Object.fromEntries((read.book.names ?? []).map(({ name, ref }) => [name, ref]))
    expect(names(xml)).toMatchObject({ __JB65536: '1', _IV65537: '2', _A1048576: '3', _XFD44114: '4', XFE1599: '5' })
    expect(names(binary)).toEqual(names(xml))
    expect(binary.book.sheets).toStrictEqual(xml.book.sheets)
    expect(binary.cached).toStrictEqual(xml.cached)
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it.each(['array-exact', 'range-sheet-quote'])('matches the %s formulas and preserves the saved cache', async (name) => {
    const [binary, xml] = await pair(name)
    expect(binary.book).toStrictEqual(xml.book)
    expect(binary.issues).toStrictEqual(xml.issues)
    // Excel recalculates some legacy array formulas when it reloads the XLS. FormulaValue in the XLS is #VALUE!.
    const at = name === 'array-exact' ? 'A100' : 'A115'
    expect(binary.cached.sheets.find(({ name }) => name === 'C')?.cells[at]?.value).toEqual({ error: 'VALUE' })
    expect(xml.cached.sheets.find(({ name }) => name === 'C')?.cells[at]?.value).not.toEqual({ error: 'VALUE' })
  })

  it.each(['protect-allow', 'protect-dialog', 'sheet-password'])('reads Excel’s %s sheet protection', async (name) => {
    const [binary, xml] = await pair(name)
    expect(binary).toStrictEqual(xml)
  })

  it('reads stored PivotTable axes, items and values with its separate compound stream', async () => {
    const [binary, xml] = await pair('pivot-show-as')
    const pivot = binary.book.sheets.find(({ name }) => name === 'P')?.pivotTables?.[0]
    const expected = xml.book.sheets.find(({ name }) => name === 'P')?.pivotTables?.[0]
    expect(pivot).toStrictEqual(expected)
    expect(binary.cached).toStrictEqual(xml.cached)
    expect(binary.issues).toStrictEqual(xml.issues)
    const cache = binary.book.pivotCaches?.[0]
    expect(cache?.source).toEqual({ type: 'range', sheet: 'D', range: 'A1:J49' })
    expect(cache?.records).toHaveLength(48)
    expect(cache?.records?.[20][3]).toEqual({ error: 'NA' })
    expect(cache?.records?.[41][3]).toEqual({ error: 'DIV0' })
  })

  it.each(['pivot-options', 'pivot-compact-options'])('reads Excel’s %s view flags and compact indent', async (name) => {
    const [binary, xml] = await pair(name)
    const pivot = (read: typeof binary) => read.book.sheets[1].pivotTables?.[0]
    expect(pivot(xml)).toMatchObject({ indent: 3, fieldHeaders: false, showDrill: false })
    expect(pivot(binary)).toStrictEqual({ ...pivot(xml), cache: pivot(binary)?.cache })
  })

  it.each([
    ['extra-pivot-display', 'Regions & groups', 'Time & sizes'],
    ['held-caption-unicode', '中文&'.repeat(84) + 'end', 'Итог'],
  ])('reads Excel’s custom compact captions in %s', async (name, rowHeader, columnHeader) => {
    const [binary, xml] = await pair(name)
    const view = binary.book.sheets[0].pivotTables?.[0]
    const reference = xml.book.sheets[0].pivotTables?.[0]
    expect(reference?.captions).toEqual({ rowHeader, columnHeader })
    expect(view).toStrictEqual({ ...reference, cache: view?.cache })
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it.each([
    ['held-print-area-source', { type: 'name', name: 'Data!_xlnm.Print_Area' }],
    ['held-scenario-pivot', { type: 'scenario' }],
  ] as const)('reads the PivotCache of %s', async (name, source) => {
    const [binary, xml] = await pair(name)
    const model = binary.book.pivotCaches?.[0]
    const reference = xml.book.pivotCaches?.[0]
    expect(reference?.source).toEqual(source)
    expect(model).toStrictEqual({ ...reference, id: model?.id })
    const view = binary.book.sheets[0].pivotTables?.[0]
    expect(view).toStrictEqual({ ...xml.book.sheets[0].pivotTables?.[0], cache: model?.id })
    expect(binary.cached).toStrictEqual(xml.cached)
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it.each([
    ['external-range', 3, null],
    ['external-name', 3, null],
    ['external-scoped', 3, "'Data Scope'!ScopedLedger"],
    ['local-scoped', 3, "'Local Data'!ScopedLedger"],
    ['consolidation', 6, null],
  ] as const)('reads Excel for Windows’s %s PivotTable source and records', async (name, rows, scopedName) => {
    const [binary, xml] = await pair('pivot-source-' + name)
    const cache = binary.book.pivotCaches?.[0]
    const expected = xml.book.pivotCaches?.[0]
    expect(expected?.records).toHaveLength(rows)
    if (scopedName) expect(expected?.source).toMatchObject({ type: 'name', name: scopedName })
    expect(cache).toStrictEqual({ ...expected, id: cache?.id })
    const table = (read: typeof binary) => read.book.sheets.find((sheet) => sheet.name === 'Report')?.pivotTables?.[0]
    expect(table(binary)).toStrictEqual({ ...table(xml), cache: cache?.id })
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it.each(['external-pivot-range', 'external-pivot-name'])('reads a relative source workbook in %s', async (name) => {
    const [binary, xml] = await pair(name)
    const cache = (read: typeof binary) => read.book.pivotCaches?.[0]
    expect(cache(xml)?.source).toMatchObject({ book: 'ledger.xlsx' })
    expect(cache(binary)?.source).toEqual(cache(xml)?.source)
    expect(cache(binary)?.records).toStrictEqual(cache(xml)?.records)
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it('keeps the linked ranges of a consolidated PivotCache through an XLSX save', async () => {
    const [binary, xml] = await pair('pivot-source-consolidation-linked')
    const source = xml.book.pivotCaches?.[0].source
    expect(source).toMatchObject({ type: 'consolidation', ranges: [{ sheet: 'North' }, { sheet: 'Data Scope', book: expect.stringContaining('pivot-ledger.xlsx') }] })
    expect(binary.book.pivotCaches?.[0].source).toEqual(source)
    expect(binary.book.pivotCaches?.[0].records).toEqual(xml.book.pivotCaches?.[0].records)
    const saved = await writeXlsx({ book: binary.book, result: binary.cached })
    expect((await readXlsx(saved.data)).book.pivotCaches?.[0].source).toEqual(source)
  })

  it('reads the saved OLAP PivotTable and cache Excel wrote against an offline cube', async () => {
    // Excel 16.0.20430 created this view against the Microsoft Press Cube Test EN offline cube.
    const [binary, xml] = await pair('olap-offline-cube')
    const actual = binary.book.pivotCaches?.[0]
    const expected = xml.book.pivotCaches?.[0]
    expect(expected?.fields.map(({ name }) => name)).toEqual(['[Products].[Product].[Product]', '[Measures].[Sales]'])
    expect(actual).toStrictEqual({ ...expected, id: actual?.id })
    expect(binary.book.sheets[0].pivotTables).toStrictEqual(xml.book.sheets[0].pivotTables?.map((pivot) => ({ ...pivot, cache: actual?.id })))
    expect(binary.book.sheets[0].cells).toStrictEqual(xml.book.sheets[0].cells)
    expect(binary.cached).toStrictEqual(xml.cached)
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it('reads the records and view of a connection-backed PivotTable', async () => {
    const [binary, xml] = await pair('pivot-external-text')
    const actual = binary.book.pivotCaches?.[0]
    const expected = xml.book.pivotCaches?.[0]
    expect(expected?.source).toEqual({ type: 'external', connection: 1 })
    expect(expected?.records).toHaveLength(4)
    expect(actual).toStrictEqual({ ...expected, id: actual?.id })
    const pivot = binary.book.sheets[0].pivotTables?.[0]
    expect(pivot).toStrictEqual({ ...xml.book.sheets[0].pivotTables?.[0], cache: actual?.id })
    expect(binary.cached).toStrictEqual(xml.cached)
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it('keeps a plain text query’s data and reports its connection', async () => {
    const [binary, xml] = await pair('source-query')
    expect(binary).toStrictEqual(xml)
    expect(binary.book.sheets[0].cells.B5).toEqual({ value: 20 })
    expect(binary.issues.some(({ code }) => code === 'connection')).toBe(true)
  })

  it.each([
    ['source-ole-pivot', true],
    ['pivot-refresh-ole-off', false],
    ['pivot-refresh-odbc-off', false],
    ['pivot-refresh-sheet-legacy-off', false],
    ['pivot-refresh-sheet-legacy-on', true],
    ['pivot-refresh-sheet-modern-off', false],
    ['pivot-refresh-sheet-modern-on', true],
  ] as const)('reads the saved refresh setting in %s', async (name, refresh) => {
    const [binary, xml] = await pair(name)
    expect(xml.book.pivotCaches?.[0].enableRefresh ?? true).toBe(refresh)
    const actual = binary.book.pivotCaches?.[0]
    expect(actual).toStrictEqual({ ...xml.book.pivotCaches?.[0], id: actual?.id })
    expect(binary.book.sheets[0].pivotTables?.[0]).toStrictEqual({ ...xml.book.sheets[0].pivotTables?.[0], cache: actual?.id })
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it('associates several caches with their connections, including a cache saved without records', async () => {
    const [binary, xml] = await pair('pivot-external-many')
    const caches = binary.book.pivotCaches!
    expect(caches.map(({ source }) => source)).toEqual([1, 2, 3].map((connection) => ({ type: 'external', connection })))
    expect(caches.map(({ records }) => records?.length)).toEqual([4, 3, undefined])
    expect(caches.map((cache, index) => ({ ...cache, id: xml.book.pivotCaches![index].id }))).toStrictEqual(xml.book.pivotCaches)
    const views = binary.book.sheets[0].pivotTables!
    expect(views.map(({ name, cache }) => [name, caches.find(({ id }) => id === cache)?.source]))
      .toEqual([2, 1, 0].map((index) => ['Report' + index, caches[index].source]))
    expect(binary.cached).toStrictEqual(xml.cached)
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it('keeps the worksheet and the other caches when a PivotCache stream ends early', async () => {
    const [binary, xml] = await pair('pivot-first-cache-short')
    const baseline = await fixture('pivot-external-many').then(readXls)
    const sheet = binary.book.sheets[0]
    expect(sheet.cells).toStrictEqual(xml.book.sheets[0].cells)
    expect(binary.book.pivotCaches).toStrictEqual(baseline.book.pivotCaches?.slice(1))
    expect(sheet.pivotTables).toStrictEqual(baseline.book.sheets[0].pivotTables?.filter(({ name }) => name !== 'Report0'))
    expect(binary.issues).toContainEqual(expect.objectContaining({ code: 'pivotTable', sheet: 'Summary' }))
  })

  it('keeps time-only and phantom dates from a 1900 PivotCache', async () => {
    const [binary, xml] = await pair('date-serial-1900-dates')
    const field = (read: typeof binary) => read.book.pivotCaches?.[0].fields[0]
    expect(field(binary)?.items).toEqual([
      '1900-01-00T00:00:00', '1900-01-00T06:00:00', '1900-01-00T12:00:00',
      '1900-01-01T00:00:00', '1900-02-28T00:00:00', '1900-02-29T00:00:00',
      '1900-03-01T00:00:00', '1900-03-01T13:55:12',
    ].map((date) => ({ date })))
    // Excel's XLSX export shifts the stored dates before March. The XLS values retain their original serials.
    expect(field(xml)?.items?.[5]).toEqual({ date: '1900-02-28T00:00:00' })
    expect(binary.book.pivotCaches?.[0].records).toEqual(xml.book.pivotCaches?.[0].records)
  })

  it('reads dates and time-only values in a 1904 PivotCache', async () => {
    const [binary, xml] = await pair('date-serial-1904-times')
    expect(binary.book.dateSystem).toBe(1904)
    const cache = binary.book.pivotCaches?.[0]
    const reference = xml.book.pivotCaches?.[0]
    expect(cache?.fields[0].items).toStrictEqual(reference?.fields[0].items)
    expect(cache?.fields[0].items?.[4]).toEqual({ date: '1904-02-29T00:00:00' })
    expect(cache?.records).toEqual(reference?.records)
  })

  it.each(['legacy-tabular', 'legacy-outline', 'legacy-mixed', 'legacy-excel2003', 'modern-include'])(
    'matches Excel’s layouts and new-item filter settings in %s',
    async (name) => {
      const [binary, xml] = await pair('pivot-' + name)
      const table = binary.book.sheets[0].pivotTables?.[0]
      const expected = xml.book.sheets[0].pivotTables?.[0]
      expect(expected?.fields[0].includeNewItems).toBe(true)
      expect(table).toStrictEqual({ ...expected, cache: table?.cache })
    },
  )

  it.each(['modern-no-cache', 'modern-no-cache-types'])('reads the field kinds and dictionaries in %s', async (name) => {
    const [binary, xml] = await pair('pivot-' + name)
    const cache = binary.book.pivotCaches?.[0]
    const expected = xml.book.pivotCaches?.[0]
    expect(expected?.records).toBeUndefined()
    expect(expected?.fields[2].facts).toEqual(name === 'modern-no-cache' ? { types: ['number'], integer: true } : { types: ['number'] })
    expect(cache).toStrictEqual({ ...expected, id: cache?.id })
  })

  it.each(['retained', 'retained-calculated'])('reads stale entries alongside live cache items in %s', async (name) => {
    const [binary, xml] = await pair('pivot-' + name)
    const cache = binary.book.pivotCaches?.[0]
    const expected = xml.book.pivotCaches?.[0]
    const field = expected?.fields[0]
    expect(field?.unused?.map((index) => field.items?.[index])).toEqual(['Stale'])
    expect(cache).toStrictEqual({ ...expected, id: cache?.id })
  })

  it.each([['data', 2], ['label', 3], ['whole', 8], ['total', 2], ['alignment', 4]] as const)(
    'reads the formatted PivotTable areas for %s cells',
    async (name, count) => {
      const [binary, xml] = await pair('pivot-format-' + name)
      const actual = binary.book.sheets[0].pivotTables?.[0]
      const expected = xml.book.sheets[0].pivotTables?.[0]
      expect(expected?.formats).toHaveLength(count)
      expect(actual).toStrictEqual({ ...expected, cache: actual?.cache })
      expect(binary.issues).toStrictEqual(xml.issues)
    },
  )

  it.each([['format-condition', 'field'], ['cf-data', 'data'], ['cf-selection', undefined]] as const)('reads the PivotTable scope of %s', async (name, scope) => {
    const [binary, xml] = await pair('pivot-' + name)
    const rules = (read: typeof binary) => read.book.sheets[0].conditionalFormats
    expect(rules(xml)?.[0].pivot?.pivot).toBe('Summary')
    expect(rules(xml)?.[0].pivot?.scope).toBe(scope)
    expect(rules(binary)).toStrictEqual(rules(xml))
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it.each(['basic-amount-sum', 'basic-page-one', 'basic-source-name', 'basic-source-table'])(
    'reads the source, page selections and lines of Excel’s %s PivotTable',
    async (name) => {
      const [binary, xml] = await pair('pivot-' + name)
      const table = binary.book.sheets.find((sheet) => sheet.name === 'P')?.pivotTables?.[0]
      const expected = xml.book.sheets.find((sheet) => sheet.name === 'P')?.pivotTables?.[0]
      expect(table).toBeDefined()
      expect(table?.columnLines).toStrictEqual(expected?.columnLines)
      expect(table?.rowLines).toStrictEqual(expected?.rowLines)
      expect(table?.pages).toStrictEqual(expected?.pages)
      expect(table?.location.pageColumns).toEqual(expected?.location.pageColumns)
      expect(table?.location.pageRows).toEqual(expected?.location.pageRows)
      expect(binary.book.pivotCaches?.[0].source).toEqual(xml.book.pivotCaches?.[0].source)
      expect(binary.book.pivotCaches?.[0].records).toHaveLength(48)
      expect(binary.book.pivotCaches?.[0].records?.some((row) => typeof row[3] === 'object' && row[3] !== null && 'error' in row[3])).toBe(true)
      expect(binary.issues.filter((issue) => issue.code === 'pivotTable')).toEqual([])
    },
  )

  it.each(['qty-edges', 'dates-columns', 'items-two-fields', 'months-days'])(
    'reads the %s cache grouping and records saved by Excel',
    async (name) => {
      const [binary, xml] = await pair('pivot-group-held-' + name)
      const cache = binary.book.pivotCaches?.[0]
      const expected = xml.book.pivotCaches?.[0]
      expect(cache?.fields.map((field) => field.group)).toStrictEqual(expected?.fields.map((field) => field.group))
      expect(cache?.records?.slice(0, 2)).toStrictEqual(expected?.records?.slice(0, 2))
      expect(cache?.records).toHaveLength(60)
      expect(binary.book.sheets.find((sheet) => sheet.name === 'P')?.pivotTables).toHaveLength(1)
      expect(binary.issues.filter((issue) => issue.code === 'pivotTable')).toEqual([])
    },
  )

  it.each([
    'layout-held-captions-outline', 'layout-held-form-mixed-repeat', 'layout-held-subtotals-none-blank',
    'sort-held-region-down', 'sort-held-second-value', 'sort-held-tabular-value',
  ])('reads field layout and ordering in Excel’s %s PivotTable', async (name) => {
    const [binary, xml] = await pair('pivot-' + name)
    const table = (read: typeof binary) => read.book.sheets.find((sheet) => sheet.name === 'P')?.pivotTables?.[0]
    expect(table(binary)?.layout).toEqual(table(xml)?.layout)
    expect(table(binary)?.fields).toStrictEqual(table(xml)?.fields)
  })

  it.each(['pivot-calculated', 'pivot-item-number'])('reads the formulas of Excel’s %s cache', async (name) => {
    const [binary, xml] = await pair(name)
    const actual = binary.book.pivotCaches?.[0]
    const expected = xml.book.pivotCaches?.[0]
    expect(actual?.fields.filter((field) => field.formula)).toEqual(expected?.fields.filter((field) => field.formula))
    expect(actual?.calculatedItems).toStrictEqual(expected?.calculatedItems)
    expect(binary.book.sheets.find((sheet) => sheet.name === 'P')?.pivotTables).toHaveLength(1)
    expect(binary.issues.filter((issue) => issue.code === 'pivotTable')).toEqual([])
  })

  it.each(['date-after-equal', 'top-count-ties', 'label-contains-long', 'value-between-average', 'bottom-percent-inner'])(
    'reads the saved %s PivotTable filter',
    async (name) => {
      const [binary, xml] = await pair('pivot-filter-' + name)
      const filters = (read: typeof binary) => read.book.sheets.find((sheet) => sheet.name === 'P')?.pivotTables?.[0]?.filters
      expect(filters(xml)).toHaveLength(1)
      expect(filters(binary)).toStrictEqual(filters(xml))
    },
  )

  it('keeps a relative quarter filter stored in XLS when Excel removes it on reopening', async () => {
    const [binary, xml] = await pair('pivot-filter-date-quarter-4')
    const filters = (read: typeof binary) => read.book.sheets.find((sheet) => sheet.name === 'P')?.pivotTables?.[0]?.filters
    expect(filters(binary)).toEqual([{ field: 2, type: 'Q4' }])
    expect(filters(xml)).toBeUndefined()
  })

  it('reads LibreOffice’s mini-stream and text split across Continue records', async () => {
    const mini = await readXls(await fixture('libreoffice-mini'))
    expect(mini.book.sheets[0].cells.A2).toEqual({ formula: 'SUM(A1,2)' })
    expect(mini.cached.sheets[0].cells.A2).toEqual({ value: 5 })
    const { book, cached, issues } = await readXls(await fixture('libreoffice-continued'))
    const text = 'a'.repeat(5000) + 'Я 😀 é ' + 'я'.repeat(9000) + ' tail '
    expect(book.sheets[0].name).toBe('Русский')
    expect(book.sheets[0].cells.A1.value).toBe(text)
    expect(cached.sheets[0].cells.A2.value).toBe(Array.from(text).length)
    expect(cached.sheets[0].cells.A3.value).toBe('Привет')
    expect(issues).toEqual([])
  })

  it('reads volume paths and web URLs in linked workbooks as Excel saves them', async () => {
    const [binary, xml] = await pair('targets')
    expect(binary).toStrictEqual(xml)
  })

  it.each(['validation-invalid-reference', 'validation-list-comma'])('reads Excel’s %s rules', async (name) => {
    const [binary, xml] = await pair(name)
    expect(binary.book.sheets.map((sheet) => sheet.validations)).toStrictEqual(xml.book.sheets.map((sheet) => sheet.validations))
    expect(binary.book.locale).toEqual(xml.book.locale)
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it('reads LibreOffice’s external URLs, internal targets and locations within an external document', async () => {
    const { book, issues } = await readXls(await fixture('libreoffice-links'))
    const cells = book.sheets[0].cells
    expect(cells.B1.link).toEqual({ url: 'https://example.org/x' })
    expect(cells.B2.link).toEqual({ location: 'T!B3' })
    expect(cells.B3.link).toEqual({ url: 'https://example.org/doc', location: 'part' })
    expect(cells.B4.link).toEqual({ url: 'mailto:bob@example.org' })
    expect(issues).toEqual([])
  })

  it.each(['text-range-links', 'link-targets'])('reads Excel’s hyperlink targets, ranges and tips in %s', async (name) => {
    const [binary, xml] = await pair(name)
    expect(binary).toStrictEqual(xml)
  })

  it('reads a BIFF5 workbook with 8-bit strings', async () => {
    const { book } = await readXls(await fixture('biff5'))
    expect(book.sheets[0].name).toBe('Old')
    expect(book.sheets[0].cells.A1.value).toBe(7)
    expect(book.sheets[0].cells.B1.value).toBe('Café')
  })

  it.each([
    [437, 'Café'], [850, 'født'], [852, 'Łódź'], [866, 'шахт'],
  ] as const)('honors the declared DOS code page %i in a BIFF4 sheet', async (page, text) => {
    const { book } = await fixture('codepage-' + page).then(readXls)
    expect(book.sheets[0].cells.B1).toMatchObject({ value: text })
  })

  it('recognizes the default number format regardless of its spelling in older files', async () => {
    const bytes = await fixture('early-biff4')
    const expected = await readXls(bytes)
    const at = bytes.indexOf('General')
    expect(at).toBeGreaterThan(0)
    bytes.write('GENERAL', at, 'ascii')
    expect(await readXls(bytes)).toStrictEqual(expected)
  })

  it('reads BIFF5 links and the cell cache carried by its worksheet', async () => {
    const [binary, xml] = await pair('biff5-external-areas')
    expect(xml.book.externalBooks?.map((link) => link.target)).toEqual(['outside-legacy.xls', 'outside.xlsx'])
    expect(binary.book).toStrictEqual(xml.book)
    expect(binary.cached).toStrictEqual(xml.cached)
    expect(binary.issues).toEqual([])
  })

  it('resolves BIFF5 names and caches with each worksheet’s own link table', async () => {
    const [binary, xml] = await pair('biff5-external-names')
    expect(binary.book.sheets).toStrictEqual(xml.book.sheets)
    expect(binary.book.names).toStrictEqual(xml.book.names)
    expect(binary.cached).toStrictEqual(xml.cached)
    const expected = xml.book.externalBooks?.[0]
    expect(expected?.names).toHaveLength(3)
    // Excel's XLSX export resets this sheetId to 0. The BIFF5 EXTERNNAME points to Other in its local link table.
    const names = expected?.names?.map((name) => name.name === 'LocalRate' ? { ...name, sheet: 1 } : name)
    expect(binary.book.externalBooks).toStrictEqual([{ ...expected, names }])
    expect(binary.issues).toEqual([])
  })

  it.each([2, 3, 4])('reads a standalone BIFF%d worksheet and its formulas', async (version) => {
    const [binary, xml] = await pair('early-biff' + version)
    const sheet = binary.book.sheets[0]
    const expected = xml.book.sheets[0]
    expect(sheet.name).toBe('Sheet1')
    expect(binary.book.dateSystem).toEqual(xml.book.dateSystem)
    expect(binary.book.font).toStrictEqual(xml.book.font)
    expect(sheet.columns).toStrictEqual(expected.columns)
    expect(binary.book.styles).toStrictEqual(xml.book.styles)
    expect(sheet.cells.A1).toStrictEqual(expected.cells.A1)
    expect(sheet.cells.A2).toStrictEqual(expected.cells.A2)
    expect(sheet.cells.A3).toStrictEqual(expected.cells.A3)
    expect(sheet.cells.A4).toMatchObject({ value: 1.25, style: 0 })
    expect(sheet.cells.B4).toStrictEqual(expected.cells.B4)
    // The stored code page is Windows-1252. Excel for Mac decodes these old files in its current locale.
    expect(sheet.cells.B1.value).toBe('Café')
    expect(binary.cached.sheets[0].cells).toStrictEqual(xml.cached.sheets[0].cells)
  })

  it.each(['gcw-overlapping', 'gcw-with-gaps', 'gcw-without-standard', 'legacy-standard-only'])(
    'reads the saved BIFF4 column widths in %s',
    async (name) => {
      const [binary, xml] = await pair(name)
      const actual = binary.book.sheets[0]
      const expected = xml.book.sheets[0]
      expect(actual.columns).toStrictEqual(expected.columns)
      expect(actual.defaultColumnWidth).toBe(name === 'gcw-without-standard' ? undefined : expected.defaultColumnWidth)
    },
  )

  it.each(['biff2-column-only', 'biff2-column-middle'])('reads BIFF2 column formats beside their widths in %s', async (name) => {
    const [binary, xml] = await pair(name)
    const actual = binary.book.sheets[0]
    const expected = xml.book.sheets[0]
    expect(actual.columns?.B?.width).toBe(expected.columns?.B?.width)
    const columns = (read: typeof binary) => Object.fromEntries(Object.entries(read.book.sheets[0].columns ?? {}).map(([range, col]) => [
      range, { ...col, style: col.style === undefined ? undefined : read.book.styles?.[col.style], numberFormat: undefined },
    ]))
    expect(columns(binary)).toStrictEqual(columns(xml))
    expect(actual.columns?.B?.numberFormat).toBe('0.00')
    expect(binary.book.styles?.[actual.cells.B1.style ?? -1]).toEqual(xml.book.styles?.[expected.cells.B1.style ?? -1])
  })

  it('reads BIFF5 rich text stored directly in an RSTRING cell', async () => {
    const [binary, xml] = await pair('legacy-rich')
    expect(binary.book.sheets[0].cells.A1).toStrictEqual(xml.book.sheets[0].cells.A1)
    expect(binary.book.sheets[0].cells.A1.richText).toBeDefined()
    expect(binary.issues).toStrictEqual(xml.issues)
  })

  it('reads a long note across BIFF3 NOTE records', async () => {
    const [binary, xml] = await pair('legacy-note')
    const note = binary.book.sheets[0].cells.A1.note
    expect(note?.text).toHaveLength(2147)
    expect(note).toEqual(xml.book.sheets[0].cells.A1.note)
  })

  it('honors a sliced byte view and the same limits as XLSB', async () => {
    const bytes = await fixture('probe')
    const padded = Buffer.concat([Buffer.alloc(11), bytes, Buffer.alloc(3)])
    await expect(readXls(padded.subarray(11, -3))).resolves.toStrictEqual(await readXls(bytes))
    await expect(readXls(bytes, { limits: { entryBytes: 10 } })).rejects.toThrow(/Workbook.*entryBytes/)
    await expect(readXls(bytes, { limits: { totalBytes: 1 } })).rejects.toThrow(/totalBytes/)
    await expect(readXls(bytes, { limits: { totalBytes: 15_000 } })).rejects.toThrow(/DOCUMENTSUMMARYINFORMATION.*totalBytes/)
    await expect(readXls(bytes, { limits: { totalBytes: 19_000 } })).rejects.toThrow(/XML.*totalBytes/)
    await expect(readXls(bytes.subarray(0, 16))).rejects.toThrow(/compound file.*truncated header/)
    await expect(readXls(await fixture('probe', 'xlsx'))).rejects.toThrow(/compound file/)
  })

  it('uses the declared Workbook length when the FAT links another sector after it', async () => {
    const bytes = await fixture('probe')
    const expected = await readXls(bytes)
    const sectorSize = 512
    const directory = bytes.readUInt32LE(48)
    const workbook = (directory + 1) * sectorSize + 128
    const start = bytes.readUInt32LE(workbook + 116)
    const count = Math.ceil(bytes.readUInt32LE(workbook + 120) / sectorSize)
    const fat = (bytes.readUInt32LE(76) + 1) * sectorSize
    let last = start
    let previous = start
    for (let index = 1; index < count; index++) {
      previous = last
      last = bytes.readUInt32LE(fat + last * 4)
    }
    expect(bytes.readUInt32LE(fat + last * 4)).toBe(0xfffffffe)
    bytes.writeUInt32LE(directory, fat + last * 4)
    expect(await readXls(bytes)).toStrictEqual(expected)

    // The directory is still refused if the Workbook would read any of its bytes.
    bytes.writeUInt32LE(directory, fat + previous * 4)
    await expect(readXls(bytes)).rejects.toThrow('overlapping sector')
  })

  it.each(['extra-sector', 'short', 'overlap'])('reads the workbook with an optional property stream marked %s', async (kind) => {
    const bytes = await fixture('probe')
    const expected = await readXls(bytes)
    const sectorSize = 512
    const directory = bytes.readUInt32LE(48)
    const properties = (directory + 1) * sectorSize + 3 * 128
    const fat = (bytes.readUInt32LE(76) + 1) * sectorSize
    const start = bytes.readUInt32LE(properties + 116)
    if (kind === 'extra-sector') {
      const count = Math.ceil(bytes.readUInt32LE(properties + 120) / sectorSize)
      let last = start
      for (let index = 1; index < count; index++) last = bytes.readUInt32LE(fat + last * 4)
      bytes.writeUInt32LE(directory, fat + last * 4)
    } else if (kind === 'short') bytes.writeUInt32LE(0xfffffffe, fat + start * 4)
    else bytes.writeUInt32LE(directory, properties + 116)

    const actual = await readXls(bytes)
    expect(actual.book).toStrictEqual(expected.book)
    expect(actual.cached).toStrictEqual(expected.cached)
    if (kind === 'extra-sector') expect(actual.issues).toStrictEqual(expected.issues)
    else expect(actual.issues).toContainEqual({ code: 'part', message: 'invalid optional workbook properties were left out' })
  })
})
