/**
 * The files of the `keep` step: books the writer writes and Excel gives what the workbook model does not hold
 * (`MADE`: charts, a chart sheet, pictures, a pivot table, formatted runs, theme colours, the print area, page breaks,
 * a note, an AutoFilter with its criteria), and the edits written over each of them by `writeKept` (`EDITS`), which
 * Excel then opens. What a file holds of those parts is counted by `partsOf`. `MADE_PROBES` asks Excel what it answers
 * to an AutoFilter over a table and a table over an AutoFilter (`PROBES`, the engine's answers).
 */
import type { BookInput, Session, StructureOp } from '../../../src/contract.js'
import { compute } from '../../../src/index.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { openArchive } from '../../../src/xlsx/zip/read.js'
import { canonical } from '../../../src/syntax/print.js'
import { openSource } from '../../../src/xlsx/keep/source.js'
import { ContentTypes } from '../../../src/xlsx/keep/xml.js'
import { parseFormula } from '../../../src/xlsx/read/formulas.js'
import { attributes, element, elements } from '../../../src/xlsx/xml/scan.js'
import { readText } from '../../../src/xlsx/xml/text.js'
import { zip } from '../../../src/xlsx/zip/write.js'

const REGIONS = ['North', 'South', 'East', 'West']

/** The book every file starts from: amounts by region and month on Data, with a total; an empty sheet for a pivot table; a plain sheet. */
export function baseBook(): BookInput {
  const cells: BookInput['sheets'][number]['cells'] = { A1: { value: 'Region' }, B1: { value: 'Amount' }, C1: { value: 'Month' }, E1: { value: 'Total' }, E2: { formula: 'SUM(B2:B21)' } }
  for (let r = 2; r <= 21; r++) {
    cells[`A${r}`] = { value: REGIONS[r % 4] }
    cells[`B${r}`] = { value: r * 7 }
    cells[`C${r}`] = { value: `M${(r % 3) + 1}` }
  }
  return { sheets: [{ name: 'Data', cells }, { name: 'Pivot', cells: {} }, { name: 'Plain', cells: { A1: { value: 'plain' }, A2: { formula: 'Data!E2*2' } } }] }
}

/**
 * A file Excel makes: the features its script adds (`excel-keep.applescript`), whether it holds a pivot table or
 * sparklines the step writes into it, a table on Plain (A1:A2), and widths of Data's columns.
 */
export interface Made {
  id: string
  features: string[]
  pivot?: boolean
  sparklines?: boolean
  table?: boolean
  columns?: Record<string, { width: number }>
}

export const MADE: Made[] = [
  { id: 'chart', features: ['chart', 'theme'] },
  { id: 'picture', features: ['picture', 'rich'] },
  { id: 'pivot', features: ['chart'], pivot: true },
  { id: 'print', features: ['chart', 'picture', 'print', 'breaks', 'rich', 'theme'] },
  { id: 'chartsheet', features: ['chartsheet', 'note', 'rich'] },
  { id: 'all', features: ['chart', 'chartsheet', 'picture', 'rich', 'theme', 'print', 'note'], pivot: true },
  { id: 'criteria', features: ['autofilter', 'criteria'] },
]

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

/**
 * A pivot table of the amounts by region on the sheet Pivot (A3), its cache over Data!A1:C21 refreshed on loading:
 * Excel's object model on the Mac makes no pivot cache, so the parts are written as Excel writes them and Excel fills
 * the table when it opens the file.
 */
export async function withPivot(data: Uint8Array): Promise<Uint8Array> {
  const a = openArchive(data)
  const parts = new Map<string, string>()
  for (const n of a.names) parts.set(n, (await a.text(n))!)
  const edit = (n: string, f: (x: string) => string) => parts.set(n, f(parts.get(n)!))
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
  const items = REGIONS.map((r) => `<s v="${r}"/>`).join('')
  parts.set(
    'xl/pivotCache/pivotCacheDefinition1.xml',
    `${head}<pivotCacheDefinition xmlns="${MAIN}" xmlns:r="${REL}" saveData="0" refreshOnLoad="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="0"><cacheSource type="worksheet"><worksheetSource ref="A1:C21" sheet="Data"/></cacheSource><cacheFields count="3"><cacheField name="Region" numFmtId="0"><sharedItems count="4">${items}</sharedItems></cacheField><cacheField name="Amount" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="14" maxValue="147"/></cacheField><cacheField name="Month" numFmtId="0"><sharedItems/></cacheField></cacheFields></pivotCacheDefinition>`,
  )
  const rowItems = `<i><x/></i>${[1, 2, 3].map((k) => `<i><x v="${k}"/></i>`).join('')}<i t="grand"><x/></i>`
  parts.set(
    'xl/pivotTables/pivotTable1.xml',
    `${head}<pivotTableDefinition xmlns="${MAIN}" name="PivotTable1" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"><location ref="A3:B8" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="3"><pivotField axis="axisRow" showAll="0"><items count="5">${[0, 1, 2, 3].map((k) => `<item x="${k}"/>`).join('')}<item t="default"/></items></pivotField><pivotField dataField="1" showAll="0"/><pivotField showAll="0"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><rowItems count="5">${rowItems}</rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField name="Sum of Amount" fld="1" baseField="0" baseItem="0"/></dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`,
  )
  parts.set('xl/pivotTables/_rels/pivotTable1.xml.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>`)
  parts.set('xl/worksheets/_rels/sheet2.xml.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>`)
  edit('[Content_Types].xml', (x) => x.replace('</Types>', `<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="${TYPE}.pivotCacheDefinition+xml"/><Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="${TYPE}.pivotTable+xml"/></Types>`))
  edit('xl/_rels/workbook.xml.rels', (x) => x.replace('</Relationships>', `<Relationship Id="rIdPC1" Type="${REL}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>`))
  edit('xl/workbook.xml', (x) => x.replace(/(<calcPr[^>]*\/>)/, '$1<pivotCaches><pivotCache cacheId="1" r:id="rIdPC1"/></pivotCaches>'))
  return zip([...parts].map(([name, text]) => ({ name, data: new TextEncoder().encode(text) })))
}

/**
 * A group of two sparklines on Data, G2 over B2:B11 and G3 over B12:B21, in the sheet's extension as Excel writes them:
 * Excel's object model on the Mac makes none.
 */
export async function withSparklines(data: Uint8Array): Promise<Uint8Array> {
  const a = openArchive(data)
  const parts = new Map<string, string>()
  for (const n of a.names) parts.set(n, (await a.text(n))!)
  const colors = ['Series', 'Negative', 'Axis', 'Markers', 'First', 'Last', 'High', 'Low'].map((c) => `<x14:color${c} rgb="FF376092"/>`).join('')
  const line = (f: string, cell: string) => `<x14:sparkline><xm:f>${f}</xm:f><xm:sqref>${cell}</xm:sqref></x14:sparkline>`
  const ext = `<extLst><ext uri="{05C60535-1F16-4FD2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:sparklineGroup displayEmptyCellsAs="gap">${colors}<x14:sparklines>${line('Data!B2:B11', 'G2')}${line('Data!B12:B21', 'G3')}</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>`
  parts.set('xl/worksheets/sheet1.xml', parts.get('xl/worksheets/sheet1.xml')!.replace('</worksheet>', `${ext}</worksheet>`))
  return zip([...parts].map(([name, text]) => ({ name, data: new TextEncoder().encode(text) })))
}

/** The book's file as the step sends it to Excel to be made. */
export async function baseFile(m: Made): Promise<Uint8Array> {
  const book = baseBook()
  if (m.table) book.sheets[2].tables = [{ name: 'Plain', range: 'A1:A2', columns: [{ name: 'plain' }] }]
  if (m.columns) book.sheets[0].columns = m.columns
  const { data } = await writeXlsx({ book, result: compute(book) })
  return m.pivot ? withPivot(data) : m.sparklines ? withSparklines(data) : data
}

/** An edit written over a made file: what it does to a session over the file's book; over the made files `made` alone where it names them. */
export interface Edit {
  id: string
  what: string
  edit(s: Session): void
  made?: string[]
}

const ok = (r: { ok?: boolean }) => {
  if (r.ok === false) throw new Error(`the edit was refused: ${JSON.stringify(r)}`)
}

export const EDITS: Edit[] = [
  { id: 'values', what: 'amounts the chart and the pivot table show changed', edit: (s) => void s.setCells([2, 3, 4].map((r) => ({ sheet: 'Data', cell: `B${r}`, input: { value: r * 100 } }))) },
  {
    id: 'links',
    what: 'a link to another workbook added with the values the host read, and a formula typed over it by its file name',
    edit: (s) => {
      s.setExternalBooks([...(s.getBook().externalBooks ?? []), { target: 'Prices.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 3, B2: 'x' } }, { name: 'Rates 2', cells: { C3: 0.5 } }] }])
      ok(s.setCellText('Data', 'F1', "=[Prices.xlsx]Sheet1!A1*2+'[Prices.xlsx]Rates 2'!C3"))
    },
  },
  {
    id: 'formulas-styles',
    what: 'a column of formulas, bold headers, a fill, a number format',
    edit: (s) => {
      s.setCells([{ sheet: 'Data', cell: 'D1', input: { value: 'Share' } }, ...Array.from({ length: 20 }, (_, i) => ({ sheet: 'Data', cell: `D${i + 2}`, input: { formula: `B${i + 2}/$E$2` } }))])
      s.patchCells([{ sheet: 'Data', range: 'A1:E1', set: { style: { font: { bold: true }, fill: { color: '#DDEEFF' } } } }, { sheet: 'Data', range: 'D2:D21', set: { numberFormat: '0.0%' } }])
    },
  },
  {
    id: 'parts',
    what: 'a merged area, a conditional format, a validation rule, a hyperlink and a note',
    edit: (s) => {
      ok(s.editStructure({ op: 'mergeCells', sheet: 'Plain', range: 'C1:E2' }))
      const data = s.getBook().sheets.find((x) => x.name === 'Data')!
      s.setLayout([
        { sheet: 'Data', conditionalFormats: [...(data.conditionalFormats ?? []), { type: 'cellIs', ranges: ['B2:B21'], operator: 'greaterThan', formulas: ['100'], style: { font: { bold: true }, fill: { color: '#FFC7CE' } } }] },
        { sheet: 'Plain', validations: [{ range: 'B1:B10', type: 'whole', operator: 'between', formula1: '1', formula2: '9' }] },
      ])
      s.patchCells([{ sheet: 'Plain', cell: 'A1', set: { link: { url: 'https://example.com/plain' }, note: { author: 'xlcalc', text: 'written over the file' } } }])
    },
  },
  {
    id: 'structure',
    what: 'rows and columns inserted on the plain sheet, a sheet added, the plain sheet renamed',
    edit: (s) => {
      ok(s.editStructure({ op: 'insertRows', sheet: 'Plain', row: 1, count: 2 }))
      ok(s.editStructure({ op: 'insertColumns', sheet: 'Plain', column: 'A' }))
      ok(s.editStructure({ op: 'insertSheet', name: 'Added', before: 'Plain' }))
      s.setCells([{ sheet: 'Added', cell: 'A1', input: { formula: 'SEQUENCE(3)' } }, { sheet: 'Added', cell: 'B1', input: { value: 'new' } }])
      ok(s.editStructure({ op: 'renameSheet', sheet: 'Plain', name: 'Plain renamed' }))
    },
  },
  {
    id: 'below',
    what: 'rows inserted on Data below what the chart and the pivot table read, and values typed there',
    edit: (s) => {
      ok(s.editStructure({ op: 'insertRows', sheet: 'Data', row: 22, count: 3 }))
      s.setCells([{ sheet: 'Data', cell: 'A23', input: { value: 'later' } }, { sheet: 'Data', cell: 'B23', input: { value: 1 } }])
    },
  },
  {
    id: 'filter',
    what: "an AutoFilter put on Data over a row more than the file's, and one on Plain",
    edit: (s) => void s.setLayout([{ sheet: 'Data', autoFilter: { range: 'A1:C22' } }, { sheet: 'Plain', autoFilter: { range: 'A1:A2' } }]),
  },
  {
    id: 'filter-off',
    what: "the file's AutoFilter removed, its rows shown",
    edit: (s) => void s.setLayout([{ sheet: 'Data', autoFilter: null }, { sheet: 'Data', filtered: false }, { sheet: 'Data', rows: '1:21', set: { hidden: null } }]),
    made: ['criteria'],
  },
  {
    id: 'rich-sheet',
    what: 'the cells with formatted runs moved by a range to another sheet',
    edit: (s) => ok(s.editStructure({ op: 'moveRange', from: { sheet: 'Data', range: 'A2:A5' }, to: { sheet: 'Plain', cell: 'D1' } })),
    made: ['picture'],
  },
]

/** What a file holds of the parts the model does not have, counted: charts, chart sheets, drawings, pictures, pivot tables and caches, formatted strings, theme colours, the print area, page breaks, notes, and the sheets' AutoFilters where it has any. */
export async function partsOf(data: Uint8Array): Promise<Record<string, number>> {
  const a = openArchive(data)
  const names = a.names
  const count = (re: RegExp) => names.filter((n) => re.test(n)).length
  const text = async (n: string) => (await a.text(n)) ?? ''
  const strings = await text('xl/sharedStrings.xml')
  const styles = await text('xl/styles.xml')
  const wb = await text('xl/workbook.xml')
  let breaks = 0
  let filters = 0
  for (const n of names.filter((x) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(x))) {
    const xml = await text(n)
    breaks += (xml.match(/<brk /g) ?? []).length
    filters += (xml.match(/<autoFilter /g) ?? []).length
  }
  return {
    charts: count(/^xl\/charts\/chart\d+\.xml$/i),
    chartsheets: count(/^xl\/chartsheets\/sheet\d+\.xml$/i),
    drawings: count(/^xl\/drawings\/drawing\d+\.xml$/i),
    pictures: count(/^xl\/media\//i),
    pivotTables: count(/^xl\/pivotTables\/pivotTable\d+\.xml$/i),
    pivotCaches: count(/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/i),
    richStrings: (strings.match(/<si><r>/g) ?? []).length,
    themeColors: (styles.match(/theme="\d+"/g) ?? []).length,
    printAreas: (wb.match(/_xlnm\.Print_Area/g) ?? []).length,
    breaks,
    notes: count(/^xl\/comments\d+\.xml$/i),
    ...(filters ? { autoFilters: filters } : {}),
  }
}

/** Files Excel makes for the operations of the structure (`STRUCT`) beside `MADE`: a picture that sizes with its cells and one fixed, and a chart and a picture that only move. */
export const MADE_STRUCT: Made[] = [
  { id: 'sized', features: ['chart', 'picture', 'sized', 'floating'] },
  { id: 'moving', features: ['chart', 'chartmove', 'picture'] },
  { id: 'filter', features: ['autofilter'] },
  // the chart and the picture of moving over columns of other widths than Excel's default, and sparklines
  { id: 'wide', features: ['chart', 'chartmove', 'picture'], columns: { F: { width: 20 }, G: { width: 4 }, H: { width: 30 } } },
  { id: 'sparklines', features: [], sparklines: true },
]

/**
 * Files Excel makes to answer what it does with an AutoFilter beside a table (`excel-keep.applescript`): one asked over
 * Plain!A1:B5, where a table lies on A1:A2, and on Data, which has an AutoFilter over A1:C21, a table asked over
 * A1:C21, over B10:D12, inside the AutoFilter (A1:B10), around it (A1:D22), over it and rows below (A1:C25), over its
 * rows but the header (A2:C21), each predicted after the answers before it. Each line of an answer starts with the
 * feature's name and what Excel did, `done` or `error`.
 */
export const MADE_PROBES: Made[] = [
  { id: 'filter-tables', features: ['autofilter', 'overtable', 'tableover'], table: true },
  { id: 'filter-table-part', features: ['autofilter', 'tablepart'] },
  { id: 'filter-table-in', features: ['autofilter', 'tablein'] },
  { id: 'filter-table-around', features: ['autofilter', 'tablearound'] },
  { id: 'filter-table-long', features: ['autofilter', 'tablelong'] },
  { id: 'filter-table-low', features: ['autofilter', 'tablelow'] },
]

/**
 * The engine's answers to the probes, each line's start: the AutoFilter made; one over a table refused (`checkTables`);
 * a table over all of the AutoFilter's range made, the AutoFilter gone, one over part of it refused (`addTable`). The
 * engine refused every table over the AutoFilter until Excel's first answer, then took any such table in its place
 * until the second (`filter-table-part`, refused), then only one over its range alone until the fourth
 * (`filter-table-around`, made); the last two were predicted after it.
 */
export const PROBES: Record<string, string[]> = {
  'filter-tables': ['autofilter $A$1:$C$21', 'overtable error', 'tableover done; autofilter mode false; tables 1'],
  'filter-table-part': ['autofilter $A$1:$C$21', 'tablepart error'],
  'filter-table-in': ['autofilter $A$1:$C$21', 'tablein error'],
  'filter-table-around': ['autofilter $A$1:$C$21', 'tablearound done; autofilter mode false; tables 1'],
  'filter-table-long': ['autofilter $A$1:$C$21', 'tablelong done; autofilter mode false; tables 1'],
  'filter-table-low': ['autofilter $A$1:$C$21', 'tablelow error'],
}

/** An operation of the structure on a made file, as the session takes it and as `excel-keep-struct.applescript` does it: op, sheet, range, target, cell. */
export interface Struct {
  id: string
  made: string
  op: StructureOp
  excel: [op: string, sheet: string, range: string, target: string, to: string]
}

const rows = (sheet: string, row: number, count: number, op: 'insertRows' | 'deleteRows'): Pick<Struct, 'op' | 'excel'> => ({ op: { op, sheet, row, count }, excel: [op, sheet, `A${row}:A${row + count - 1}`, '', ''] })

const filter = (id: string, op: StructureOp, excel: Struct['excel']): Struct => ({ id: `filter-${id}`, made: 'filter', op, excel })

/** Columns of Data inserted or deleted, from `column` on, over a made file. */
const cols = (id: string, made: string, column: string, count: number, op: 'insertColumns' | 'deleteColumns'): Struct => {
  const last = String.fromCharCode(column.charCodeAt(0) + count - 1)
  return { id, made, op: { op, sheet: 'Data', column, count }, excel: [op, 'Data', `${column}1:${last}1`, '', ''] }
}

export const STRUCT: Struct[] = [
  { id: 'chart-insert-in', made: 'chart', ...rows('Data', 3, 2, 'insertRows') },
  { id: 'chart-delete-in', made: 'chart', ...rows('Data', 3, 2, 'deleteRows') },
  { id: 'chart-delete-series', made: 'chart', ...rows('Data', 2, 5, 'deleteRows') },
  { id: 'chart-delete-chart', made: 'chart', ...rows('Data', 1, 14, 'deleteRows') },
  { id: 'chart-insert-cols', made: 'chart', op: { op: 'insertColumns', sheet: 'Data', column: 'B' }, excel: ['insertColumns', 'Data', 'B1', '', ''] },
  { id: 'chart-rename', made: 'chart', op: { op: 'renameSheet', sheet: 'Data', name: 'Sales 2025' }, excel: ['renameSheet', 'Data', '', 'Sales 2025', ''] },
  { id: 'chart-move', made: 'chart', op: { op: 'moveRange', from: { sheet: 'Data', range: 'A1:B6' }, to: { sheet: 'Plain', cell: 'D1' } }, excel: ['moveRange', 'Data', 'A1:B6', 'Plain', 'D1'] },
  { id: 'picture-delete-picture', made: 'picture', ...rows('Data', 14, 7, 'deleteRows') },
  { id: 'sized-delete-picture', made: 'sized', ...rows('Data', 14, 7, 'deleteRows') },
  { id: 'sized-insert-above', made: 'sized', ...rows('Data', 2, 2, 'insertRows') },
  { id: 'moving-insert-in', made: 'moving', ...rows('Data', 3, 2, 'insertRows') },
  { id: 'pivot-source-deleted', made: 'pivot', op: { op: 'deleteSheet', sheet: 'Data' }, excel: ['deleteSheet', 'Data', '', '', ''] },
  { id: 'pivot-insert-above', made: 'pivot', ...rows('Pivot', 1, 2, 'insertRows') },
  { id: 'pivot-insert-source', made: 'pivot', ...rows('Data', 5, 1, 'insertRows') },
  { id: 'print-delete', made: 'print', ...rows('Data', 21, 1, 'deleteRows') },
  { id: 'print-insert', made: 'print', ...rows('Data', 5, 2, 'insertRows') },
  { id: 'print-delete-break-row', made: 'print', ...rows('Data', 12, 1, 'deleteRows') },
  { id: 'chartsheet-rename', made: 'chartsheet', op: { op: 'renameSheet', sheet: 'Data', name: 'Amounts' }, excel: ['renameSheet', 'Data', '', 'Amounts', ''] },
  { id: 'chartsheet-source-deleted', made: 'chartsheet', op: { op: 'deleteSheet', sheet: 'Data' }, excel: ['deleteSheet', 'Data', '', '', ''] },
  { id: 'all-insert-cols', made: 'all', op: { op: 'insertColumns', sheet: 'Data', column: 'A' }, excel: ['insertColumns', 'Data', 'A1', '', ''] },
  // the AutoFilter of Data over A1:C21
  { ...rows('Data', 3, 2, 'insertRows'), id: 'filter-insert-in', made: 'filter' },
  { ...rows('Data', 22, 1, 'insertRows'), id: 'filter-insert-below', made: 'filter' },
  { ...rows('Data', 1, 1, 'insertRows'), id: 'filter-insert-above', made: 'filter' },
  { ...rows('Data', 1, 1, 'deleteRows'), id: 'filter-delete-header', made: 'filter' },
  { ...rows('Data', 1, 21, 'deleteRows'), id: 'filter-delete-all', made: 'filter' },
  filter('delete-column', { op: 'deleteColumns', sheet: 'Data', column: 'B' }, ['deleteColumns', 'Data', 'B1', '', '']),
  filter('delete-columns-all', { op: 'deleteColumns', sheet: 'Data', column: 'A', count: 3 }, ['deleteColumns', 'Data', 'A1:C1', '', '']),
  filter('insert-column', { op: 'insertColumns', sheet: 'Data', column: 'B' }, ['insertColumns', 'Data', 'B1', '', '']),
  filter('insert-column-right', { op: 'insertColumns', sheet: 'Data', column: 'D' }, ['insertColumns', 'Data', 'D1', '', '']),
  filter('move', { op: 'moveRange', from: { sheet: 'Data', range: 'A1:C21' }, to: { sheet: 'Data', cell: 'H3' } }, ['moveRange', 'Data', 'A1:C21', 'Data', 'H3']),
  filter('move-sheet', { op: 'moveRange', from: { sheet: 'Data', range: 'A1:C21' }, to: { sheet: 'Plain', cell: 'D1' } }, ['moveRange', 'Data', 'A1:C21', 'Plain', 'D1']),
  filter('move-header', { op: 'moveRange', from: { sheet: 'Data', range: 'A1:C1' }, to: { sheet: 'Data', cell: 'H1' } }, ['moveRange', 'Data', 'A1:C1', 'Data', 'H1']),
  filter('move-over', { op: 'moveRange', from: { sheet: 'Data', range: 'H1:J21' }, to: { sheet: 'Data', cell: 'A1' } }, ['moveRange', 'Data', 'H1:J21', 'Data', 'A1']),
  filter('rename', { op: 'renameSheet', sheet: 'Data', name: 'Sales' }, ['renameSheet', 'Data', '', 'Sales', '']),
  filter('delete-sheet', { op: 'deleteSheet', sheet: 'Data' }, ['deleteSheet', 'Data', '', '', '']),
  // held out after the rules for the header row and for moves were fitted to the cases above
  { ...rows('Data', 1, 2, 'deleteRows'), id: 'filter-delete-head-two', made: 'filter' },
  filter('delete-first-column', { op: 'deleteColumns', sheet: 'Data', column: 'A' }, ['deleteColumns', 'Data', 'A1', '', '']),
  filter('move-over-end', { op: 'moveRange', from: { sheet: 'Data', range: 'H1:J5' }, to: { sheet: 'Data', cell: 'A17' } }, ['moveRange', 'Data', 'H1:J5', 'Data', 'A17']),
  filter('move-end-down', { op: 'moveRange', from: { sheet: 'Data', range: 'A15:C21' }, to: { sheet: 'Data', cell: 'A30' } }, ['moveRange', 'Data', 'A15:C21', 'Data', 'A30']),
  // the rules for columns: a drawing whose columns are all deleted, the size of one that moves without sizing by the
  // columns' widths; and the refusals for deleted cells of a pivot cache's source and of sparklines
  cols('sized-delete-picture-cols', 'sized', 'F', 3, 'deleteColumns'),
  cols('moving-delete-picture-cols', 'moving', 'F', 3, 'deleteColumns'),
  cols('moving-delete-from-col', 'moving', 'F', 1, 'deleteColumns'),
  cols('moving-insert-cols-in', 'moving', 'G', 2, 'insertColumns'),
  cols('wide-delete-from-col', 'wide', 'F', 1, 'deleteColumns'),
  cols('wide-delete-cols-in', 'wide', 'G', 1, 'deleteColumns'),
  cols('wide-insert-cols-in', 'wide', 'G', 2, 'insertColumns'),
  cols('wide-delete-picture-cols', 'wide', 'F', 3, 'deleteColumns'),
  cols('wide-delete-before', 'wide', 'D', 2, 'deleteColumns'),
  cols('pivot-delete-source-cells', 'pivot', 'A', 3, 'deleteColumns'),
  { id: 'sparklines-delete-cell', made: 'sparklines', ...rows('Data', 2, 1, 'deleteRows') },
  { id: 'sparklines-delete-data', made: 'sparklines', ...rows('Data', 12, 10, 'deleteRows') },
  cols('sparklines-delete-data-col', 'sparklines', 'B', 1, 'deleteColumns'),
  // held out after the rules for a drawing's size by the columns' widths, a pivot cache's source and a sparkline's cell were fitted to the cases above
  cols('wide-delete-wide-col', 'wide', 'H', 1, 'deleteColumns'),
  cols('wide-delete-to-col', 'wide', 'I', 1, 'deleteColumns'),
  cols('moving-delete-cols-in', 'moving', 'G', 2, 'deleteColumns'),
  { id: 'pivot-delete-source-rows', made: 'pivot', ...rows('Data', 1, 21, 'deleteRows') },
  { id: 'sparklines-delete-both', made: 'sparklines', ...rows('Data', 2, 2, 'deleteRows') },
]

/**
 * What a file's parts the model does not have refer to: the series of its charts (canonical, sorted), each sheet's
 * drawing anchors (kind, `editAs`, each corner's column and row with their offsets), its pivot tables' places and their
 * caches' sources, the built-in names, each sheet's page breaks and, where a sheet has them, its AutoFilter's range and its sparklines (cell and data). What `keep struct` compares between Excel's file and
 * the one `writeKept` writes.
 */
export async function refsOf(data: Uint8Array): Promise<Record<string, unknown>> {
  const src = await openSource(data)
  const types = new ContentTypes(src.types)
  const canon = (text: string) => {
    const f = parseFormula(text, { union: true })
    return f.ast ? canonical(f.ast) : text
  }
  const series: string[] = []
  for (const part of types.parts((t) => /drawingml\.chart\+xml$/i.test(t))) for (const f of elements((await src.archive.text(part)) ?? '', 'f')) series.push(canon(readText(f.body ?? '')))
  const sheetName = (i: number) => src.decls[src.sheets[i].decl].name
  const anchors: Record<string, string[]> = {}
  const breaks: Record<string, string> = {}
  const filters: Record<string, string> = {}
  for (const [i, s] of src.sheets.entries()) {
    const own = attributes(element(s.xml, 'autoFilter')?.attrs ?? '').ref
    if (own !== undefined) filters[sheetName(i)] = own
    for (const r of s.rels.filter((x) => x.kind === 'drawing' && !x.external)) {
      const xml = (await src.archive.text(r.target)) ?? ''
      const corner = (body: string, tag: string) => {
        const p = element(body, tag)?.body ?? ''
        return ['col', 'colOff', 'row', 'rowOff'].map((k) => element(p, k)?.body ?? '?').join(' ')
      }
      for (const local of ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'])
        for (const e of elements(xml, local)) (anchors[sheetName(i)] ??= []).push(`${local} ${attributes(e.attrs).editAs ?? ''}: ${corner(e.body ?? '', 'from')} / ${corner(e.body ?? '', 'to')}`)
    }
    const ids = (local: string) => elements(element(s.xml, local)?.body ?? '', 'brk').map((b) => attributes(b.attrs).id)
    if (ids('rowBreaks').length || ids('colBreaks').length) breaks[sheetName(i)] = `rows ${ids('rowBreaks').join(',')} columns ${ids('colBreaks').join(',')}`
  }
  const pivots: string[] = []
  for (const part of types.parts((t) => /pivotTable\+xml$/i.test(t))) pivots.push(`location ${attributes(element((await src.archive.text(part)) ?? '', 'location')?.attrs ?? '').ref}`)
  for (const part of types.parts((t) => /pivotCacheDefinition\+xml$/i.test(t))) {
    const a = attributes(element((await src.archive.text(part)) ?? '', 'worksheetSource')?.attrs ?? '')
    pivots.push(`source ${a.name !== undefined ? `name ${a.name}` : `${a.sheet ?? '-'}!${a.ref ?? '-'}`}`)
  }
  const sparklines: Record<string, string[]> = {}
  for (const [i, s] of src.sheets.entries())
    for (const x of elements(s.xml, 'sparkline')) (sparklines[sheetName(i)] ??= []).push(`${element(x.body ?? '', 'sqref')?.body ?? '?'} ${canon(readText(element(x.body ?? '', 'f')?.body ?? ''))}`)
  const names: string[] = []
  for (const d of elements(element(src.wb, 'definedNames')?.body ?? '', 'definedName')) {
    const a = attributes(d.attrs)
    if (/^_xlnm\./.test(a.name ?? '')) names.push(`${a.name}@${a.localSheetId === undefined ? '' : src.decls[Number(a.localSheetId)]?.name}=${canon(readText(d.body ?? ''))}`)
  }
  return { series: series.sort(), anchors, pivots, names: names.sort(), breaks, ...(Object.keys(filters).length ? { filters } : {}), ...(Object.keys(sparklines).length ? { sparklines } : {}) }
}
