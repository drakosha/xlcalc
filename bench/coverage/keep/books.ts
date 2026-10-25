/**
 * The files of the `keep` step: books the writer writes and Excel gives what the workbook model does not hold
 * (`MADE`: charts, a chart sheet, pictures, a pivot table, formatted runs, theme colours, the print area, page breaks,
 * a note), and the edits written over each of them by `writeKept` (`EDITS`), which Excel then opens. What a file holds
 * of those parts is counted by `partsOf`.
 */
import type { BookInput, Session } from '../../../src/contract.js'
import { compute } from '../../../src/index.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { openArchive } from '../../../src/xlsx/zip/read.js'
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

/** A file Excel makes: the features its script adds (`excel-keep.applescript`), and whether it holds a pivot table the step writes into it. */
export interface Made {
  id: string
  features: string[]
  pivot?: boolean
}

export const MADE: Made[] = [
  { id: 'chart', features: ['chart', 'theme'] },
  { id: 'picture', features: ['picture', 'rich'] },
  { id: 'pivot', features: ['chart'], pivot: true },
  { id: 'print', features: ['chart', 'picture', 'print', 'breaks', 'rich', 'theme'] },
  { id: 'chartsheet', features: ['chartsheet', 'note', 'rich'] },
  { id: 'all', features: ['chart', 'chartsheet', 'picture', 'rich', 'theme', 'print', 'note'], pivot: true },
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

/** The book's file as the step sends it to Excel to be made. */
export async function baseFile(m: Made): Promise<Uint8Array> {
  const book = baseBook()
  const { data } = await writeXlsx({ book, result: compute(book) })
  return m.pivot ? withPivot(data) : data
}

/** An edit written over a made file: what it does to a session over the file's book. */
export interface Edit {
  id: string
  what: string
  edit(s: Session): void
}

const ok = (r: { ok?: boolean }) => {
  if (r.ok === false) throw new Error(`the edit was refused: ${JSON.stringify(r)}`)
}

export const EDITS: Edit[] = [
  { id: 'values', what: 'amounts the chart and the pivot table show changed', edit: (s) => void s.setCells([2, 3, 4].map((r) => ({ sheet: 'Data', cell: `B${r}`, input: { value: r * 100 } }))) },
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
]

/** What a file holds of the parts the model does not have, counted: charts, chart sheets, drawings, pictures, pivot tables and caches, formatted strings, theme colours, the print area, page breaks, notes. */
export async function partsOf(data: Uint8Array): Promise<Record<string, number>> {
  const a = openArchive(data)
  const names = a.names
  const count = (re: RegExp) => names.filter((n) => re.test(n)).length
  const text = async (n: string) => (await a.text(n)) ?? ''
  const strings = await text('xl/sharedStrings.xml')
  const styles = await text('xl/styles.xml')
  const wb = await text('xl/workbook.xml')
  let breaks = 0
  for (const n of names.filter((x) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(x))) breaks += ((await text(n)).match(/<brk /g) ?? []).length
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
  }
}
