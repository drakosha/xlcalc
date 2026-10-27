import { withoutView } from './xlsx-fixtures.js'
/**
 * Hidden rows and columns and filters through xlsx: hidden lines read and written with their sizes, a line of size 0
 * taken as hidden, the filter mode read from each form a file holds a filter in (an autofilter with criteria, the mark of
 * an advanced filter, a table's filter) and written as Excel's mark with criteria every row meets; SUBTOTAL and
 * AGGREGATE compute as in the book written. Excel's answers to the same forms are checked in `test/xlsx-excel.test.ts`.
 */
import { describe, expect, it } from 'vitest'
import { compute, type BookInput, type SheetInput } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { KEEP_ALL } from '../src/xlsx/write/sheet.js'
import { codes, pack, partText as text, parts } from './xlsx-fixtures.js'

/** A1:A6 hold 1, 2, 4, 8, 16, 32 and B1:B6 keys; C1 to C4 SUBTOTAL and AGGREGATE over A1:A6. */
const ROWS = Array.from({ length: 6 }, (_, i) => `<c r="A${i + 1}"><v>${2 ** i}</v></c><c r="B${i + 1}" t="inlineStr"><is><t>${i % 3 ? 'x' : 'y'}</t></is></c>`)
const FORMULAS = ['SUBTOTAL(9,A1:A6)', 'SUBTOTAL(109,A1:A6)', 'AGGREGATE(9,5,A1:A6)', 'AGGREGATE(9,4,A1:A6)']

/** A sheet with rows 1 and 4 hidden and what goes around its data: a filter's parts. */
async function filtered(opts: { before?: string; after?: string; table?: string; workbook?: string }) {
  const rows = ROWS.map((cells, i) => `<row r="${i + 1}"${i % 3 ? '' : ' hidden="1"'}>${cells}${i < FORMULAS.length ? `<c r="C${i + 1}"><f>${FORMULAS[i]}</f></c>` : ''}</row>`)
  const f = parts({ sheets: [{ name: 'S', body: `${opts.before ?? ''}<sheetData>${rows.join('')}</sheetData>${opts.after ?? ''}${opts.table ? '<tableParts count="1"><tablePart r:id="rT"/></tableParts>' : ''}` }], workbook: { after: opts.workbook } })
  if (opts.table) {
    f['xl/worksheets/_rels/sheet1.xml.rels'] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>`
    f['xl/tables/table1.xml'] = `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:B6" headerRowCount="0">${opts.table}<tableColumns count="2"><tableColumn id="1" name="v"/><tableColumn id="2" name="k"/></tableColumns></table>`
  }
  const r = await readXlsx(await pack(f))
  const values = compute(r.book, { display: false }).sheets[0].cells
  return { filtered: r.book.sheets[0].filtered ?? false, values: FORMULAS.map((_, i) => values[`C${i + 1}`].value), issues: r.issues.map((i) => i.message) }
}

const CRITERIA = '<filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn>'
/** Hidden rows 1 and 4 (1 + 8) left out, or counted. */
const SKIPPED = [54, 54, 54, 63]
const COUNTED = [63, 54, 54, 63]

describe('xlsx filters: reading the filter mode', () => {
  it("an autofilter with criteria makes the sheet filtered, with Excel's mark or without it", async () => {
    const mark = '<sheetPr filterMode="1"/>'
    expect(await filtered({ before: mark, after: `<autoFilter ref="A1:B6">${CRITERIA}</autoFilter>` })).toEqual({
      filtered: true,
      values: SKIPPED,
      issues: ["an AutoFilter's criteria and sort are not in the model: the sheet is read as filtered, and the rows they hid as hidden"],
    })
    expect((await filtered({ after: `<autoFilter ref="A1:B6">${CRITERIA}</autoFilter>` })).values).toEqual(SKIPPED)
  })

  it('an autofilter without criteria leaves the sheet unfiltered, the mark beside it too; a column with buttons alone has no criteria; a sort is reported', async () => {
    expect(await filtered({ before: '<sheetPr filterMode="1"/>', after: '<autoFilter ref="A1:B6"/>' })).toEqual({
      filtered: false,
      values: COUNTED,
      issues: [],
    })
    expect((await filtered({ after: '<autoFilter ref="A1:B6"><sortState ref="A2:B6"><sortCondition ref="A2:A6"/></sortState></autoFilter>' })).issues).toEqual(["an AutoFilter's sort is not in the model"])
    expect((await filtered({ after: '<autoFilter ref="A1:B6"><filterColumn colId="0" hiddenButton="1"/><filterColumn colId="1" showButton="0"></filterColumn></autoFilter>' })).filtered).toBe(false)
  })

  it('the mark without an autofilter is an advanced filter: the sheet is filtered', async () => {
    const r = await filtered({ before: '<sheetPr filterMode="1"><tabColor rgb="FFFF0000"/></sheetPr>', workbook: '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">S!$A$1:$B$6</definedName></definedNames>' })
    expect(r.filtered).toBe(true)
    expect(r.values).toEqual(SKIPPED)
    expect(r.issues).toContain('an advanced filter is not in the model: the sheet is read as filtered, and the rows it hid as hidden')
  })

  it("a table's filter with criteria makes its sheet filtered; one without criteria does not", async () => {
    expect(await filtered({ table: `<autoFilter ref="A1:B6">${CRITERIA}</autoFilter>` })).toEqual({
      filtered: true,
      values: SKIPPED,
      issues: ["a table's filter criteria are not in the model: the sheet is read as filtered, and the rows they hid as hidden"],
    })
    expect(await filtered({ before: '<sheetPr filterMode="1"/>', table: `<autoFilter ref="A1:B6">${KEEP_ALL}</autoFilter>` })).toMatchObject({ filtered: true, values: SKIPPED })
    expect(await filtered({ table: '<autoFilter ref="A1:B6"/>' })).toEqual({ filtered: false, values: COUNTED, issues: [] })
  })
})

describe('xlsx filters: hidden lines and the filter mode written and read back', () => {
  const sheet = (more: Partial<SheetInput>): BookInput => ({
    sheets: [
      {
        name: 'S',
        cells: { ...Object.fromEntries(Array.from({ length: 6 }, (_, i) => [`B${i + 2}`, { value: 2 ** i }])), D1: { formula: 'SUBTOTAL(9,B2:B7)' }, D2: { formula: 'SUBTOTAL(109,B2:B7)' }, D3: { formula: 'AGGREGATE(9,5,B2:B7)' } },
        ...more,
      },
    ],
  })

  it('hidden rows and columns keep their sizes and looks', async () => {
    const book = sheet({ rows: { 3: { hidden: true }, 5: { height: 30, hidden: true, numberFormat: '0.0' }, 9: { hidden: true } }, columns: { C: { width: 12, hidden: true }, 'E:F': { hidden: true, style: 0 } } })
    book.styles = [{ fill: { color: '#FFFF00' } }]
    const { data, issues } = await writeXlsx({ book }, { values: 'omit' })
    expect(issues).toEqual([])
    const xml = await text(data, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('<row r="5" ht="30" customHeight="1" hidden="1" s="1" customFormat="1">')
    expect(xml).toContain('<row r="9" hidden="1"/>')
    expect(xml).toContain('<col min="3" max="3" width="12" customWidth="1" hidden="1"/><col min="5" max="6" width="8.83203125" hidden="1" style="3"/>')
    // where row 5's format crosses the columns' fill, the cells with both, as Excel writes them
    expect(xml).toContain('<c r="E5" s="2"/><c r="F5" s="2"/>')
    expect(xml).not.toContain('filterMode')
    const r = await readXlsx(data)
    expect(withoutView(r.book)).toEqual(book)
    expect(r.issues).toEqual([])
    expect(compute(r.book).sheets[0].cells.D2.value).toBe(1 + 4 + 16 + 32)
  })

  it('a line of size 0 is written as the book has it and read back hidden, as Excel takes it', async () => {
    const r = await readXlsx((await writeXlsx({ book: sheet({ rows: { 3: { height: 0 } }, columns: { C: { width: 0 } } }) })).data)
    expect(r.book.sheets[0].rows).toEqual({ 3: { height: 0, hidden: true } })
    expect(r.book.sheets[0].columns).toEqual({ C: { width: 0, hidden: true } })
  })

  it("a filtered sheet gets Excel's mark and an autofilter over its cells and hidden rows whose criteria every row meets", async () => {
    const book = sheet({ rows: { 3: { hidden: true }, '11:12': { hidden: true } }, filtered: true })
    const { data, issues } = await writeXlsx({ book }, { values: 'omit' })
    expect(issues).toEqual([])
    const xml = await text(data, 'xl/worksheets/sheet1.xml')
    expect(xml).toMatch(/^<\?xml[^>]*>\n<worksheet [^>]*><sheetPr filterMode="1"\/><dimension /)
    expect(xml).toContain(`</sheetData><autoFilter ref="B1:D12">${KEEP_ALL}</autoFilter></worksheet>`)
    const r = await readXlsx(data)
    // the autofilter over its cells reads back as the sheet's AutoFilter
    expect(withoutView(r.book)).toEqual({ ...book, sheets: [{ ...book.sheets[0], autoFilter: { range: 'B1:D12' } }] })
    expect(codes(r.issues)).toEqual(['autoFilter'])
    const values = compute(r.book).sheets[0].cells
    expect([values.D1.value, values.D2.value, values.D3.value]).toEqual([61, 61, 61])
  })

  it("a filtered sheet's criteria go on its first table with a header row, as Excel keeps a table's filter, without the sheet's mark", async () => {
    const tables = [
      { name: 'Bare', range: 'F1:F3', headers: false, columns: [{ name: 'c' }] },
      { name: 'Data', range: 'B1:B7', columns: [{ name: 'n' }] },
    ]
    const book = sheet({ rows: { 3: { hidden: true } }, filtered: true, tables })
    book.sheets[0].cells.B1 = { value: 'n' }
    const { data, issues } = await writeXlsx({ book }, { values: 'omit' })
    expect(issues).toEqual([])
    expect(await text(data, 'xl/worksheets/sheet1.xml')).not.toMatch(/filterMode|autoFilter/)
    expect(await text(data, 'xl/tables/table2.xml')).toContain(`<autoFilter ref="B1:B7">${KEEP_ALL}</autoFilter>`)
    const r = await readXlsx(data)
    expect(withoutView(r.book)).toEqual(book)
    expect(compute(r.book).sheets[0].cells.D1.value).toBe(61)
  })

  it('a filtered sheet whose tables have no header row and meet its cells loses its filter mode, reported', async () => {
    const book = sheet({ rows: { 3: { hidden: true } }, filtered: true, tables: [{ name: 'Bare', range: 'B2:B7', headers: false, columns: [{ name: 'n' }] }] })
    const { data, issues } = await writeXlsx({ book }, { values: 'omit' })
    expect(codes(issues)).toEqual(['autoFilter'])
    expect(await text(data, 'xl/worksheets/sheet1.xml')).not.toMatch(/filterMode|autoFilter/)
    expect((await readXlsx(data)).book.sheets[0].filtered).toBeUndefined()
  })

  it('a filtered sheet without cells or hidden rows gets an autofilter on A1', async () => {
    const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells: {}, filtered: true }] } })
    expect(await text(data, 'xl/worksheets/sheet1.xml')).toContain(`<autoFilter ref="A1">${KEEP_ALL}</autoFilter>`)
    expect((await readXlsx(data)).book.sheets[0].filtered).toBe(true)
  })
})
