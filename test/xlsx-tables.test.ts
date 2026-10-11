/**
 * Tables through xlsx: a book's tables written as table parts and read back as they were (names, ranges, header and
 * totals rows, calculated columns, totals), structured references in the file's spelling, and what a file can hold
 * that the model cannot, reported. Excel's own files are checked in `test/xlsx-excel.test.ts`.
 */
import { describe, expect, it } from 'vitest'
import { compute, type BookInput } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { pack, parts as fixture } from './xlsx-fixtures.js'

const text = async (data: Uint8Array, name: string) => (await openArchive(data).text(name))!

const book: BookInput = {
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 'Region' }, B1: { value: 'Qty' }, C1: { value: '@at' }, D1: { value: 'Total' },
        A2: { value: 'North' }, B2: { value: 2 }, C2: { value: 10 }, D2: { formula: '[@Qty]*[@[@at]]' },
        A3: { value: 'South' }, B3: { value: 3 }, C3: { value: 20 }, D3: { formula: "[@Qty]*[@['@at]]" },
        A4: { value: 'Total' }, B4: { formula: 'SUBTOTAL(109,[Qty])' }, D4: { formula: 'SUM([Total])' },
        F2: { value: 1 }, F3: { value: 2 },
        H1: { formula: "SUM(Sales[Qty])+SUM(Sales['@at])+ROWS(Bare)" },
      },
      tables: [
        {
          name: 'Sales',
          range: 'A1:D4',
          totals: true,
          columns: [{ name: 'Region', totals: { label: 'Total' } }, { name: 'Qty', totals: { function: 'sum' } }, { name: '@at' }, { name: 'Total', formula: "Sales[[#This Row],[Qty]]*Sales[[#This Row],['@at]]" }],
        },
        { name: 'Bare', range: 'F2:F3', headers: false, columns: [{ name: 'n' }] },
      ],
    },
  ],
  names: [{ name: 'qty', ref: 'Sales[Qty]' }],
}

describe('tables in xlsx', () => {
  it('go into table parts and come back as they were', async () => {
    const result = compute(book)
    const { data, issues } = await writeXlsx(book, { result })
    expect(issues).toEqual([])
    const t1 = await text(data, 'xl/tables/table1.xml')
    expect(t1).toContain('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:D4" totalsRowCount="1">')
    expect(t1).toContain('<tableColumn id="3" name="@at"/>')
    expect(t1).toContain('<calculatedColumnFormula>Sales[[#This Row],[Qty]]*Sales[[#This Row],[@at]]</calculatedColumnFormula>')
    expect(t1).toContain('<tableColumn id="4" name="Total" totalsRowFunction="custom"><calculatedColumnFormula>')
    expect(await text(data, 'xl/tables/table2.xml')).toContain('headerRowCount="0"')
    const s1 = await text(data, 'xl/worksheets/sheet1.xml')
    expect(s1).toContain('<f>SUBTOTAL(109,Sales[Qty])</f>')
    expect(s1).toContain('<f t="array" ref="D3">Sales[[#This Row],[Qty]]*Sales[[#This Row],[@at]]</f>')
    expect(await text(data, 'xl/workbook.xml')).toContain('<definedName name="qty">Sales[Qty]</definedName>')
    const read = await readXlsx(data)
    expect(read.issues).toEqual([])
    const tables = read.book.sheets[0].tables!
    expect(tables[0]).toEqual({ ...book.sheets[0].tables![0], columns: book.sheets[0].tables![0].columns.map((c) => (c.name === 'Total' ? { ...c, totals: { function: 'custom' } } : c)) })
    expect(tables[1]).toEqual(book.sheets[0].tables![1])
    expect(compute(read.book).hash).toBe(result.hash)
    expect(read.book.sheets[0].cells.D2.formula).toBe("Sales[[#This Row],[Qty]]*Sales[[#This Row],['@at]]")
  })

  it('a header cell without its name is given it, a table the model cannot hold is left out, both reported', async () => {
    const sheet =
      '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="B1"><v>5</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="D1:E1"/></mergeCells><tableParts count="3"><tablePart r:id="rId1"/><tablePart r:id="rId2"/><tablePart r:id="rId3"/></tableParts>'
    const table = (id: number, name: string, ref: string, cols: string[]) =>
      `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${name}" displayName="${name}" ref="${ref}" headerRowCount="${ref.startsWith('D') ? 0 : 1}"><tableColumns count="${cols.length}">${cols.map((c, i) => `<tableColumn id="${i + 1}" name="${c}"/>`).join('')}</tableColumns></table>`
    const rels = [1, 2, 3].map((i) => `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${i}.xml"/>`).join('')
    const data = await pack(
      fixture({
        sheets: [{ name: 'S', body: sheet }],
        parts: {
          'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
          'xl/tables/table1.xml': table(1, 'T', 'A1:B2', ['a', 'b']),
          'xl/tables/table2.xml': table(2, 'Over', 'B1:C2', ['x', 'y']),
          'xl/tables/table3.xml': table(3, 'Merged', 'D1:E2', ['p', 'q']),
        },
      }),
    )
    const r = await readXlsx(data)
    expect(r.book.sheets[0].tables).toEqual([{ name: 'T', range: 'A1:B2', columns: [{ name: 'a' }, { name: 'b' }] }])
    expect(r.book.sheets[0].cells.B1).toEqual({ value: 'b' })
    expect(r.issues.map((i) => [i.code, i.message.replace(/:.*/, '')])).toEqual([
      ['table', "a header cell that does not hold its column's name as text is read with the name \"b\""],
      ['table', 'the table Over is left out'],
      ['table', 'the table Merged is left out'],
    ])
  })

  it("a calculated column formula is reported as a cell's formula is, at the column's first cell; one that differs from its cells' is kept", async () => {
    const f = (r: number) => `<c r="B${r}"><f>T[[#This Row],[a]]*3</f></c>`
    const sheet = `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="B1" t="inlineStr"><is><t>b</t></is></c><c r="C1" t="inlineStr"><is><t>c</t></is></c><c r="D1" t="inlineStr"><is><t>d</t></is></c></row><row r="2"><c r="A2"><v>1</v></c>${f(2)}</row><row r="3"><c r="A3"><v>2</v></c>${f(3)}</row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts>`
    const calc = (name: string, formula: string) => `<tableColumn id="1" name="${name}"><calculatedColumnFormula>${formula}</calculatedColumnFormula></tableColumn>`
    const table = `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:D3"><tableColumns count="4"><tableColumn id="1" name="a"/>${calc('b', 'T[[#This Row],[a]]*2')}${calc('c', '_xludf.MYFN(T[[#This Row],[a]])')}${calc('d', 'T[[#This Row],[a]]+')}</tableColumns></table>`
    const data = await pack(
      fixture({
        sheets: [{ name: 'S', body: sheet }],
        parts: {
          'xl/worksheets/_rels/sheet1.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
          'xl/tables/table1.xml': table,
        },
      }),
    )
    const r = await readXlsx(data)
    expect(r.book.sheets[0].tables![0].columns.map((c) => c.formula)).toEqual([undefined, 'T[[#This Row],[a]]*2', 'MYFN(T[[#This Row],[a]])', 'T[[#This Row],[a]]+'])
    expect(r.book.sheets[0].cells.B2.formula).toBe('T[[#This Row],[a]]*3')
    expect(r.issues.map((i) => [i.code, i.cells])).toEqual([
      ['userFunction', ['C1']],
      ['formula', ['D1']],
    ])
  })
})
