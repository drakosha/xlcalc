/**
 * Cells without a look of their own in a row or column with one. Excel shows a stored cell by its own cell format alone,
 * so the writer stores each cell with the whole of the look the model shows it by (its own style and format over its
 * row's, else its column's), and the reader keeps of a stored cell only what its row or column does not give it.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput } from '../src/contract.js'
import { compute } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { read } from './xlsx-fixtures.js'

const DATE = 'dddd, mmmm d, yyyy'

const book = (): BookInput => ({
  styles: [{ font: { bold: true } }, { fill: { color: '#FFFF00' } }],
  sheets: [
    {
      name: 'S',
      cells: {
        B2: { value: 46096 },
        B3: { value: 7, style: 0 },
        B5: { formula: 'SEQUENCE(2,1,46096)' },
        D4: { value: 1234.5 },
        B4: { value: 5 },
        E4: { value: 'text' },
        B9: { value: 46096, format: 'General' },
      },
      columns: { B: { format: DATE, style: 1, width: 30 } },
      rows: { 4: { format: '#,##0.00' } },
    },
  ],
})

/** The `s` of each stored cell, and the row's and column's cell formats. */
async function stored(data: Uint8Array) {
  const xml = (await openArchive(data).text('xl/worksheets/sheet1.xml'))!
  const cells = Object.fromEntries([...xml.matchAll(/<c r="([A-Z]+\d+)"(?: s="(\d+)")?/g)].map((m) => [m[1], Number(m[2] ?? 0)]))
  const col = Number(/<col [^>]*style="(\d+)"/.exec(xml)?.[1] ?? 0)
  const row = Number(/<row r="4" s="(\d+)" customFormat="1"/.exec(xml)?.[1] ?? 0)
  return { cells, col, row }
}

describe.each(['closures', 'eval'] as const)('cells in formatted rows and columns (%s)', (execution) => {
  it('a cell without a format of its own is stored with its column\'s cell format, or its row\'s', async () => {
    const b = book()
    const { data } = await writeXlsx(b, { result: compute(b, { execution }) })
    const { cells, col, row } = await stored(data)
    expect(col).toBeGreaterThan(0)
    expect(row).toBeGreaterThan(0)
    // the column's cells and the cells its array spilled into; the row's number and text
    expect([cells.B2, cells.B5, cells.B6]).toEqual([col, col, col])
    expect([cells.D4, cells.E4]).toEqual([row, row])
    // the row's format with the column's fill (each from the row, else the column); its own style over the column's
    // format; General of its own with the column's fill
    expect(new Set([cells.B4, cells.B3, cells.B9, col, row, 0]).size).toBe(6)
  })

  it('reads back as written: no cell gets a look its row or column gave it, and every cell shows the same', async () => {
    const b = book()
    const result = compute(b, { execution })
    const back = (await readXlsx((await writeXlsx(b, { result })).data)).book
    const styles = back.styles!
    const cells = back.sheets[0].cells
    expect(cells.B2).toEqual({ value: 46096 })
    expect(cells.D4).toEqual({ value: 1234.5 })
    expect(cells.B4).toEqual({ value: 5 })
    expect(styles[cells.B3.style!]).toEqual({ font: { bold: true } })
    expect(cells.B9).toEqual({ value: 46096, format: 'General' })
    const texts = (r: ReturnType<typeof compute>) => Object.fromEntries(Object.entries(r.sheets[0].cells).map(([a, c]) => [a, c.text]))
    expect(texts(compute(back, { execution }))).toEqual(texts(result))
    expect(texts(result)).toMatchObject({ B2: 'Sunday, March 15, 2026', B9: '46096', D4: '1,234.50', B6: 'Monday, March 16, 2026' })
  })

  it('values only: the cells of a spill keep their column\'s look', async () => {
    const b = book()
    const { data } = await writeXlsx(b, { result: compute(b, { execution }), mode: 'values' })
    const { cells, col } = await stored(data)
    expect([cells.B5, cells.B6]).toEqual([col, col])
    expect((await readXlsx(data)).book.sheets[0].cells.B6).toEqual({ value: 46097 })
  })
})

describe('reading a stored cell over its row and column', () => {
  const STYLES = '<numFmts count="1"><numFmt numFmtId="164" formatCode="d mmmm yyyy"/></numFmts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><cellXfs count="3"><xf/><xf numFmtId="164" fillId="2"/><xf numFmtId="4"/></cellXfs>'
  const COLS = '<cols><col min="2" max="2" width="20" customWidth="1" style="1"/></cols>'

  it('a cell with its column\'s cell format has no look of its own; one stored in General shows General', async () => {
    const r = await read({
      sheets: [{ name: 'S', body: `${COLS}<sheetData><row r="1"><c r="B1" s="1"><v>46096</v></c><c r="C1"><v>1</v></c></row><row r="2"><c r="B2"><v>46096</v></c><c r="B3" s="2"><v>46096</v></c></row><row r="4"><c r="B4"/></row></sheetData>` }],
      styles: STYLES,
    })
    const s = r.book.sheets[0]
    const plain = r.book.styles!.findIndex((x) => Object.keys(x).length === 0)
    expect(plain).toBeGreaterThanOrEqual(0)
    expect(s.cells).toEqual({ B1: { value: 46096 }, C1: { value: 1 }, B2: { value: 46096, style: plain, format: 'General' }, B3: { value: 46096, style: plain, format: 4 }, B4: { style: plain, format: 'General' } })
    const shown = compute(r.book).sheets[0].cells
    expect([shown.B1.text, shown.B2.text, shown.B3.text]).toEqual(['15 March 2026', '46096', '46,096.00'])
  })

  it("a row's format comes before the column's, and the column's fill where the row has none", async () => {
    const r = await read({ sheets: [{ name: 'S', body: `${COLS}<sheetData><row r="1" s="2" customFormat="1"><c r="B1" s="2"><v>5</v></c><c r="C1" s="2"><v>5</v></c><c r="D1" s="1"><v>5</v></c></row></sheetData>` }], styles: STYLES })
    const plain = r.book.styles!.findIndex((x) => Object.keys(x).length === 0)
    expect(r.book.sheets[0].rows).toEqual({ 1: { format: 4 } })
    expect(r.book.sheets[0].cells).toEqual({ B1: { value: 5, style: plain }, C1: { value: 5 }, D1: { value: 5, style: 0, format: 'd mmmm yyyy' } })
  })
})

describe('where a row crosses a column', () => {
  // row 6 bold alone, row 8 a format alone; columns E a fill alone, G a percent alone (as the set lines/files measures)
  const crossing = (): BookInput => ({
    styles: [{ font: { bold: true } }, { fill: { color: '#FFFF00' } }],
    sheets: [{ name: 'S', cells: { A1: { value: 1 } }, columns: { E: { style: 1 }, G: { format: '0.00%' } }, rows: { 6: { style: 0 }, 8: { format: '0.000' } } }],
  })

  it("the writer writes the empty cells Excel would show by the row's whole format, with the model's", async () => {
    const { data } = await writeXlsx(crossing())
    const xml = (await openArchive(data).text('xl/worksheets/sheet1.xml'))!
    expect([...xml.matchAll(/<c r="([A-Z]+\d+)" s="\d+"\/>/g)].map((m) => m[1])).toEqual(['G6', 'E8'])
    const back = await readXlsx(data)
    expect(back.book.sheets[0].cells).toEqual({ A1: { value: 1 } })
    expect(back.book.sheets[0].rows).toEqual({ 6: { style: 0 }, 8: { format: '0.000' } })
  })

  it("a file without those cells reads with the row's look spelled out, as Excel shows them", async () => {
    const STYLES = '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><cellXfs count="5"><xf/><xf fontId="1"/><xf fillId="2"/><xf numFmtId="10"/><xf numFmtId="2"/></cellXfs>'
    const cols = '<cols><col min="5" max="5" width="9" style="2"/><col min="7" max="7" width="9" style="3"/></cols>'
    const r = await read({ sheets: [{ name: 'S', body: `${cols}<sheetData><row r="6" s="1" customFormat="1"><c r="E6" s="1"><v>5</v></c></row><row r="8" s="4" customFormat="1"/><row r="9" s="4" customFormat="1"><c r="E9" s="4"/></row></sheetData>` }], styles: STYLES })
    const plain = r.book.styles!.findIndex((x) => Object.keys(x).length === 0)
    // row 6: column G has a format and no cell in it, the row none: General; E6 holds its row's look
    // row 8: no cell under the filled column E: the plain style; row 9 has its cell there, with the row's look
    expect(r.book.sheets[0].rows).toEqual({ 6: { style: 0, format: 'General' }, 8: { format: 2, style: plain }, 9: { format: 2 } })
    expect(r.book.sheets[0].cells).toEqual({ E6: { value: 5 }, E9: { style: plain } })
    const shown = compute({ ...r.book, sheets: [{ ...r.book.sheets[0], cells: { ...r.book.sheets[0].cells, G6: { value: 46096 } } }] }).sheets[0].cells
    expect(shown.G6.text).toBe('46096')
  })
})
