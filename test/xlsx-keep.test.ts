/**
 * `xlcalc/xlsx/keep`: a book written over the file it was read from keeps what the model does not have (charts,
 * pictures, pivot tables, formatted runs, cell formats the model cannot hold) byte for byte, writes what the book
 * changed, and refuses a change that would move the references of parts the model does not hold.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { createSession, InputError, type BookInput } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { KeptRefusal, openKept, writeKept } from '../src/xlsx-keep.js'
import { changedParts, partsOf, withParts } from './keep-files.js'

const book = (): BookInput => ({
  sheets: [
    {
      name: 'Data',
      cells: {
        A1: { value: 'Item' },
        B1: { value: 'Amount' },
        A2: { value: 'North' },
        B2: { value: 10, style: 0 },
        A3: { value: 'South' },
        B3: { value: 20 },
        A4: { value: 'East' },
        B4: { value: 30 },
        B5: { formula: 'SUM(B2:B4)' },
      },
    },
    { name: 'Other', cells: { A1: { value: 1 }, A2: { formula: 'A1*2' } } },
    { name: 'Plain', cells: { A1: { value: 'p' } } },
  ],
  styles: [{ font: { bold: true } }],
})

/** A file with a chart over Data!B2:B4 and a picture on Other (anchored at H1:M6 and H10:J14), its book opened kept, and a session over it. */
async function charted(extra = {}) {
  const data = await withParts(book(), { chart: { sheet: 1, series: ['Data!$B$2:$B$4'], anchor: [0, 7, 5, 12] }, picture: { sheet: 1, anchor: [9, 7, 13, 9] }, ...extra })
  const read = await openKept(data)
  return { data, read, session: createSession(read.book, { cached: read.cached }) }
}

const KEPT_PARTS = ['xl/charts/chart1.xml', 'xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels', 'xl/media/image1.png', 'xl/worksheets/sheet2.xml', 'xl/worksheets/_rels/sheet2.xml.rels', 'xl/styles.xml', 'xl/sharedStrings.xml']

describe('writeKept: the file as it was', () => {
  it('a book that did not change gives the original bytes, as read and from a session', async () => {
    const { data, read, session } = await charted()
    expect(Buffer.compare(Buffer.from((await writeKept(data, { book: read.book, result: read.cached })).data), Buffer.from(data))).toBe(0)
    expect(Buffer.compare(Buffer.from((await writeKept(data, session.getSnapshot())).data), Buffer.from(data))).toBe(0)
    // a file Excel saved, with a calculation chain and a theme
    const saved = readFileSync(new URL('./excel/xlsx/saved/areas-1.xlsx', import.meta.url))
    const r = await openKept(saved)
    expect(Buffer.compare(Buffer.from((await writeKept(saved, createSession(r.book, { cached: r.cached }).getSnapshot())).data), saved)).toBe(0)
  })

  it('an edit of a value writes its sheet; the chart, the drawing, the picture and the other sheets stay byte for byte', async () => {
    const { data, session } = await charted()
    session.setCells([{ sheet: 'Data', cell: 'B3', input: { value: 25 } }])
    const w = await writeKept(data, session.getSnapshot())
    expect(w.issues).toEqual([])
    // the workbook asks Excel to compute it on loading already, as writeXlsx wrote it
    expect(await changedParts(data, w.data)).toEqual(['xl/worksheets/sheet1.xml'])
    const [before, after] = [await partsOf(data), await partsOf(w.data)]
    for (const p of KEPT_PARTS) expect(after.get(p), p).toBe(before.get(p))
    const back = await readXlsx(w.data)
    expect(back.book.sheets[0].cells.B3).toEqual({ value: 25 })
    expect(back.cached.sheets[0].cells.B5).toEqual({ value: 65 })
    // Excel computes the book again on loading
    expect(after.get('xl/workbook.xml')).toContain('fullCalcOnLoad="1"')
  })

  it('a cell that did not change is the file\'s own: its formatted runs and a cell format the model cannot hold stay; a new value goes inline', async () => {
    const b = book()
    // B1 and B2 share the bold cell format, which gets a rotation; A2 gets formatted runs
    const data = await withParts(b, { rich: { North: '<r><rPr><b/></rPr><t>No</t></r><r><t>rth</t></r>' }, rotated: [1] })
    const read = await openKept(data)
    const s = createSession(read.book, { cached: read.cached })
    s.setCells([{ sheet: 'Data', cell: 'B2', input: { value: 11, style: read.book.sheets[0].cells.B2.style } }])
    s.setCells([{ sheet: 'Data', cell: 'A3', input: { value: 'Southwest' } }])
    const sheet = (await partsOf((await writeKept(data, s.getSnapshot())).data)).get('xl/worksheets/sheet1.xml')!
    const was = (await partsOf(data)).get('xl/worksheets/sheet1.xml')!
    const cell = (xml: string, a: string) => new RegExp(`<c r="${a}"[^>]*?(/>|>.*?</c>)`).exec(xml)![0]
    // A2 points at its shared string with runs as it did; B2 keeps its cell format (rotated) with its new value
    expect(cell(sheet, 'A2')).toBe(cell(was, 'A2'))
    expect(cell(sheet, 'B2')).toBe(cell(was, 'B2').replace('<v>10</v>', '<v>11</v>'))
    expect(cell(sheet, 'A3')).toBe('<c r="A3" t="inlineStr"><is><t>Southwest</t></is></c>')
    // the shared strings stay as the file has them
    expect((await partsOf((await writeKept(data, s.getSnapshot())).data)).get('xl/sharedStrings.xml')).toBe((await partsOf(data)).get('xl/sharedStrings.xml'))
  })

  it('a new look is added at the end of the cell formats; one the file has takes its index', async () => {
    const { data, session } = await charted()
    session.patchCells([{ sheet: 'Data', cell: 'A2', set: { style: { font: { bold: true } } } }])
    let w = await writeKept(data, session.getSnapshot())
    // bold is B2's cell format: styles.xml stays
    expect(await changedParts(data, w.data)).toEqual(['xl/worksheets/sheet1.xml'])
    session.patchCells([{ sheet: 'Data', cell: 'A3', set: { style: { fill: { color: '#FFEE00' } } } }])
    w = await writeKept(data, session.getSnapshot())
    const [before, after] = [(await partsOf(data)).get('xl/styles.xml')!, (await partsOf(w.data)).get('xl/styles.xml')!]
    const count = (xml: string) => Number(/<cellXfs count="(\d+)"/.exec(xml)![1])
    expect(count(after)).toBe(count(before) + 1)
    expect(after.startsWith(before.slice(0, before.indexOf('<fills')))).toBe(true)
    const back = await readXlsx(w.data)
    expect(back.book.styles![back.book.sheets[0].cells.A3.style!]).toEqual({ fill: { color: '#FFEE00' } })
  })

  it('a shared formula whose master changes is written in each of its copies', async () => {
    const data = await withParts(book(), {})
    const parts = await partsOf(data)
    // the formulas of C2:C4 as Excel shares them, C2 the master
    const sheet = parts.get('xl/worksheets/sheet1.xml')!.replace('<c r="B2" s="1"><v>10</v></c>', '<c r="B2" s="1"><v>10</v></c><c r="C2"><f t="shared" ref="C2:C4" si="0">B2*2</f><v>20</v></c>').replace('<c r="B3"><v>20</v></c>', '<c r="B3"><v>20</v></c><c r="C3"><f t="shared" si="0"/><v>40</v></c>').replace('<c r="B4"><v>30</v></c>', '<c r="B4"><v>30</v></c><c r="C4"><f t="shared" si="0"/><v>60</v></c>')
    const { zip } = await import('../src/xlsx/zip/write.js')
    const file = await zip([...parts].map(([name, text]) => ({ name, data: new TextEncoder().encode(name === 'xl/worksheets/sheet1.xml' ? sheet : text) })))
    const read = await openKept(file)
    const s = createSession(read.book, { cached: read.cached })
    s.setCells([{ sheet: 'Data', cell: 'C2', input: { formula: 'B2*3' } }])
    const w = await writeKept(file, s.getSnapshot())
    const out = (await partsOf(w.data)).get('xl/worksheets/sheet1.xml')!
    expect(out).not.toContain('t="shared"')
    const back = await readXlsx(w.data)
    expect([back.book.sheets[0].cells.C2.formula, back.book.sheets[0].cells.C3.formula, back.book.sheets[0].cells.C4.formula]).toEqual(['B2*3', 'B3*2', 'B4*2'])
    expect(back.cached.sheets[0].cells.C2.value).toBe(30)
  })

  it('the calculation chain goes, and a dynamic array formula gets the metadata where the file has none', async () => {
    const saved = readFileSync(new URL('./excel/xlsx/saved/areas-1.xlsx', import.meta.url))
    const parts = await partsOf(saved)
    expect(parts.has('xl/calcChain.xml')).toBe(true)
    const r = await openKept(saved)
    const s = createSession(r.book, { cached: r.cached })
    s.setCells([{ sheet: r.book.sheets[0].name, cell: 'Z1', input: { formula: 'SEQUENCE(3)' } }])
    const w = await writeKept(saved, s.getSnapshot())
    const after = await partsOf(w.data)
    expect(after.has('xl/calcChain.xml')).toBe(false)
    expect(after.get('[Content_Types].xml')).not.toContain('calcChain')
    expect(after.get('xl/_rels/workbook.xml.rels')).not.toContain('calcChain')
    const back = await readXlsx(w.data)
    expect(back.cached.sheets[0].cells.Z3).toEqual({ value: 3, spillOf: 'Z1' })
  })
})

describe('writeKept: what the model has, written', () => {
  it('a note, a hyperlink, a merged area, a rule, a validation, a name and a new sheet', async () => {
    const { data, session } = await charted()
    session.patchCells([{ sheet: 'Data', cell: 'A2', set: { note: { author: 'Ann', text: 'checked' }, link: { url: 'https://example.com/north' } } }])
    session.editStructure({ op: 'mergeCells', sheet: 'Other', range: 'C1:D1' })
    session.setLayout([
      { sheet: 'Data', conditionalFormats: [{ type: 'cellIs', ranges: ['B2:B4'], operator: 'greaterThan', formulas: ['15'], style: { font: { bold: true } } }] },
      { sheet: 'Other', validations: [{ range: 'A1', type: 'whole', operator: 'greaterThan', formula1: '0' }] },
    ])
    // the carriers stay among the names
    session.setNames([...session.getBook().names!, { name: 'total', ref: "'Data'!$B$5" }])
    session.editStructure({ op: 'insertSheet', name: 'Notes', before: 'Other' })
    session.setCells([{ sheet: 'Notes', cell: 'A1', input: { formula: 'total*2' } }])
    const w = await writeKept(data, session.getSnapshot())
    const back = await readXlsx(w.data)
    const [d, notes, other] = back.book.sheets
    expect(back.book.sheets.map((x) => x.name)).toEqual(['Data', 'Notes', 'Other', 'Plain'])
    expect(d.cells.A2).toEqual({ value: 'North', note: { author: 'Ann', text: 'checked' }, link: { url: 'https://example.com/north' } })
    expect(d.conditionalFormats).toEqual([{ type: 'cellIs', ranges: ['B2:B4'], operator: 'greaterThan', formulas: ['15'], style: { font: { bold: true } } }])
    expect(other.merges).toEqual(['C1:D1'])
    expect(other.validations).toEqual([{ range: 'A1', type: 'whole', operator: 'greaterThan', formula1: '0' }])
    expect(back.book.names).toEqual([{ name: 'total', ref: "'Data'!$B$5" }])
    expect(back.cached.sheets[1].cells.A1).toEqual({ value: 120 })
    expect(notes.cells.A1).toEqual({ formula: 'total*2' })
    // the chart and the picture stay, and the carriers are no names of the file
    const [before, after] = [await partsOf(data), await partsOf(w.data)]
    for (const p of ['xl/charts/chart1.xml', 'xl/drawings/drawing1.xml', 'xl/media/image1.png']) expect(after.get(p), p).toBe(before.get(p))
    expect(after.get('xl/workbook.xml')).not.toContain('_xlcalc.keep')
  })

  it('rows inserted on a sheet no part refers to are written, the chart stays', async () => {
    const { data, session } = await charted()
    expect(session.editStructure({ op: 'insertRows', sheet: 'Plain', row: 1 }).ok).toBe(true)
    const w = await writeKept(data, session.getSnapshot())
    expect(await changedParts(data, w.data)).toEqual(['xl/worksheets/sheet3.xml'])
    expect((await readXlsx(w.data)).book.sheets[2].cells).toEqual({ A2: { value: 'p' } })
  })

  it('rows inserted below the chart\'s series on their sheet are written; rows inserted where nothing moves change nothing', async () => {
    const { data, session } = await charted()
    expect(session.editStructure({ op: 'insertRows', sheet: 'Other', row: 20 }).ok).toBe(true)
    expect(await changedParts(data, (await writeKept(data, session.getSnapshot())).data)).toEqual([])
    expect(session.editStructure({ op: 'insertRows', sheet: 'Data', row: 5 }).ok).toBe(true)
    const w = await writeKept(data, session.getSnapshot())
    expect(await changedParts(data, w.data)).toEqual(['xl/worksheets/sheet1.xml'])
    expect((await readXlsx(w.data)).book.sheets[0].cells.B6).toEqual({ formula: 'SUM(B2:B4)' })
  })
})

describe('writeKept refuses what would point the file at other cells', () => {
  const refused = async (edit: (s: ReturnType<typeof createSession>) => void, extra = {}) => {
    const { data, session } = await charted(extra)
    edit(session)
    return writeKept(data, session.getSnapshot()).then(
      () => 'written',
      (e: Error) => (e instanceof KeptRefusal ? e.message : `${e.name}: ${e.message}`),
    )
  }

  it('rows inserted above the chart\'s series, a move of them, a rename of their sheet', async () => {
    expect(await refused((s) => s.editStructure({ op: 'insertRows', sheet: 'Data', row: 2 }))).toMatch(/^xlsx\/keep: the chart xl\/charts\/chart1\.xml refers to 'Data'!\$B\$2:\$B\$4, which the change made 'Data'!\$B\$3:\$B\$5/)
    expect(await refused((s) => s.editStructure({ op: 'moveRange', from: { sheet: 'Data', range: 'B2:B4' }, to: { sheet: 'Other', cell: 'D1' } }))).toMatch(/the chart xl\/charts\/chart1\.xml refers to 'Data'!\$B\$2:\$B\$4, which the change made 'Other'!\$D\$1:\$D\$3/)
    expect(await refused((s) => s.editStructure({ op: 'renameSheet', sheet: 'Data', name: 'Sales' }))).toMatch(/the chart .* which the change made 'Sales'!\$B\$2:\$B\$4/)
  })

  it('a column inserted where the drawing lies, a sheet with drawings deleted', async () => {
    expect(await refused((s) => s.editStructure({ op: 'insertColumns', sheet: 'Other', column: 'C' }))).toBe("xlsx/keep: a drawing on 'Other' (xl/drawings/drawing1.xml) refers to 'Other'!$H$1:$M$6, which the change made 'Other'!$I$1:$N$6; writeKept does not move the references of parts the workbook does not hold, so the file would point at other cells")
    expect(await refused((s) => s.editStructure({ op: 'deleteSheet', sheet: 'Other' }))).toBe("xlsx/keep: the sheet 'Other' was deleted, and it has a part the workbook does not hold (drawing xl/drawings/drawing1.xml)")
  })

  it('the print area and page breaks the change would move; a pivot table\'s source table renamed', async () => {
    expect(await refused((s) => s.editStructure({ op: 'deleteRows', sheet: 'Plain', row: 8 }), { printArea: { sheet: 2, ref: "'Plain'!$A$1:$B$9" } })).toMatch(/the name _xlnm\.Print_Area refers to 'Plain'!\$A\$1:\$B\$9, which the change made 'Plain'!\$A\$1:\$B\$8/)
    expect(await refused((s) => s.editStructure({ op: 'insertRows', sheet: 'Plain', row: 3 }), { breaks: { sheet: 2, rows: [4] } })).toMatch(/the page breaks of 'Plain'/)
    const table: BookInput = { sheets: [{ name: 'Data', cells: { A1: { value: 'k' }, B1: { value: 'v' }, A2: { value: 'x' }, B2: { value: 1 } }, tables: [{ name: 'Sales', range: 'A1:B2', columns: [{ name: 'k' }, { name: 'v' }] }] }, { name: 'Other', cells: {} }] }
    const data = await withParts(table, { pivot: { sheet: 1, location: 'A3:B5', source: { name: 'Sales' } } })
    const r = await openKept(data)
    const s = createSession(r.book, { cached: r.cached })
    s.editStructure({ op: 'renameTable', table: 'Sales', name: 'Orders' })
    await expect(writeKept(data, s.getSnapshot())).rejects.toThrow("xlsx/keep: the source of the pivot cache xl/pivotCache/pivotCacheDefinition1.xml is the table or name Sales, which the workbook no longer has")
  })

  it('a book without the carriers of a file with charts: open it with openKept', async () => {
    const { data } = await charted()
    const r = await readXlsx(data)
    await expect(writeKept(data, { book: r.book, result: r.cached })).rejects.toThrow("xlsx/keep: the workbook has no name _xlcalc.keep.0, which holds a reference of a drawing on 'Other' (xl/drawings/drawing1.xml): open the file with openKept and keep the hidden names it adds")
  })
})

describe('openKept and writeKept take what readXlsx and writeXlsx take', () => {
  it('openKept: the book read, with hidden names that carry the references of the parts the model does not have', async () => {
    const { data, read } = await charted()
    const plain = await readXlsx(data)
    expect({ ...read.book, names: undefined }).toEqual({ ...plain.book, names: undefined })
    expect(read.book.names).toEqual([
      { name: '_xlcalc.keep.s0', ref: "'Data'!$XFD$1048576", hidden: true },
      { name: '_xlcalc.keep.s1', ref: "'Other'!$XFD$1048576", hidden: true },
      { name: '_xlcalc.keep.s2', ref: "'Plain'!$XFD$1048576", hidden: true },
      { name: '_xlcalc.keep.0', ref: "'Other'!$H$1:$M$6", hidden: true },
      { name: '_xlcalc.keep.1', ref: "'Other'!$H$10:$J$14", hidden: true },
      { name: '_xlcalc.keep.2', ref: "'Data'!$B$2:$B$4", hidden: true },
    ])
    expect(read.issues).toEqual(plain.issues)
  })

  it('their input is checked as readXlsx and writeXlsx check theirs', async () => {
    const { data, session } = await charted()
    await expect(openKept('book.xlsx' as never)).rejects.toThrow(InputError)
    await expect(writeKept(data, session as never)).rejects.toThrow(/a session is not a snapshot/)
    await expect(writeKept(data, session.getSnapshot(), { mode: 'values' } as never)).rejects.toThrow(InputError)
    await expect((writeKept as (...a: unknown[]) => Promise<unknown>)(data, session.getSnapshot(), { values: 'omit' }, {})).rejects.toThrow(/writeKept/)
    const w = await writeKept(new Uint8Array(data).buffer, { book: session.getBook() }, { values: 'omit', limits: { entryBytes: 2 ** 24 } })
    expect(w.data).toBeInstanceOf(Uint8Array)
  })

  it('writeXlsx of the same book reads back the same, the file\'s parts aside', async () => {
    const { data, session } = await charted()
    session.setCells([{ sheet: 'Data', cell: 'B4', input: { value: 31 } }])
    const snap = session.getSnapshot()
    const plain = { ...snap, book: { ...snap.book, names: snap.book.names?.filter((n) => !n.name.startsWith('_xlcalc.keep.')) } }
    if (!plain.book.names?.length) delete plain.book.names
    const [a, b] = [await readXlsx((await writeKept(data, snap)).data), await readXlsx((await writeXlsx(plain)).data)]
    expect(a.book).toEqual(b.book)
    expect(a.cached).toEqual(b.cached)
  })
})
