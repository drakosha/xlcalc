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
import { KEPT_CELLS } from '../src/xlsx/keep/carriers.js'
import { changedParts, partsOf, withParts } from './keep-files.js'
import { ruRU } from './locales.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { zip } from '../src/xlsx/zip/write.js'

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
    const data = await withParts(b, { rich: { North: '<r><rPr><b/></rPr><t>No</t></r><r><t>rth</t></r>' }, shrunk: [1] })
    const read = await openKept(data)
    const s = createSession(read.book, { cached: read.cached })
    s.setCells([{ sheet: 'Data', cell: 'B2', input: { value: 11, style: read.book.sheets[0].cells.B2.style } }])
    s.setCells([{ sheet: 'Data', cell: 'A3', input: { value: 'Southwest' } }])
    const sheet = (await partsOf((await writeKept(data, s.getSnapshot())).data)).get('xl/worksheets/sheet1.xml')!
    const was = (await partsOf(data)).get('xl/worksheets/sheet1.xml')!
    const cell = (xml: string, a: string) => new RegExp(`<c r="${a}"[^>]*?(/>|>.*?</c>)`).exec(xml)![0]
    // A2 points at its shared string with runs as it did; B2 keeps its cell format (shrunk to fit) with its new value
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

describe('writeKept moves the cells that hold what only the file holds', () => {
  const RUNS = '<r><rPr><b/></rPr><t>No</t></r><r><t>rth</t></r>'
  /** A file whose A2 (North) holds formatted runs and B3 a rich value, and the Data sheet's cells after `edit`: each cell's XML by address, as the file has it and as written. */
  const after = async (edit: (s: ReturnType<typeof createSession>) => void) => {
    const data = await withParts(book(), { rich: { North: RUNS }, attrs: { sheet: 0, cells: { B3: 'vm="1"' } } })
    const r = await openKept(data)
    const s = createSession(r.book, { cached: r.cached })
    edit(s)
    const w = await writeKept(data, s.getSnapshot())
    const cells = (xml: string) => Object.fromEntries([...xml.matchAll(/<c r="(\w+)"[^>]*?(?:\/>|>.*?<\/c>)/g)].map((m) => [m[1], m[0]]))
    const [was, now] = [await partsOf(data), await partsOf(w.data)]
    return { was: cells(was.get('xl/worksheets/sheet1.xml')!), now: cells(now.get('xl/worksheets/sheet1.xml')!), strings: now.get('xl/sharedStrings.xml') === was.get('xl/sharedStrings.xml'), issues: w.issues }
  }
  const at = (xml: string, addr: string) => xml.replace(/^<c r="\w+"/, `<c r="${addr}"`)

  it('openKept gives each a carrier of its own', async () => {
    const data = await withParts(book(), { rich: { North: RUNS }, attrs: { sheet: 0, cells: { B3: 'vm="1"', B4: 'ph="1"' } } })
    expect((await openKept(data)).book.names!.filter((n) => /^_xlcalc\.keep\.c/.test(n.name))).toEqual([
      { name: '_xlcalc.keep.c0.0', ref: "'Data'!$A$2", hidden: true },
      { name: '_xlcalc.keep.c0.1', ref: "'Data'!$B$3", hidden: true },
      { name: '_xlcalc.keep.c0.2', ref: "'Data'!$B$4", hidden: true },
    ])
  })

  it('rows and columns inserted, a range moved: the file\'s cell at its new address; its old address takes what came there', async () => {
    let r = await after((s) => s.editStructure({ op: 'insertRows', sheet: 'Data', row: 1, count: 2 }))
    expect([r.now.A4, r.now.B5]).toEqual([at(r.was.A2, 'A4'), at(r.was.B3, 'B5')])
    expect(r.strings).toBe(true)
    r = await after((s) => s.editStructure({ op: 'insertColumns', sheet: 'Data', column: 'A' }))
    expect([r.now.B2, r.now.C3]).toEqual([at(r.was.A2, 'B2'), at(r.was.B3, 'C3')])
    r = await after((s) => s.editStructure({ op: 'moveRange', from: { sheet: 'Data', range: 'A2' }, to: { sheet: 'Data', cell: 'D10' } }))
    expect([r.now.D10, r.now.A2]).toEqual([at(r.was.A2, 'D10'), undefined])
    // the row of North deleted: South moves up into A2, written from the workbook
    r = await after((s) => s.editStructure({ op: 'deleteRows', sheet: 'Data', row: 2 }))
    expect([r.now.A2, r.now.B2]).toEqual(['<c r="A2" t="inlineStr"><is><t>South</t></is></c>', at(r.was.B3, 'B2')])
  })

  it('a cell changed after it moved is written from the workbook; one that stayed is the file\'s', async () => {
    let r = await after((s) => {
      s.editStructure({ op: 'insertRows', sheet: 'Data', row: 1 })
      s.setCells([{ sheet: 'Data', cell: 'A3', input: { value: 'Northeast' } }])
    })
    expect(r.now.A3).toBe('<c r="A3" t="inlineStr"><is><t>Northeast</t></is></c>')
    r = await after((s) => s.editStructure({ op: 'insertRows', sheet: 'Data', row: 10 }))
    expect([r.now.A2, r.now.B3]).toEqual([r.was.A2, r.was.B3])
  })

  it(`past ${KEPT_CELLS} of them on a sheet, none has a carrier: rows or columns moved there are reported`, async () => {
    const cells = Object.fromEntries(Array.from({ length: KEPT_CELLS + 1 }, (_, i) => [`A${i + 1}`, { value: 'North' }]))
    const data = await withParts({ sheets: [{ name: 'Big', cells }] }, { rich: { North: RUNS } })
    const r = await openKept(data)
    expect(r.book.names!.filter((n) => /^_xlcalc\.keep\.c/.test(n.name))).toEqual([])
    const s = createSession(r.book, { cached: r.cached })
    s.editStructure({ op: 'insertRows', sheet: 'Big', row: 1 })
    expect((await writeKept(data, s.getSnapshot())).issues).toEqual([
      { code: 'richText', sheet: 'Big', message: `rows or columns moved on a sheet with ${KEPT_CELLS + 1} cells holding formatted runs of text or rich values that the workbook does not have, more than writeKept follows: those that moved are written from the workbook, without them` },
    ])
  })
})

/** After `edit` over the charted file: the written file's parts as text, or the refusal's message. */
async function moved(edit: (s: ReturnType<typeof createSession>) => void, extra = {}): Promise<Map<string, string> | string> {
  const { data, session } = await charted(extra)
  edit(session)
  return writeKept(data, session.getSnapshot()).then(
    async (w) => partsOf(w.data),
    (e: Error) => (e instanceof KeptRefusal ? `${e.code}: ${e.message}` : `${e.name}: ${e.message}`),
  )
}

/** The chart's series formulas, and each anchor of the drawing: its corners' column, offset, row and offset. */
const series = (parts: Map<string, string> | string) => (typeof parts === 'string' ? parts : [...parts.get('xl/charts/chart1.xml')!.matchAll(/<c:f>([^<]*)<\/c:f>/g)].map((m) => m[1]))
const anchors = (parts: Map<string, string> | string) =>
  typeof parts === 'string' ? parts : [...parts.get('xl/drawings/drawing1.xml')!.matchAll(/<xdr:(from|to)><xdr:col>(\d+)<\/xdr:col><xdr:colOff>(\d+)<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row><xdr:rowOff>(\d+)<\/xdr:rowOff>/g)].map((m) => `${m[1]} ${m[2]}+${m[3]} ${m[4]}+${m[5]}`)

describe('writeKept moves the references of the parts the model does not have', () => {
  it("a chart's series: rows inserted above them and in them, deleted in them, all of them deleted, a move, a rename, their sheet deleted", async () => {
    expect(series(await moved((s) => s.editStructure({ op: 'insertRows', sheet: 'Data', row: 2 })))).toEqual(['Data!$B$3:$B$5'])
    expect(series(await moved((s) => s.editStructure({ op: 'insertRows', sheet: 'Data', row: 3, count: 2 })))).toEqual(['Data!$B$2:$B$6'])
    expect(series(await moved((s) => s.editStructure({ op: 'deleteRows', sheet: 'Data', row: 3 })))).toEqual(['Data!$B$2:$B$3'])
    expect(series(await moved((s) => s.editStructure({ op: 'deleteRows', sheet: 'Data', row: 2, count: 3 })))).toEqual(['Data!#REF!'])
    expect(series(await moved((s) => s.editStructure({ op: 'moveRange', from: { sheet: 'Data', range: 'B2:B4' }, to: { sheet: 'Plain', cell: 'D1' } })))).toEqual(['Plain!$D$1:$D$3'])
    expect(series(await moved((s) => s.editStructure({ op: 'renameSheet', sheet: 'Data', name: 'Sales 2' })))).toEqual(["'Sales 2'!$B$2:$B$4"])
    expect(series(await moved((s) => s.editStructure({ op: 'deleteSheet', sheet: 'Data' })))).toEqual(['#REF!'])
  })

  it('drawings by editAs: the chart moves and sizes with its cells, the picture only moves, a deleted line takes its corner to where it was', async () => {
    // the chart H1:M6 (twoCell), the picture H10:J14 (oneCell), each corner at offsets 1000 and 2000
    const was = ['from 7+1000 0+2000', 'to 12+1000 5+2000', 'from 7+1000 9+2000', 'to 9+1000 13+2000']
    expect(anchors(await moved((s) => s.editStructure({ op: 'insertRows', sheet: 'Plain', row: 1 })))).toEqual(was)
    expect(anchors(await moved((s) => s.editStructure({ op: 'insertColumns', sheet: 'Other', column: 'C' })))).toEqual(['from 8+1000 0+2000', 'to 13+1000 5+2000', 'from 8+1000 9+2000', 'to 10+1000 13+2000'])
    // rows inserted inside both: the chart grows, the picture keeps its size
    expect(anchors(await moved((s) => s.editStructure({ op: 'insertRows', sheet: 'Other', row: 3, count: 2 })))).toEqual(['from 7+1000 0+2000', 'to 12+1000 7+2000', 'from 7+1000 11+2000', 'to 9+1000 15+2000'])
    expect(anchors(await moved((s) => s.editStructure({ op: 'insertRows', sheet: 'Other', row: 12 })))).toEqual([...was.slice(0, 3), 'to 9+1000 13+2000'])
    // the chart's last row deleted: its bottom where that row was; the picture's rows all deleted: it moves up to where
    // they were, its height kept (its top's offset gone, its bottom's goes too: four rows of the default height)
    expect(anchors(await moved((s) => s.editStructure({ op: 'deleteRows', sheet: 'Other', row: 6 })))).toEqual(['from 7+1000 0+2000', 'to 12+1000 5+0', 'from 7+1000 8+2000', 'to 9+1000 12+2000'])
    expect(anchors(await moved((s) => s.editStructure({ op: 'deleteRows', sheet: 'Other', row: 8, count: 10 })))).toEqual([...was.slice(0, 2), 'from 7+1000 7+0', 'to 9+1000 11+0'])
    // all the chart's rows deleted: Excel deletes it, its chart part with it
    const gone = (await moved((s) => s.editStructure({ op: 'deleteRows', sheet: 'Other', row: 1, count: 6 }))) as Map<string, string>
    expect(anchors(gone)).toEqual(['from 7+1000 3+2000', 'to 9+1000 7+2000'])
    expect([...gone.keys()].filter((p) => /chart/.test(p))).toEqual([])
    expect(gone.get('xl/drawings/_rels/drawing1.xml.rels')).not.toContain('chart')
    expect(gone.get('[Content_Types].xml')).not.toContain('chart1')
    // a fixed drawing stays where it is
    expect(anchors(await moved((s) => s.editStructure({ op: 'insertRows', sheet: 'Other', row: 1 }), { picture: { sheet: 1, anchor: [9, 7, 13, 9], editAs: 'absolute' } }))).toEqual(['from 7+1000 1+2000', 'to 12+1000 6+2000', ...was.slice(2)])
  })

  it('a range moved over a drawing is refused; a sheet with drawings deleted takes them with it', async () => {
    expect(await moved((s) => s.editStructure({ op: 'moveRange', from: { sheet: 'Other', range: 'H1:M6' }, to: { sheet: 'Other', cell: 'H20' } }))).toBe("moved: xlsx/keep: a drawing on 'Other' (xl/drawings/drawing1.xml) lies in a range that moved, which writeKept does not move a drawing with")
    // a range moved onto all the chart's cells: the chart stays where it is
    const over = (await moved((s) => s.editStructure({ op: 'moveRange', from: { sheet: 'Other', range: 'A20:F25' }, to: { sheet: 'Other', cell: 'H1' } }))) as Map<string, string>
    expect(anchors(over)).toEqual(['from 7+1000 0+2000', 'to 12+1000 5+2000', 'from 7+1000 9+2000', 'to 9+1000 13+2000'])
    expect(over.has('xl/charts/chart1.xml')).toBe(true)
    const parts = (await moved((s) => s.editStructure({ op: 'deleteSheet', sheet: 'Other' }))) as Map<string, string>
    expect([...parts.keys()].filter((p) => /drawing|chart|media/.test(p))).toEqual([])
    expect(parts.get('[Content_Types].xml')).not.toMatch(/drawing1|chart1/)
    // a drawing that loses its only anchor goes, with its sheet's <drawing>
    const alone = (await moved((s) => s.editStructure({ op: 'deleteRows', sheet: 'Other', row: 1, count: 6 }), { picture: undefined })) as Map<string, string>
    expect([...alone.keys()].filter((p) => /drawing|chart/.test(p))).toEqual([])
    expect(alone.get('xl/worksheets/sheet2.xml')).not.toContain('<drawing')
    expect(alone.get('xl/worksheets/_rels/sheet2.xml.rels')).not.toContain('drawing')
  })

  it('the print area, page breaks, a pivot table and its source by cells and by a table renamed', async () => {
    const plain = (extra: object, edit: (s: ReturnType<typeof createSession>) => void) => moved(edit, extra)
    const name = (parts: Map<string, string> | string) => (typeof parts === 'string' ? parts : /<definedName name="_xlnm.Print_Area"[^>]*>([^<]*)</.exec(parts.get('xl/workbook.xml')!)![1])
    expect(name(await plain({ printArea: { sheet: 2, ref: "'Plain'!$A$1:$B$9" } }, (s) => s.editStructure({ op: 'deleteRows', sheet: 'Plain', row: 8 })))).toBe('Plain!$A$1:$B$8')
    const breaks = (parts: Map<string, string> | string) => (typeof parts === 'string' ? parts : (/<rowBreaks[^>]*>.*<\/rowBreaks>|<rowBreaks[^>]*\/>/.exec(parts.get('xl/worksheets/sheet3.xml')!)?.[0] ?? 'none'))
    expect(breaks(await plain({ breaks: { sheet: 2, rows: [4, 9] } }, (s) => s.editStructure({ op: 'insertRows', sheet: 'Plain', row: 3 })))).toBe('<rowBreaks count="2" manualBreakCount="2"><brk id="5" max="16383" man="1"/><brk id="10" max="16383" man="1"/></rowBreaks>')
    // the row a break comes before deleted: the break goes
    expect(breaks(await plain({ breaks: { sheet: 2, rows: [4, 9] } }, (s) => s.editStructure({ op: 'deleteRows', sheet: 'Plain', row: 5 })))).toBe('<rowBreaks count="1" manualBreakCount="1"><brk id="8" max="16383" man="1"/></rowBreaks>')
    const pivot = { pivot: { sheet: 2, location: 'D3:E5', source: { sheet: 'Data', ref: 'A1:B4' } } }
    const pivots = (parts: Map<string, string> | string) => (typeof parts === 'string' ? parts : [/<location ref="([^"]*)"/.exec(parts.get('xl/pivotTables/pivotTable1.xml')!)![1], /<worksheetSource [^>]*\/>/.exec(parts.get('xl/pivotCache/pivotCacheDefinition1.xml')!)![0]])
    expect(pivots(await plain(pivot, (s) => s.editStructure({ op: 'insertRows', sheet: 'Plain', row: 1 })))).toEqual(['D4:E6', '<worksheetSource ref="A1:B4" sheet="Data"/>'])
    expect(pivots(await plain(pivot, (s) => { s.editStructure({ op: 'insertRows', sheet: 'Data', row: 3 }); s.editStructure({ op: 'renameSheet', sheet: 'Data', name: 'Raw' }) }))).toEqual(['D3:E5', '<worksheetSource ref="A1:B5" sheet="Raw"/>'])
    // the source's sheet deleted: Excel makes the source a link to the file by its path, refused; its cells deleted: refused
    expect(await plain(pivot, (s) => s.editStructure({ op: 'deleteSheet', sheet: 'Data' }))).toBe('deleted: xlsx/keep: the source of the pivot cache xl/pivotCache/pivotCacheDefinition1.xml is on a sheet the change deleted, which Excel makes a link to the file by its path')
    expect(await plain(pivot, (s) => s.editStructure({ op: 'deleteRows', sheet: 'Data', row: 1, count: 4 }))).toBe('deleted: xlsx/keep: the source of the pivot cache xl/pivotCache/pivotCacheDefinition1.xml lies in cells the change deleted')
    expect(await plain(pivot, (s) => s.editStructure({ op: 'insertRows', sheet: 'Plain', row: 4 }))).toBe("moved: xlsx/keep: the pivot table xl/pivotTables/pivotTable1.xml on 'Plain' would be cut by the change, which Excel refuses")
    const table: BookInput = { sheets: [{ name: 'Data', cells: { A1: { value: 'k' }, B1: { value: 'v' }, A2: { value: 'x' }, B2: { value: 1 } }, tables: [{ name: 'Sales', range: 'A1:B2', columns: [{ name: 'k' }, { name: 'v' }] }] }, { name: 'Other', cells: {} }] }
    const data = await withParts(table, { pivot: { sheet: 1, location: 'A3:B5', source: { name: 'Sales' } } })
    const r = await openKept(data)
    const s = createSession(r.book, { cached: r.cached })
    s.editStructure({ op: 'renameTable', table: 'Sales', name: 'Orders' })
    expect((await partsOf((await writeKept(data, s.getSnapshot())).data)).get('xl/pivotCache/pivotCacheDefinition1.xml')).toContain('<worksheetSource name="Orders"/>')
    s.editStructure({ op: 'deleteTable', table: 'Orders' })
    await expect(writeKept(data, s.getSnapshot())).rejects.toMatchObject({ code: 'deleted' })
    await expect(writeKept(data, s.getSnapshot())).rejects.toThrow('xlsx/keep: the source of the pivot cache xl/pivotCache/pivotCacheDefinition1.xml is the table or name Sales, which the workbook no longer has')
  })

  it('a book without the carriers of a file with charts: open it with openKept', async () => {
    const { data } = await charted()
    const r = await readXlsx(data)
    await expect(writeKept(data, { book: r.book, result: r.cached })).rejects.toMatchObject({ code: 'carriers' })
    await expect(writeKept(data, { book: r.book, result: r.cached })).rejects.toThrow(/^xlsx\/keep: the workbook has no name _xlcalc\.keep\.\d+, which holds a reference of .*: open the file with openKept and keep the hidden names it adds$/)
  })
})

describe('openKept and writeKept take what readXlsx and writeXlsx take', () => {
  it('openKept: the book read, with hidden names that carry the references of the parts the model does not have', async () => {
    const { data, read } = await charted()
    const plain = await readXlsx(data)
    expect({ ...read.book, names: undefined }).toEqual({ ...plain.book, names: undefined })
    // a carrier of each sheet's last cell, of the chart's series, and of each anchor: its rectangle, the lines below and beside it, its lines
    expect(read.book.names!.every((n) => n.hidden && n.name.startsWith('_xlcalc.keep.'))).toBe(true)
    expect(read.book.names!.slice(0, 3).map((n) => n.ref)).toEqual(["'Data'!$XFD$1048576", "'Other'!$XFD$1048576", "'Plain'!$XFD$1048576"])
    expect(read.book.names!.map((n) => n.ref)).toContain("'Data'!$B$2:$B$4")
    expect(read.book.names!.map((n) => n.ref)).toContain("'Other'!$H$1:$M$6")
    expect(read.book.names!.map((n) => n.ref)).toContain("'Other'!$10:$10")
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

describe("the workbook's default font", () => {
  it('is kept from the file and written as its bytes; another one is refused', async () => {
    const book: BookInput = { font: { name: 'Arial', size: 10 }, styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 1, style: 0 } } }] }
    const data = (await writeXlsx({ book }, { values: 'omit' })).data
    const opened = await openKept(data)
    expect(opened.book.font).toEqual(book.font)
    expect(Buffer.from((await writeKept(data, { book: opened.book })).data).equals(Buffer.from(data))).toBe(true)
    for (const font of [{ name: 'Arial', size: 12 }, undefined])
      await expect(writeKept(data, { book: { ...opened.book, font } })).rejects.toMatchObject({ name: 'KeptRefusal', code: 'font', message: 'xlsx/keep: the default font changed, which writeKept does not write' })
  })
})

describe("a refusal's message", () => {
  it("is in the language of writeKept's locale, a loaded locale's tag or a locale object, en-US by default", async () => {
    const data = (await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] } }, { values: 'omit' })).data
    const book = { ...(await openKept(data)).book, font: { name: 'Arial', size: 12 } }
    const ru = 'xlsx/keep: Изменился шрифт по умолчанию, а writeKept его не записывает'
    for (const locale of ['ru-RU', ruRU]) await expect(writeKept(data, { book }, { locale })).rejects.toMatchObject({ name: 'KeptRefusal', code: 'font', message: ru })
    await expect(writeKept(data, { book }, { locale: 'en-US' })).rejects.toMatchObject({ message: 'xlsx/keep: the default font changed, which writeKept does not write' })
  })
})

describe('the carriers outside writeKept', () => {
  it('writeXlsx leaves them out of a new file; a name of the book stays', async () => {
    const data = (await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 1 } } }], names: [{ name: 'rate', ref: "'S'!$A$1" }] } }, { values: 'omit' })).data
    const opened = await openKept(data)
    expect(opened.book.names?.some((n) => n.name.startsWith('_xlcalc.keep.'))).toBe(true)
    const file = await writeXlsx({ book: opened.book }, { values: 'omit' })
    expect(file.issues).toEqual([])
    expect((await readXlsx(file.data)).book.names).toEqual([{ name: 'rate', ref: "'S'!$A$1" }])
  })
})

describe('validation the workbook could not read is refused, never lost; conditional formats are carried', () => {
  /** S with a rule of each kind the workbook reads, and the file given a conditional format and a validation rule it does not; T without. */
  async function unread(): Promise<Uint8Array> {
    const book: BookInput = {
      sheets: [
        { name: 'S', cells: { A1: { value: 1 }, B1: { value: 2 } }, conditionalFormats: [{ type: 'cellIs', ranges: ['A1:A3'], operator: 'greaterThan', formulas: ['0'], style: { font: { bold: true } } }], validations: [{ range: 'B1', type: 'whole', operator: 'greaterThan', formula1: '0' }] },
        { name: 'T', cells: { A1: { value: 3 } }, conditionalFormats: [{ type: 'cellIs', ranges: ['A1'], operator: 'greaterThan', formulas: ['0'], style: { font: { bold: true } } }] },
      ],
    }
    const archive = openArchive((await writeXlsx({ book })).data)
    const parts = new Map<string, string | Uint8Array>()
    for (const name of archive.names) parts.set(name, (await archive.text(name))!)
    const sheet = parts.get('xl/worksheets/sheet1.xml') as string
    // a rule of a type the workbook does not have, and a rule's IME mode, which it does not keep
    const cf = '<conditionalFormatting sqref="C1:C3"><cfRule type="unknownKind" priority="9"/></conditionalFormatting>'
    const at = sheet.indexOf('<dataValidations')
    parts.set('xl/worksheets/sheet1.xml', (sheet.slice(0, at) + cf + sheet.slice(at)).replace('<dataValidation ', '<dataValidation imeMode="hiragana" '))
    return zip([...parts].map(([name, data]) => ({ name, data: typeof data === 'string' ? new TextEncoder().encode(data) : data })))
  }

  it("the sheet's validation changed, or moved by its rows, is refused with rules; its conditional formats, a value or another sheet's rules are written", async () => {
    const data = await unread()
    expect((await readXlsx(data)).issues.map((x) => [x.code, x.sheet])).toEqual(expect.arrayContaining([['conditionalFormat', 'S'], ['dataValidation', 'S']]))
    const open = async () => createSession((await openKept(data)).book, { clock: { now: '2026-10-27T12:00:00Z', seed: 1 } })
    const rule = { type: 'cellIs' as const, ranges: ['B1'], operator: 'lessThan' as const, formulas: ['9'], style: { font: { italic: true } } }
    for (const edit of [
      (s: ReturnType<typeof createSession>) => s.setLayout([{ sheet: 'S', validations: [] }]),
      (s: ReturnType<typeof createSession>) => s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 }),
    ]) {
      const s = await open()
      edit(s)
      await expect(writeKept(data, s.getSnapshot())).rejects.toMatchObject({ name: 'KeptRefusal', code: 'rules' })
    }
    for (const edit of [
      (s: ReturnType<typeof createSession>) => s.setCells([{ sheet: 'S', cell: 'A2', input: { value: 5 } }]),
      (s: ReturnType<typeof createSession>) => s.setLayout([{ sheet: 'S', conditionalFormats: [...s.getBook().sheets[0].conditionalFormats!, rule] }]),
      (s: ReturnType<typeof createSession>) => s.setLayout([{ sheet: 'T', conditionalFormats: [rule] }]),
    ]) {
      const s = await open()
      edit(s)
      const parts = await partsOf((await writeKept(data, s.getSnapshot())).data)
      expect(parts.get('xl/worksheets/sheet1.xml')).toContain('type="unknownKind"')
      expect(parts.get('xl/worksheets/sheet1.xml')).toContain('imeMode="hiragana"')
    }
  })
})
