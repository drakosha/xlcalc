/**
 * `xlcalc/xlsx/keep`: a cell, row or column whose cell format the workbook does not tell from an earlier one of the file
 * (a quote prefix, a cell style, a theme colour, a pattern fill) keeps it when rows or columns are inserted or deleted,
 * and a column keeps it when the columns are written again; a cell that comes to the address of such a cell does not
 * take it. Formats are compared by their XML, not by their numbers.
 */
import { describe, expect, it } from 'vitest'
import { createSession, type BookInput, type StructureOp } from '../src/index.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { KEPT_FORMATS } from '../src/xlsx/keep/carriers.js'
import { formatAddress } from '../src/model.js'
import { partsOf, withParts } from './keep-files.js'
import { zip } from '../src/xlsx/zip/write.js'

/**
 * The file's cell formats past `writeXlsx`'s two (0 plain, 1 bold #4472C4), each looking as one of them to the
 * workbook: 2 plain with a quote prefix, 3 bold in the theme's first accent (#4472C4), 4 plain in a cell style, 5 plain
 * with a pattern fill.
 */
const TWINS = [
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1"/>',
  '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1"/>',
  '<xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyFill="1"/>',
]

/**
 * A file of the sheet's cells with the twin formats (a bold cell far off gives `writeXlsx` its bold format), cells and
 * rows given them by address (`A1`, `8`), a `<cols>` of `cols`.
 */
async function twinned(cells: SheetCells, formats: Record<string, number>, cols = ''): Promise<Uint8Array> {
  const parts = await partsOf(await withParts({ sheets: [{ name: 'S', cells: { ...cells, Z1000: { value: 0, style: 0 } } }], styles: [{ font: { bold: true, color: '#4472C4' } }] }, {}))
  let styles = parts.get('xl/styles.xml')!
  styles = styles
    .replace(/<fonts count="2">(.*?)<\/fonts>/, '<fonts count="3">$1<font><b/><sz val="11"/><color theme="4"/><name val="Calibri"/></font></fonts>')
    .replace(/<cellStyleXfs count="1">(.*?)<\/cellStyleXfs>/, '<cellStyleXfs count="2">$1<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
    .replace(/<cellXfs count="2">(.*?)<\/cellXfs>/, `<cellXfs count="${2 + TWINS.length}">$1${TWINS.join('')}</cellXfs>`)
    .replace(/<cellStyles count="1">(.*?)<\/cellStyles>/, '<cellStyles count="2">$1<cellStyle name="Kept" xfId="1"/></cellStyles>')
  let sheet = parts.get('xl/worksheets/sheet1.xml')!
  for (const [at, s] of Object.entries(formats)) {
    if (/^\d+$/.test(at)) sheet = sheet.replace(new RegExp(`<row r="${at}"( [^>]*)?>`), `<row r="${at}" s="${s}" customFormat="1">`)
    else sheet = sheet.replace(new RegExp(`<c r="${at}"(?: s="\\d+")?`), `<c r="${at}" s="${s}"`)
  }
  if (cols) sheet = sheet.replace(/<sheetData/, `<cols>${cols}</cols><sheetData`)
  parts.set('xl/styles.xml', styles)
  // a theme whose first accent is #4472C4
  parts.set('xl/theme/theme1.xml', `<a:theme xmlns:a="${DRAWING}" name="T"><a:themeElements><a:clrScheme name="T"><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme></a:themeElements></a:theme>`)
  parts.set('xl/_rels/workbook.xml.rels', parts.get('xl/_rels/workbook.xml.rels')!.replace('</Relationships>', `<Relationship Id="rIdT" Type="${REL}/theme" Target="theme/theme1.xml"/></Relationships>`))
  parts.set('[Content_Types].xml', parts.get('[Content_Types].xml')!.replace('</Types>', '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'))
  parts.set('xl/worksheets/sheet1.xml', sheet)
  return zip([...parts].map(([name, text]) => ({ name, data: new TextEncoder().encode(text) })))
}

/** The cell formats of a file's styles, as written. */
const xfsOf = (styles: string) => /<cellXfs[^>]*>(.*?)<\/cellXfs>/.exec(styles)![1].match(/<xf\b[^>]*?(?:\/>|>.*?<\/xf>)/g)!

/** What a file gives a cell, a row (`8`) or a column (`D`) of its first sheet: its cell format's XML. */
async function formatOf(data: Uint8Array): Promise<(at: string) => string | undefined> {
  const parts = await partsOf(data)
  const xfs = xfsOf(parts.get('xl/styles.xml')!)
  const sheet = parts.get('xl/worksheets/sheet1.xml')!
  return (at) => {
    let tag: string | undefined
    if (/^\d+$/.test(at)) tag = new RegExp(`<row r="${at}"[^>]*>`).exec(sheet)?.[0]
    else if (/^[A-Z]+$/.test(at)) {
      const n = [...at].reduce((k, ch) => k * 26 + ch.charCodeAt(0) - 64, 0)
      tag = [...sheet.matchAll(/<col [^>]*>/g)].map((m) => m[0]).find((t) => Number(/min="(\d+)"/.exec(t)![1]) <= n && n <= Number(/max="(\d+)"/.exec(t)![1]))
    } else tag = new RegExp(`<c r="${at}"[^>]*>`).exec(sheet)?.[0]
    if (tag === undefined) return undefined
    const s = new RegExp(`\\s${/^[A-Z]+$/.test(at) ? 'style' : 's'}="(\\d+)"`).exec(tag)?.[1]
    return xfs[Number(s ?? 0)]
  }
}

/** The file written after the operations, from a session over the file opened kept. */
async function after(data: Uint8Array, ops: StructureOp[], edit?: (s: ReturnType<typeof createSession>) => void) {
  const read = await openKept(data)
  const s = createSession(read.book, { cached: read.cached })
  for (const op of ops) expect(s.editStructure(op).ok).toBe(true)
  edit?.(s)
  const w = await writeKept(data, s.getSnapshot())
  s.close()
  return w
}

const DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/** `writeXlsx`'s plain cell format. */
const PLAIN = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'

type SheetCells = BookInput['sheets'][number]['cells']

const values = (addrs: string[], style?: number): SheetCells => Object.fromEntries(addrs.map((a, i) => [a, { value: i + 1, ...(style === undefined ? {} : { style }) }]))

describe('writeKept: cell formats the workbook does not tell apart, where rows and columns move', () => {
  it('a moved cell keeps its own format; one that comes to the address of a formatted cell does not take it', async () => {
    // A1:A4 one of each twin; C1 a quote prefix over C2's same value; E1 plain over E2's quote prefix
    const cells = { ...values(['A1', 'A2', 'A3', 'A4']), A2: { value: 2, style: 0 }, C1: { value: 7 }, C2: { value: 7 }, E1: { value: 1 }, E2: { value: 2 } }
    const data = await twinned(cells, { A1: 2, A2: 3, A3: 4, A4: 5, C1: 2, E2: 2 })
    const was = await formatOf(data)
    // each a twin of the workbook's: a carrier holds each format's cells
    const carried = (await openKept(data)).book.names!.filter((n) => n.name.startsWith('_xlcalc.keep.x')).map((n) => n.ref)
    expect(carried.sort()).toEqual(["'S'!$A$1,'S'!$C$1,'S'!$E$2", "'S'!$A$2", "'S'!$A$3", "'S'!$A$4"])
    for (const op of [{ op: 'insertRows', sheet: 'S', row: 1 }, { op: 'insertRows', sheet: 'S', row: 1, count: 3 }] as StructureOp[]) {
      const k = op.op === 'insertRows' ? (op.count ?? 1) : 0
      const now = await formatOf((await after(data, [op])).data)
      for (const a of ['A1', 'A2', 'A3', 'A4', 'C1', 'C2', 'E1', 'E2']) expect(now(`${a[0]}${Number(a.slice(1)) + k}`), a).toBe(was(a))
    }
    const up = await formatOf((await after(data, [{ op: 'deleteRows', sheet: 'S', row: 1 }])).data)
    expect([up('A1'), up('A2'), up('A3'), up('C1'), up('E1')]).toEqual([was('A2'), was('A3'), was('A4'), was('C2'), was('E2')])
    const right = await formatOf((await after(data, [{ op: 'insertColumns', sheet: 'S', column: 'A', count: 2 }])).data)
    expect(['C1', 'C2', 'C3', 'C4', 'E1', 'G2'].map(right)).toEqual(['A1', 'A2', 'A3', 'A4', 'C1', 'E2'].map(was))
    const left = await formatOf((await after(data, [{ op: 'deleteColumns', sheet: 'S', column: 'B' }])).data)
    expect(['A1', 'B1', 'B2', 'D2'].map(left)).toEqual(['A1', 'C1', 'C2', 'E2'].map(was))
  })

  it('a cell moved and edited keeps its format while it looks the same, and takes the first of its look when it does not', async () => {
    const data = await twinned({ A1: { value: 'x' }, A2: { value: 'y', style: 0 } }, { A1: 2, A2: 3 })
    const was = await formatOf(data)
    const w = await after(data, [{ op: 'insertRows', sheet: 'S', row: 1 }], (s) =>
      s.setCells([
        { sheet: 'S', cell: 'A2', input: { value: 'x2' } },
        { sheet: 'S', cell: 'A3', input: { value: 'y2' } },
      ]),
    )
    const now = await formatOf(w.data)
    expect([now('A2'), now('A3')]).toEqual([was('A1'), was('A2')])
    const plain = await formatOf((await after(data, [{ op: 'insertRows', sheet: 'S', row: 1 }], (s) => s.patchCells([{ sheet: 'S', cell: 'A3', set: { style: null } }]))).data)
    expect(plain('A3')).toBe(PLAIN)
  })

  it('rows inserted inside a block of a format give the new cells of its look that format, as the row above has it', async () => {
    const data = await twinned(values(['B1', 'B2', 'B3', 'B4'], 0), { B1: 3, B2: 3, B3: 3, B4: 3 })
    const was = await formatOf(data)
    const now = await formatOf((await after(data, [{ op: 'insertRows', sheet: 'S', row: 3, count: 2 }])).data)
    expect(['B1', 'B2', 'B3', 'B4', 'B5', 'B6'].map(now)).toEqual(Array(6).fill(was('B1')))
  })

  it("a row's format and a column's move with them, and a column keeps its own when the columns are written again", async () => {
    const cells = { ...values(['A1', 'A8', 'D2']), A9: { value: 9 } }
    const cols = '<col min="4" max="5" width="12" customWidth="1" style="3"/><col min="6" max="6" width="9" customWidth="1" style="1"/>'
    const data = await twinned(cells, { 8: 3, 9: 1 }, cols)
    const was = await formatOf(data)
    expect([was('8'), was('D')]).toEqual([TWINS[1], TWINS[1]])
    const down = await formatOf((await after(data, [{ op: 'insertRows', sheet: 'S', row: 2, count: 2 }, { op: 'insertColumns', sheet: 'S', column: 'B' }])).data)
    expect([down('10'), down('11'), down('E'), down('F'), down('G')]).toEqual([was('8'), was('9'), was('D'), was('E'), was('F')])
    const read = await openKept(data)
    const s = createSession(read.book, { cached: read.cached })
    s.setLayout([{ sheet: 'S', columns: 'D', set: { width: 20 } }])
    const wide = await formatOf((await writeKept(data, s.getSnapshot())).data)
    expect([wide('D'), wide('E'), wide('F')]).toEqual([was('D'), was('E'), was('F')])
  })

  it('a sheet whose every cell is in a cell style takes a carrier per block, and its cells keep the style', async () => {
    const addrs = Array.from({ length: 200 }, (_, r) => ['A', 'B', 'C', 'D', 'E'].map((c) => `${c}${r + 1}`)).flat()
    const data = await twinned(values(addrs), Object.fromEntries(addrs.map((a) => [a, 4])))
    const read = await openKept(data)
    expect(read.book.names!.filter((n) => n.name.startsWith('_xlcalc.keep.x')).map((n) => n.ref)).toEqual(["'S'!$A$1:$E$200"])
    const now = await formatOf((await after(data, [{ op: 'insertRows', sheet: 'S', row: 50, count: 3 }, { op: 'deleteColumns', sheet: 'S', column: 'A' }])).data)
    const was = await formatOf(data)
    expect(['A1', 'D49', 'A53', 'D203'].every((a) => now(a) === was('A1'))).toBe(true)
  })

  it(`past ${KEPT_FORMATS} blocks on a sheet its formats have no carriers, and a change of its rows or columns says so`, async () => {
    const addrs: string[] = []
    // cells apart from one another: each a block of its own
    for (let r = 0; addrs.length <= KEPT_FORMATS; r += 2) for (let c = 0; c < 100; c += 2) addrs.push(formatAddress(r, c))
    const data = await twinned(values(addrs), Object.fromEntries(addrs.map((a) => [a, 2])))
    const read = await openKept(data)
    expect(read.book.names!.some((n) => n.name.startsWith('_xlcalc.keep.x'))).toBe(false)
    const w = await after(data, [{ op: 'insertRows', sheet: 'S', row: 1 }])
    expect(w.issues).toEqual([{ code: 'style', sheet: 'S', message: expect.stringContaining(`sheet whose ${addrs.length} cells in cell formats`) }])
  })

  it('a host that drops the carriers of formats leaves the cells their formats by address, without a refusal', async () => {
    const data = await twinned({ A1: { value: 1 }, A2: { value: 2 } }, { A1: 2 })
    const read = await openKept(data)
    read.book.names = read.book.names!.filter((n) => !n.name.startsWith('_xlcalc.keep.x'))
    const s = createSession(read.book, { cached: read.cached })
    expect(s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 }).ok).toBe(true)
    const w = await writeKept(data, s.getSnapshot())
    const now = await formatOf(w.data)
    expect(now('A2')).toBe((await formatOf(data))('A2'))
  })
})
