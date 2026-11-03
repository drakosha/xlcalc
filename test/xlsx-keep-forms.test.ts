/**
 * `xlcalc/xlsx/keep` writes a cell in the oldest form that loses nothing: a formula the move of rows only readdressed,
 * or one a host typed, goes as a formula of the old kind (`<f>`) where Excel reads that as the workbook computes it, and
 * as a dynamic array (`cm`, the XLDAPR metadata) only where it spills or the old kind would compute it otherwise. A row or
 * a cell written again keeps the extensions the file gives it (`<extLst>`), a cell's moving with it as Excel moves them.
 */
import { describe, expect, it } from 'vitest'
import { createSession, type BookInput } from '../src/index.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { partsOf, withParts } from './keep-files.js'
import './locales.js'

/** A cell of a sheet's XML by its address. */
const cellOf = (xml: string | undefined, addr: string) => new RegExp(`<c r="${addr}"[^>]*?(?:/>|>.*?</c>)`).exec(xml ?? '')?.[0]

/** A cell of `writeXlsx`'s with a number given a formula of the old kind, as a file Excel 2016 saved holds it. */
const old = (addr: string, formula: string) => (x: string) => x.replace(new RegExp(`<c r="${addr}">(<v>[^<]*</v>)</c>`), `<c r="${addr}"><f>${formula}</f>$1</c>`)

const values: BookInput = {
  sheets: [
    { name: 'Лист1', cells: { A1: { value: 1 }, A3: { value: 3 }, B5: { value: 6 } } },
    { name: 'Ф', cells: { A1: { value: 30 }, C1: { value: 10 } } },
  ],
}

/** Formulas of the old kind on both sheets, no dynamic array in the file and no cell metadata. */
const plainFile = () =>
  withParts(values, {
    edit: {
      'xl/worksheets/sheet1.xml': old('B5', 'A3*2'),
      'xl/worksheets/sheet2.xml': (x) => old('C1', 'Лист1!A1:A3*10')(old('A1', "'Лист1'!A3*10")(x)),
    },
  })

async function written(data: Uint8Array, edit: (s: ReturnType<typeof createSession>) => void) {
  const read = await openKept(data, { locale: 'ru-RU' })
  const s = createSession(read.book, { cached: read.cached })
  edit(s)
  return partsOf((await writeKept(data, s.getSnapshot(), { locale: 'ru-RU' })).data)
}

describe('writeKept: formulas in the oldest form that loses nothing', () => {
  it('formulas of the old kind that rows inserted readdressed stay of the old kind, on their sheet and another, without cell metadata', async () => {
    const data = await plainFile()
    const parts = await written(data, (s) => s.editStructure({ op: 'insertRows', sheet: 'Лист1', row: 2 }))
    expect(cellOf(parts.get('xl/worksheets/sheet1.xml'), 'B6')).toBe('<c r="B6"><f>A4*2</f><v>6</v></c>')
    const f = parts.get('xl/worksheets/sheet2.xml')
    expect(cellOf(f, 'A1')).toBe(`<c r="A1"><f>'Лист1'!A4*10</f><v>30</v></c>`)
    // the implicit intersection the old kind does itself: `@` in the workbook, none in the file
    expect(cellOf(f, 'C1')).toBe(`<c r="C1"><f>'Лист1'!A1:A4*10</f><v>10</v></c>`)
    expect([...parts.keys()].filter((n) => /metadata/i.test(n))).toEqual([])
    expect(parts.get('[Content_Types].xml')).toBe((await partsOf(data)).get('[Content_Types].xml'))
  })

  it('a formula a host typed is of the old kind where that computes the same; one that spills or computes an array is a dynamic array', async () => {
    const parts = await written(await plainFile(), (s) =>
      s.setCells([
        { sheet: 'Ф', cell: 'B1', input: { formula: 'Лист1!A1+1' } },
        { sheet: 'Ф', cell: 'D1', input: { formula: 'SUM(IF(Лист1!A1:A3>1,1,0))' } },
        { sheet: 'Ф', cell: 'E1', input: { formula: 'SEQUENCE(2)' } },
      ]),
    )
    const f = parts.get('xl/worksheets/sheet2.xml')
    expect(cellOf(f, 'B1')).toBe(`<c r="B1"><f>'Лист1'!A1+1</f><v>2</v></c>`)
    expect(cellOf(f, 'D1')).toBe(`<c r="D1" cm="1"><f t="array" ref="D1">SUM(IF('Лист1'!A1:A3&gt;1,1,0))</f><v>1</v></c>`)
    expect(cellOf(f, 'E1')).toMatch(/^<c r="E1" cm="1"><f t="array" ref="E1:E2">_xlfn\.SEQUENCE\(2\)<\/f>/)
    // the dynamic arrays' mark added, in a part of its own as the file had none
    expect([...parts].filter(([n]) => /metadata/i.test(n)).map(([, x]) => x.includes('XLDAPR'))).toEqual([true])
  })

  it('a dynamic array of the file that spills stays one after rows inserted', async () => {
    const book: BookInput = { sheets: [values.sheets[0], { name: 'Ф', cells: { F1: { formula: 'Лист1!A1:A3*10' } } }] }
    const data = await withParts(book, {})
    const parts = await written(data, (s) => s.editStructure({ op: 'insertRows', sheet: 'Лист1', row: 2 }))
    expect(cellOf(parts.get('xl/worksheets/sheet2.xml'), 'F1')).toBe(`<c r="F1" cm="1"><f t="array" ref="F1:F4">'Лист1'!A1:A4*10</f><v>10</v></c>`)
    expect(parts.get('xl/metadata.xml')).toBe((await partsOf(data)).get('xl/metadata.xml'))
  })
})

describe('writeKept: the extensions of rows and cells', () => {
  /** Row 1 and cell A2 with extensions of their own, as Excel 16.114 opened and saved them: the cell's moved with it, the row's dropped. */
  const extended = () =>
    withParts(
      { sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 5 } } }] },
      { edit: { 'xl/worksheets/sheet1.xml': (x) => x.replace('<c r="A1"><v>1</v></c>', '$&<extLst><ext uri="{x}"/></extLst>').replace('<v>5</v></c>', '<v>5</v><extLst><ext uri="{y}"/></extLst></c>') } },
    )
  const sheet = async (edit: (s: ReturnType<typeof createSession>) => void) => (await written(await extended(), edit)).get('xl/worksheets/sheet1.xml')!

  it("a row written again keeps its extensions, a cell whose value changed its own", async () => {
    const xml = await sheet((s) => s.setCells([{ sheet: 'S', cell: 'B1', input: { value: 2 } }, { sheet: 'S', cell: 'A2', input: { value: 6 } }]))
    expect(xml).toContain('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><extLst><ext uri="{x}"/></extLst></row>')
    expect(xml).toContain('<c r="A2"><v>6</v><extLst><ext uri="{y}"/></extLst></c>')
  })

  it("a cell's extensions move with it as rows are inserted above it", async () => {
    const xml = await sheet((s) => s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 }))
    expect(cellOf(xml, 'A3')).toBe('<c r="A3"><v>5</v><extLst><ext uri="{y}"/></extLst></c>')
    expect(cellOf(xml, 'A2')).toBe('<c r="A2"><v>1</v></c>')
  })
})
