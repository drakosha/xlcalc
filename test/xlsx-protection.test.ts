/**
 * Protection in xlsx: a cell format's `<protection>` is a style's (`CellStyle.protection`), a sheet's
 * `<sheetProtection>` its `SheetInput.protection`, the attributes saying the opposite of the model's fields. `writeXlsx`
 * writes both without a password; `writeKept` keeps the file's element while the protection did not change, sets its
 * attributes in place when it did (the password stays), and writes a cell whose look did not change with its own cell
 * format, so `locked="0"` and what the model does not hold of a cell format stay.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { zip } from '../src/xlsx/zip/write.js'

const book = (): BookInput => ({
  styles: [{ protection: { locked: false } }, { font: { bold: true }, protection: { hidden: true } }],
  sheets: [
    { name: 'S', cells: { A1: { value: 1, style: 0 }, A2: { formula: 'A1*2', style: 1 }, A3: { value: 3 } }, protection: { formatCells: true, selectLockedCells: false } },
    { name: 'T', cells: { A1: { value: 'open' } } },
  ],
})

const written = async (b: BookInput) => (await writeXlsx({ book: b, result: compute(b) })).data

async function partText(data: Uint8Array, name: string): Promise<string> {
  return (await openArchive(data).text(name))!
}

/** The file with the text of the parts `edits` names replaced, each `[from, to]` in turn; a text not found throws. */
async function edited(data: Uint8Array, edits: Record<string, [string | RegExp, string][]>): Promise<Uint8Array> {
  const a = openArchive(data)
  const files: { name: string; data: Uint8Array }[] = []
  for (const name of a.names) {
    let text = (await a.text(name))!
    for (const [from, to] of edits[name] ?? []) {
      if (!(typeof from === 'string' ? text.includes(from) : from.test(text))) throw new Error(`no ${from} in ${name}`)
      text = text.replace(from, to)
    }
    files.push({ name, data: edits[name] ? new TextEncoder().encode(text) : (await a.bytes(name))! })
  }
  return zip(files)
}

const sheetProtection = async (data: Uint8Array) => /<sheetProtection[^>]*\/>/.exec(await partText(data, 'xl/worksheets/sheet1.xml'))?.[0]
const cellXml = async (data: Uint8Array, addr: string) => new RegExp(`<c r="${addr}"[^>]*>`).exec(await partText(data, 'xl/worksheets/sheet1.xml'))?.[0]

describe('writeXlsx and readXlsx', () => {
  it("writes a cell format's protection and a sheet's, Excel's default protection with objects and scenarios protected; reads them back", async () => {
    const data = await written(book())
    const styles = await partText(data, 'xl/styles.xml')
    expect(styles).toContain('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf>')
    expect(styles).toContain('<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyProtection="1"><protection hidden="1"/></xf>')
    const xml = await partText(data, 'xl/worksheets/sheet1.xml')
    // after the cells, where the schema puts it
    expect(xml).toContain('</sheetData><sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0" selectLockedCells="1"/>')
    expect(await partText(data, 'xl/worksheets/sheet2.xml')).not.toContain('sheetProtection')
    const r = await readXlsx(data)
    expect(r.book.styles).toEqual(book().styles)
    expect(r.book.sheets.map((s) => s.protection)).toEqual([{ formatCells: true, selectLockedCells: false }, undefined])
    expect(r.issues).toEqual([])
    expect(await sheetProtection(await written({ sheets: [{ name: 'S', cells: {}, protection: {} }] }))).toBe('<sheetProtection sheet="1" objects="1" scenarios="1"/>')
  })

  it('reads a protection as the schema has it: no sheet without sheet="1", a missing attribute by its default, the password reported', async () => {
    const data = await written(book())
    const read = async (tag: string) => (await readXlsx(await edited(data, { 'xl/worksheets/sheet1.xml': [[/<sheetProtection[^>]*\/>/, tag]] })))
    expect((await read('<sheetProtection objects="1"/>')).book.sheets[0].protection).toBeUndefined()
    expect((await read('<sheetProtection sheet="0" objects="1"/>')).book.sheets[0].protection).toBeUndefined()
    expect((await read('<sheetProtection sheet="true" formatRows="false" deleteRows="0"/>')).book.sheets[0].protection).toEqual({ objects: true, scenarios: true, formatRows: true, deleteRows: true })
    const locked = await read('<sheetProtection algorithmName="SHA-512" hashValue="aGFzaA==" saltValue="c2FsdA==" spinCount="100000" sheet="1" objects="1" scenarios="1"/>')
    expect(locked.book.sheets[0].protection).toEqual({})
    expect(locked.issues.map((i) => [i.code, i.sheet, i.message])).toEqual([['protection', 'S', "the sheet protection's password is not in the model: a file written anew protects the sheet without one"]])
  })

  it("reads a cell format's own protection, whatever applyProtection says; a cell style's protection is not the cell's", async () => {
    const data = await written(book())
    const r = await readXlsx(
      await edited(data, {
        'xl/styles.xml': [
          ['applyProtection="1"><protection locked="0"/>', 'applyProtection="0"><protection locked="0"/>'],
          ['<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>', '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1"><protection locked="0" hidden="1"/></xf></cellStyleXfs>'],
        ],
      }),
    )
    expect(r.book.styles).toEqual(book().styles)
    expect(r.book.sheets[0].cells.A3).toEqual({ value: 3 })
  })
})

describe('writeKept', () => {
  it('keeps the file\'s protection while it did not change, sets its attributes in place keeping the password, removes it with the protection', async () => {
    const hash = 'algorithmName="SHA-512" hashValue="aGFzaA==" saltValue="c2FsdA==" spinCount="100000" '
    const original = await edited(await written(book()), { 'xl/worksheets/sheet1.xml': [['<sheetProtection sheet="1"', `<sheetProtection ${hash}sheet="1"`]] })
    const k = await openKept(original)
    const s = createSession(k.book, { cached: k.cached })
    s.setCells([{ sheet: 'S', cell: 'A3', input: { value: 4 } }])
    expect(await sheetProtection((await writeKept(original, s.getSnapshot())).data)).toBe(`<sheetProtection ${hash}sheet="1" objects="1" scenarios="1" formatCells="0" selectLockedCells="1"/>`)
    s.setLayout([{ sheet: 'S', protection: { sort: true, objects: true } }])
    expect(await sheetProtection((await writeKept(original, s.getSnapshot())).data)).toBe(`<sheetProtection ${hash}sheet="1" scenarios="1" sort="0"/>`)
    s.setLayout([{ sheet: 'S', protection: null }, { sheet: 'T', protection: { insertRows: true } }])
    const out = (await writeKept(original, s.getSnapshot())).data
    expect(await sheetProtection(out)).toBeUndefined()
    expect(await partText(out, 'xl/worksheets/sheet2.xml')).toContain('</sheetData><sheetProtection sheet="1" objects="1" scenarios="1" insertRows="0"/>')
    expect((await readXlsx(out)).book.sheets.map((x) => x.protection)).toEqual([undefined, { insertRows: true }])
  })

  it("a cell whose value alone changed keeps its own cell format: protection, a theme's colour, quotePrefix, a cell style's xfId", async () => {
    // cell formats that look as the default to the model: a theme colour of the default font, a quote prefix, another cell style
    const base = await written({ sheets: [{ name: 'S', cells: { A1: { value: 1, style: 0 }, B1: { value: 1 }, C1: { value: 1 }, D1: { value: 1 }, E1: { value: 1 } } }], styles: [{ protection: { locked: false } }] })
    const original = await edited(base, {
      'xl/styles.xml': [
        [/<cellXfs count="2">([\s\S]*?)<\/cellXfs>/, '<cellXfs count="5">$1<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1"/></cellXfs>'],
        [/<fonts count="1">([\s\S]*?)<\/fonts>/, '<fonts count="2">$1<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>'],
        ['<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>', '<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'],
        ['<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>', '<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Mine" xfId="1"/></cellStyles>'],
      ],
      'xl/worksheets/sheet1.xml': [
        ['<c r="B1">', '<c r="B1" s="2">'],
        ['<c r="C1">', '<c r="C1" s="3">'],
        ['<c r="D1">', '<c r="D1" s="4">'],
      ],
    })
    const k = await openKept(original)
    expect(k.book.sheets[0].cells.B1).toEqual({ value: 1 })
    const s = createSession(k.book, { cached: k.cached })
    s.setCells(['A1', 'B1', 'C1', 'D1', 'E1'].map((cell) => ({ sheet: 'S', cell, input: { value: 7 } })))
    const out = (await writeKept(original, s.getSnapshot())).data
    expect(await Promise.all(['A1', 'B1', 'C1', 'D1', 'E1'].map((a) => cellXml(out, a)))).toEqual(['<c r="A1" s="1">', '<c r="B1" s="2">', '<c r="C1" s="3">', '<c r="D1" s="4">', '<c r="E1">'])
    expect(await partText(out, 'xl/styles.xml')).toBe(await partText(original, 'xl/styles.xml'))
  })

  it('a cell that is not locked, moved by rows inserted above it, stays not locked; a new cell takes the cell format of its look', async () => {
    const original = await written({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2, style: 0 } } }], styles: [{ protection: { locked: false } }] })
    const k = await openKept(original)
    const s = createSession(k.book, { cached: k.cached })
    expect(s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 }).ok).toBe(true)
    s.setCells([{ sheet: 'S', cell: 'B1', input: { value: 5, style: 0 } }])
    const out = (await writeKept(original, s.getSnapshot())).data
    expect([await cellXml(out, 'A3'), await cellXml(out, 'B1')]).toEqual(['<c r="A3" s="1">', '<c r="B1" s="1">'])
    expect((await readXlsx(out)).book.styles).toEqual([{ protection: { locked: false } }])
  })
})
