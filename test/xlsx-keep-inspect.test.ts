/**
 * What a host checks of a file without the workbook (`inspectKept`): where its sheets lie and which of their cells have
 * formatted runs of text; and `writeKept`'s `writeValues`, which stores the values of a book with no other change over the
 * file, where it gives the original bytes without it.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { compute, sameValue, type BookInput } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { inspectKept, openKept, writeKept } from '../src/xlsx-keep.js'
import { zip } from '../src/xlsx/zip/write.js'
import { partsOf } from './keep-files.js'

const enc = new TextEncoder()

/** The file with its parts changed by `f`. */
async function edited(data: Uint8Array, f: (parts: Map<string, string>) => void): Promise<Uint8Array> {
  const parts = await partsOf(data)
  f(parts)
  return zip([...parts].map(([name, text]) => ({ name, data: enc.encode(text) })))
}

const made = async (book: BookInput) => (await writeXlsx({ book })).data

describe('inspectKept', () => {
  it("gives each sheet's name and part, in the book's order", async () => {
    const data = await made({ sheets: [{ name: 'Итог', cells: {} }, { name: 'Исходные данные', cells: {} }] })
    const { sheets } = await inspectKept(data)
    expect(sheets.map((s) => [s.name, s.part])).toEqual([
      ['Итог', 'xl/worksheets/sheet1.xml'],
      ['Исходные данные', 'xl/worksheets/sheet2.xml'],
    ])
  })

  it('names the cells with formatted runs of text, shared or inline, and not those with phonetic text alone', async () => {
    const data = await edited(await made({ sheets: [{ name: 'S', cells: { A1: { value: 'a' }, B1: { value: 'b' }, C1: { value: 'c' }, D1: { value: 'd' } } }] }), (parts) => {
      parts.set(
        'xl/sharedStrings.xml',
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><rPr><b/></rPr><t>x</t></r><r><t>y</t></r></si><si><t>plain</t></si><si><t>z</t><rPh sb="0" eb="1"><t>p</t></rPh></si></sst>',
      )
      parts.set(
        'xl/worksheets/sheet1.xml',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="inlineStr"><is><r><t>i</t></r></is></c></row></sheetData></worksheet>',
      )
    })
    const { sheets } = await inspectKept(data)
    expect(sheets[0].formattedText).toEqual(['A1', 'D1'])
  })

  it('refuses options it does not take', async () => {
    await expect(inspectKept(await made({ sheets: [{ name: 'S', cells: {} }] }), { locale: 'en-US' } as never)).rejects.toThrow(/locale/)
  })
})

describe("writeKept's writeValues", () => {
  const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*10' } } }] }
  // a file whose formula has no value stored, as openpyxl writes one
  const stale = async () => (await writeXlsx({ book }, { values: 'omit' })).data

  it('stores the values of a book with no other change, the other parts as the file has them', async () => {
    const data = await stale()
    const { book: read } = await openKept(data)
    const result = compute(read)
    expect((await writeKept(data, { book: read, result })).data).toBe(data)
    const { data: written } = await writeKept(data, { book: read, result }, { writeValues: true })
    expect(written).not.toBe(data)
    const [was, now] = [await partsOf(data), await partsOf(written)]
    expect(now.get('xl/worksheets/sheet1.xml')).toMatch(/<c r="A2"[^>]*><f>A1\*10<\/f><v>20<\/v><\/c>/)
    expect(now.get('xl/styles.xml')).toBe(was.get('xl/styles.xml'))
  })

  it('leaves a formula the result has no value for as the file has it (`values: "preserve"`)', async () => {
    const data = await stale()
    const { book: read } = await openKept(data)
    const { data: written } = await writeKept(data, { book: read, result: { sheets: [{ name: 'S', cells: {} }] } }, { writeValues: true, values: 'preserve' })
    expect((await partsOf(written)).get('xl/worksheets/sheet1.xml')).toMatch(/<c r="A2"[^>]*><f>A1\*10<\/f><\/c>/)
  })

  it('gives the original bytes where no value differs, or without a result', async () => {
    const { data } = await writeXlsx({ book, result: compute(book) })
    const { book: read } = await openKept(data)
    expect((await writeKept(data, { book: read, result: compute(read) }, { writeValues: true })).data).toBe(data)
    const none = await stale()
    expect((await writeKept(none, { book: (await openKept(none)).book }, { writeValues: true })).data).toBe(none)
  })

  it('takes a boolean', async () => {
    const data = await stale()
    await expect(writeKept(data, { book: (await openKept(data)).book }, { writeValues: 1 } as never)).rejects.toThrow(/writeValues/)
  })

  describe('over a file Excel saved, with shared and array formulas and values of each type', () => {
    const saved = new Uint8Array(readFileSync(new URL('./excel/xlsx/saved/areas-1.xlsx', import.meta.url)))
    const FORMULA_CELL = /<c\b([^>]*)>(<f\b[^>]*?(?:\/>|>[\s\S]*?<\/f>))[\s\S]*?<\/c>/g
    const isSheet = (name: string) => /^xl\/worksheets\/[^/]+\.xml$/.test(name)
    /** The file with no value on its formulas, as a generator writes one. */
    const withoutValues = () => edited(saved, (parts) => {
      for (const [name, xml] of parts) if (isSheet(name)) parts.set(name, xml.replace(FORMULA_CELL, (_, a: string, f: string) => `<c${a.replace(/ (?:t|vm)="\w+"/g, '')}>${f}</c>`))
    })
    /** Each sheet's formulas as the file has them, by cell. */
    const formulas = (parts: Map<string, string>) => new Map([...parts].filter(([name]) => isSheet(name)).map(([name, xml]) => [name, [...xml.matchAll(FORMULA_CELL)].map((m) => `${/ r="\w+"/.exec(m[1])![0]} ${m[2]}`)]))
    const sheetC = (parts: Map<string, string>) => parts.get('xl/worksheets/sheet5.xml')!
    const cellOf = (xml: string, addr: string) => new RegExp(`<c r="${addr}"[^>]*>[\\s\\S]*?</c>`).exec(xml)![0]

    it("stores each formula's value and leaves the formulas, the book and the other parts as the file has them", async () => {
      const data = await withoutValues()
      const { book: read } = await openKept(data)
      const result = compute(read)
      const { data: written, issues } = await writeKept(data, { book: read, result }, { writeValues: true })
      expect(issues).toEqual([])
      const [was, now] = [await partsOf(data), await partsOf(written)]
      expect(formulas(now)).toEqual(formulas(was))
      // the parts a change of the file's calculation takes along, and no other
      expect([...new Set([...was.keys(), ...now.keys()])].filter((name) => !isSheet(name) && was.get(name) !== now.get(name)).sort()).toEqual(['[Content_Types].xml', 'xl/_rels/workbook.xml.rels', 'xl/calcChain.xml', 'xl/workbook.xml'])
      expect(now.has('xl/calcChain.xml')).toBe(false)
      expect(now.get('xl/workbook.xml')).toContain('fullCalcOnLoad="1"')
      const [back, excel] = [await readXlsx(written), await readXlsx(saved)]
      expect(back.book).toEqual((await readXlsx(data)).book)
      const types = new Set([...now].filter(([name]) => isSheet(name)).flatMap(([, xml]) => [...xml.matchAll(FORMULA_CELL)].map((m) => / t="(\w+)"/.exec(m[1])?.[1])))
      expect(types).toEqual(new Set([undefined, 'str', 'b', 'e']))
      // the values read back are the result's; Excel's own, but for the last digits of what the engine computes (HARMEAN, KURT)
      const near = (a: unknown, b: unknown) => typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) <= 1e-12 * Math.abs(b)
      for (const [i, sheet] of back.cached.sheets.entries())
        for (const [addr, { value }] of Object.entries(result.sheets[i].cells)) {
          if (read.sheets[i].cells[addr]?.formula === undefined) continue
          expect(sameValue(sheet.cells[addr]?.value, value), `${sheet.name}!${addr}`).toBe(true)
          const stored = excel.cached.sheets[i].cells[addr]?.value
          expect(sameValue(stored, value) || near(stored, value), `${sheet.name}!${addr} in Excel`).toBe(true)
        }
    }, 30_000)

    it('gives the same bytes when written over again, the values stored', async () => {
      const data = await withoutValues()
      const opened = await openKept(data)
      const { data: written } = await writeKept(data, { book: opened.book, result: compute(opened.book) }, { writeValues: true })
      const again = await openKept(written)
      expect((await writeKept(written, { book: again.book, result: compute(again.book) }, { writeValues: true })).data).toBe(written)
    }, 30_000)

    it('writes a string, a boolean, an error and a number by their types, over a shared formula and an array one', async () => {
      const { book: read } = await openKept(saved)
      const result = compute(read)
      Object.assign(result.sheets[4].cells, { C33: { value: 'a<b&c' }, C34: { value: true }, C35: { value: { error: 'NA' } }, A14: { value: 5 } })
      const { data: written } = await writeKept(saved, { book: read, result }, { writeValues: true })
      const [was, now] = [sheetC(await partsOf(saved)), sheetC(await partsOf(written))]
      expect(cellOf(was, 'A14')).toMatch(/ t="e" cm="1" vm="1"/)
      expect(cellOf(now, 'C33')).toBe('<c r="C33" t="str"><f t="shared" si="1"/><v>a&lt;b&amp;c</v></c>')
      expect(cellOf(now, 'C34')).toBe('<c r="C34" t="b"><f t="shared" si="1"/><v>1</v></c>')
      expect(cellOf(now, 'C35')).toBe('<c r="C35" t="e"><f t="shared" si="1"/><v>#N/A</v></c>')
      // a number leaves the error's type and rich value behind, the array's mark and formula stay
      const f = /<f\b[\s\S]*<\/f>/.exec(cellOf(was, 'A14'))![0]
      expect(cellOf(now, 'A14')).toBe(`<c r="A14" cm="1">${f}<v>5</v></c>`)
    }, 30_000)

    it('asks once for a full calculation in a file that already does', async () => {
      const data = await edited(saved, (parts) => parts.set('xl/workbook.xml', parts.get('xl/workbook.xml')!.replace('<calcPr ', '<calcPr fullCalcOnLoad="1" ')))
      const { book: read } = await openKept(data)
      const result = compute(read)
      expect((await writeKept(data, { book: read, result })).data).toBe(data)
      const { data: written } = await writeKept(data, { book: read, result }, { writeValues: true })
      expect((await partsOf(written)).get('xl/workbook.xml')!.match(/fullCalcOnLoad/g)).toHaveLength(1)
    }, 30_000)
  })
})
