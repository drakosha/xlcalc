/**
 * The form `writeXlsx` writes a formula in, against Excel: Excel 365 for Mac and for Windows saved formulas of each kind
 * typed into a book (`bench/coverage/xlsx/forms.ts`), of the old kind (`<f>`) or as dynamic arrays (`cm`), and the
 * writer writes each as Excel saved it, but where the old kind loses nothing and Excel still made a dynamic array. Excel
 * for Mac opened the writer's file and showed every formula as the book has it, with the engine's values.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { FORM_CELLS, formsBook, savedForms, type FormSaved } from '../bench/coverage/xlsx/forms.js'
import { compareCached } from '../bench/coverage/xlsx/compare.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import type { StoredRun } from '../bench/coverage/xlsx.js'
import { compute } from '../src/index.js'
import { fromLegacy, readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import type { BookInput } from '../src/contract.js'

const file = (path: string) => new Uint8Array(readFileSync(new URL(`./excel/${path}`, import.meta.url)))

/** Where the writer's form is not the one Excel saved, and why. */
const OTHERWISE: Record<string, string> = {
  E8: 'SUMPRODUCT((A1:A5>2)*B1:B5): Excel makes a dynamic array of an operation over a range, though Excel before 365 computes it alike in SUMPRODUCT',
  E11: 'MAX(INDEX(A1:A5*B1:B5,0)): the same, INDEX with 0 taking the array whole',
  E20: 'SUM(FILTER(A1:A5,A1:A5>2)): the same; Excel 365 reads FILTER of the old kind without `@`',
}

const book = formsBook()
const result = compute(book)
const written = async () => (await writeXlsx({ book, result })).data

describe('writeXlsx: a formula in the form Excel 365 saves it in', () => {
  for (const [os, path] of [['Mac', 'xlsx/forms/forms.xlsx'], ['Windows', 'win/xlsx/forms/forms.xlsx']] as const)
    it(`each formula of the old kind or a dynamic array as Excel for ${os} saved it, its text as Excel wrote it`, async () => {
      const excel = await savedForms(file(path))
      const ours = await savedForms(await written())
      const misses: string[] = []
      for (const a of FORM_CELLS) {
        const [e, o] = [excel[a], ours[a]] as FormSaved[]
        // a shared formula's copies hold no text of their own
        const same = e.form === o.form && (e.text === undefined || e.text === o.text) && e.ref === o.ref
        if (!same) misses.push(a)
      }
      expect(misses).toEqual(Object.keys(OTHERWISE).sort((x, y) => Number(x.slice(1)) - Number(y.slice(1))))
    })

  it('Excel for Mac shows every formula of the file written as the book has it, with the values the engine computes', async () => {
    const data = await written()
    const run = JSON.parse(readFileSync(new URL('./excel/xlsx/forms.json', import.meta.url), 'utf8')) as StoredRun
    const sent = run.books.written
    expect(sent?.sent, 'the writer changed since Excel saw the file: run the xlsx forms step').toBe(await partsHash(data))
    expect(sent.shown).toEqual(Object.fromEntries(FORM_CELLS.map((a) => [a, `=${book.sheets[0].cells[a].formula}`])))
    expect(compareCached(result, (await readXlsx(file('xlsx/forms/written.xlsx'))).cached)).toEqual([])
  })

  it('a book with no dynamic array formula has no cell metadata', async () => {
    const plain: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { formula: 'A1*2' }, C1: { formula: 'SUMPRODUCT(A1:A3,A1:A3)' } } }] }
    const { data } = await writeXlsx({ book: plain, result: compute(plain) })
    const read = await readXlsx(data)
    expect(read.book.sheets[0].cells).toEqual(plain.sheets[0].cells)
    expect(openArchive(data).names.filter((n) => /metadata/.test(n))).toEqual([])
  })

  it('formulas of one shape take the form of their own references: a cell, a range, a number that is a place or not', async () => {
    const cells: BookInput['sheets'][0]['cells'] = { A1: { value: 1 }, A2: { value: 2 }, B1: { value: 3 }, B2: { value: 4 } }
    // one shape: `SUM(A1*2)` and `SUM(A1:A2*2)` (the old kind intersects the range), INDEX at a place, at a fraction and past 65535, at MATCH
    const formulas = { D1: 'SUM(A1*2)', D2: 'SUM(A1:A2*2)', E2: 'SUM(A2*2)', F1: 'INDEX(A1:B2,1,2)', F2: 'INDEX(A1:B2,1.5,2)', G1: 'INDEX(A1:B2,2,1)', G2: 'INDEX(A1:B2,70000,1)', H1: 'INDEX(A1:B2,MATCH(2,A1:A2,0),1)' }
    for (const [a, f] of Object.entries(formulas)) cells[a] = { formula: f }
    const b: BookInput = { sheets: [{ name: 'S', cells }] }
    const xml = (await openArchive((await writeXlsx({ book: b, result: compute(b) })).data).text('xl/worksheets/sheet1.xml'))!
    const form = (a: string) => (new RegExp(`<c r="${a}"[^>]*>`).exec(xml)![0].includes('cm=') ? 'dynamic' : 'old')
    expect(Object.fromEntries(Object.keys(formulas).map((a) => [a, form(a)]))).toEqual({ D1: 'old', D2: 'dynamic', E2: 'old', F1: 'old', F2: 'dynamic', G1: 'old', G2: 'dynamic', H1: 'old' })
  })

  it("`@` the old kind does itself goes; on one cell it stays as _xlfn.SINGLE; a LAMBDA's call is a dynamic array, as Excel saves them", async () => {
    const b: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, B2: { formula: '@A1:A2*2' }, C1: { formula: '@A1+1' }, D1: { formula: 'LAMBDA(x,x*2)(A2)' } } }] }
    const { data } = await writeXlsx({ book: b, result: compute(b) })
    const xml = (await openArchive(data).text('xl/worksheets/sheet1.xml'))!
    expect(xml).toContain('<c r="B2"><f>A1:A2*2</f><v>4</v></c>')
    expect(xml).toContain('<c r="C1"><f>_xlfn.SINGLE(A1)+1</f><v>2</v></c>')
    expect(xml).toContain('<c r="D1" cm="1"><f t="array" ref="D1">_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(A2)</f><v>4</v></c>')
    expect((await readXlsx(data)).book.sheets[0].cells).toEqual(b.sheets[0].cells)
  })

  it('a formula over another workbook is of the old kind where a sheet\'s would be: a range where one value goes keeps it a dynamic array', async () => {
    const formulas = { A1: '[1]Data!A1*2', A2: 'A1*2', A3: 'SUM([1]Data!A:A)', A4: 'VLOOKUP(7,[1]Data!A1:B3,2,FALSE)', A5: 'ISBLANK([1]Data!Z9)', C1: '[1]Data!A1:A3', D1: 'ABS([1]Data!A1:A3)', E1: 'SUM(IF([1]Data!A1:A3>1,1,0))', F1: '[1]!Rate' }
    const b: BookInput = {
      sheets: [{ name: 'S', cells: Object.fromEntries(Object.entries(formulas).map(([a, formula]) => [a, { formula }])) }],
      externalBooks: [{ target: 'Book2.xlsx', sheets: [{ name: 'Data', cells: { A1: 7, A2: 8, A3: 9, B1: 1 } }], names: [{ name: 'Rate', ref: 'Data!$A$1' }] }],
    }
    const xml = (await openArchive((await writeXlsx({ book: b, result: compute(b) })).data).text('xl/worksheets/sheet1.xml'))!
    expect(xml).toContain('<c r="A1"><f>[1]Data!A1*2</f><v>14</v></c>')
    expect(xml).toContain('<c r="A2"><f>A1*2</f><v>28</v></c>')
    for (const a of ['A3', 'A4', 'A5']) expect(xml).toMatch(new RegExp(`<c r="${a}"[^>]*><f>`))
    for (const a of ['C1', 'D1', 'E1', 'F1']) expect(xml).toMatch(new RegExp(`<c r="${a}" cm="1"[^>]*><f t="array"`))
  })

  it("a range of a link of the old kind reads with `@`, as a sheet's, and is written back of the old kind", () => {
    expect(fromLegacy('[1]S!B1:B3')).toBe("@'[1]S'!B1:B3")
    expect(fromLegacy('ABS([1]S!B1:B3)')).toBe("ABS(@'[1]S'!B1:B3)")
    expect(fromLegacy('SUM([1]S!B1:B3)')).toBe("SUM('[1]S'!B1:B3)")
    expect(fromLegacy('INDEX([1]S!B1:B3,MATCH(1,[1]S!A1:A3,0))')).toBe("INDEX('[1]S'!B1:B3,MATCH(1,'[1]S'!A1:A3,0))")
    expect(fromLegacy('[1]S!B2')).toBe("'[1]S'!B2")
  })
})
