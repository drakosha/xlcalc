/**
 * The xlsx subpath on names of any formula, the @ operator, references across sheets and array formulas of the old
 * kind: the writer writes them as Excel does (a name's formula as it is, `@x` as `_xlfn.SINGLE(x)`, an old-kind array
 * over its area without the dynamic array metadata), and the reader reads them back into the same book.
 */
import { describe, expect, it } from 'vitest'
import { storedOf, storedValues } from './xlsx-fixtures.js'
import type { BookInput } from '../src/contract.js'
import { compute } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'

const part = async (data: Uint8Array, name: string) => (await openArchive(data).text(name))!

const book: BookInput = {
  sheets: [
    { name: 'Jan', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 } } },
    {
      name: 'Feb',
      cells: {
        A1: { value: 10 },
        A2: { value: 20 },
        A3: { value: 30 },
        B1: { formula: 'A1:A3*2', array: 'B1:B3' },
        C1: { formula: 'SUM(A1:A3*A1:A3)', array: 'C1' },
        D2: { formula: '@A1:A3+rate' },
        E1: { formula: "SUM('Jan:Feb'!A1:A3)" },
        F1: { formula: 'SUM(both)+SUM(twice)+left' },
      },
    },
  ],
  names: [
    { name: 'rate', ref: '0.5' },
    { name: 'twice', ref: "'Jan'!$A$1:$A$3*2" },
    { name: 'both', ref: "'Jan'!$A$1:$A$2,'Jan'!$A$3" },
    { name: 'left', ref: "'Feb'!XFD1" },
  ],
}

describe('xlsx: names of any formula, @, references across sheets, old-kind arrays', () => {
  it('writes them as Excel stores them', async () => {
    const { data, issues } = await writeXlsx({ book, result: compute(book) })
    expect(issues).toEqual([])
    const sheet = await part(data, 'xl/worksheets/sheet2.xml')
    expect(sheet).toContain('<c r="B1"><f t="array" ref="B1:B3">A1:A3*2</f><v>20</v></c><c r="C1"><f t="array" ref="C1">SUM(A1:A3*A1:A3)</f><v>1400</v></c>')
    expect(sheet).toContain('<c r="D2" cm="1"><f t="array" ref="D2">_xlfn.SINGLE(A1:A3)+rate</f><v>20.5</v></c>')
    expect(sheet).toContain(`<f t="array" ref="E1">SUM('Jan:Feb'!A1:A3)</f><v>66</v>`)
    expect(await part(data, 'xl/workbook.xml')).toContain(
      `<definedName name="rate">0.5</definedName><definedName name="twice">'Jan'!$A$1:$A$3*2</definedName><definedName name="both">'Jan'!$A$1:$A$2,'Jan'!$A$3</definedName><definedName name="left">'Feb'!XFD1</definedName>`,
    )
  })

  it('reads them back into the same book, with the values computed', async () => {
    const result = compute(book)
    const r = await readXlsx((await writeXlsx({ book, result })).data)
    expect(r.book).toEqual(book)
    expect(r.issues).toEqual([])
    expect(r.cached).toEqual(storedOf(result))
    // `left` used in F1 is E1, the cell to its left
    expect(storedValues(r, 1)).toMatchObject({ B2: 40, F1: 6 + 12 + 66 })
  })
})
