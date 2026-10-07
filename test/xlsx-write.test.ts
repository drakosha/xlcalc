/**
 * The writer: a book comes back from its file as written (formulas, values, names, layout, locale), formulas
 * are dynamic array formulas in the file's spelling with their values and spills, the values mode writes values
 * only, errors Excel keeps as rich values (#SPILL!, #CALC!) go through value metadata, and what a file cannot hold
 * is reported. The same book gives the same bytes.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, ComputeResult } from '../src/contract.js'
import { compute, InputError } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { codes } from './xlsx-fixtures.js'

const clock = { now: '2026-10-07T12:00:00Z', seed: 7 }

const part = async (data: Uint8Array, name: string) => (await openArchive(data).text(name))!

const book: BookInput = {
  sheets: [
    {
      name: 'Data',
      cells: {
        A1: { value: 1 },
        A2: { value: 2.5 },
        A3: { value: 'текст' },
        A4: { value: true },
        A5: { value: { date: '2025-01-31' } },
        B1: { formula: 'SEQUENCE(3)' },
        C1: { formula: 'B1#*2' },
        D1: { formula: 'XLOOKUP(2.5,A1:A2,A1:A2)' },
        E1: { formula: 'LET(x,A1,x+1)' },
        F1: { formula: 'FILTER(A1:A2,A1:A2>5)' },
        G1: { formula: 'SEQUENCE(2)' },
        G2: { value: 'blocks' },
        H1: { formula: "SUM('It''s'!A1:A2)" },
      },
    },
    { name: "It's", cells: { A1: { value: 10 }, A2: { formula: 'rate*2' } } },
  ],
  names: [{ name: 'rate', ref: "'Data'!$A$1" }, { name: 'here', ref: "'It''s'!$A$1", scope: "It's" }],
  locale: 'ru-RU',
}

describe('xlsx writer', () => {
  it('writes formulas as dynamic array formulas in the file spelling, with their values and spills', async () => {
    const result = compute(book, { clock })
    const { data, issues } = await writeXlsx(book, { result })
    expect(issues).toEqual([])
    const sheet = await part(data, 'xl/worksheets/sheet1.xml')
    expect(sheet).toContain('<c r="B1" cm="1"><f t="array" ref="B1:B3">_xlfn.SEQUENCE(3)</f><v>1</v></c>')
    expect(sheet).toContain('<c r="B2"><v>2</v></c>')
    expect(sheet).toContain('<f t="array" ref="C1:C3">_xlfn.ANCHORARRAY(B1)*2</f>')
    expect(sheet).toContain('<f t="array" ref="D1">_xlfn.XLOOKUP(2.5,A1:A2,A1:A2)</f>')
    expect(sheet).toContain('<f t="array" ref="E1">_xlfn.LET(_xlpm.x,A1,_xlpm.x+1)</f>')
    expect(sheet).toMatch(/<c r="F1" t="e" vm="\d" cm="1"><f t="array" ref="F1">_xlfn._xlws.FILTER\(A1:A2,A1:A2&gt;5\)<\/f><v>#VALUE!<\/v><\/c>/)
    expect(sheet).toMatch(/<c r="G1" t="e" vm="\d" cm="1">/)
    expect(await part(data, 'xl/workbook.xml')).toContain('<definedName name="rate">\'Data\'!$A$1</definedName><definedName name="here" localSheetId="1">\'It\'\'s\'!$A$1</definedName></definedNames><calcPr fullCalcOnLoad="1"/>')
    expect(await part(data, 'xl/richData/rdrichvalue.xml')).toContain('<rv s="0"><v>13</v></rv><rv s="0"><v>8</v></rv>')
    expect(await part(data, 'docProps/custom.xml')).toContain('name="xlcalc.locale"><vt:lpwstr>ru-RU</vt:lpwstr>')
  })

  it('the book comes back from its file, and the stored values are the result', async () => {
    const result = compute(book, { clock })
    const r = await readXlsx((await writeXlsx(book, { result })).data)
    expect(r.book).toEqual({
      ...book,
      sheets: [{ ...book.sheets[0], cells: { ...book.sheets[0].cells, A5: { value: 45688 }, H1: { formula: "SUM('It''s'!A1:A2)" } } }, book.sheets[1]],
    })
    expect(r.issues).toEqual([])
    const values = (s: ComputeResult['sheets'][number]) => Object.fromEntries(Object.entries(s.cells).flatMap(([a, c]) => (c.formula !== undefined || c.spillOf !== undefined ? [[a, c.value]] : [])))
    expect(r.cached).toEqual(result.sheets.map((s) => ({ name: s.name, cells: values(s) })))
  })

  it('values only: formulas give way to their values, spills too; errors are stored as errors', async () => {
    const result = compute(book, { clock })
    const { data } = await writeXlsx(book, { result, mode: 'values' })
    const r = await readXlsx(data)
    expect(r.book.sheets[0].cells).toMatchObject({ B1: { value: 1 }, B3: { value: 3 }, C3: { value: 6 }, D1: { value: 2.5 }, G1: { formula: '#SPILL!' }, H1: { value: 12 } })
    expect(r.book.sheets[0].cells.F1).toBeUndefined()
    expect(await openArchive(data).text('xl/metadata.xml')).not.toContain('XLDAPR')
    expect(codes(r.issues)).toEqual(['errorValue F1', 'errorConstant G1'])
  })

  it('formulas without a result: the file asks Excel to compute them', async () => {
    const { data } = await writeXlsx({ sheets: [{ name: 'S', cells: { A1: { formula: 'SEQUENCE(2)' }, B1: { value: 1 } } }] })
    const sheet = await part(data, 'xl/worksheets/sheet1.xml')
    expect(sheet).toContain('<c r="A1" cm="1"><f t="array" ref="A1">_xlfn.SEQUENCE(2)</f></c>')
    expect(await part(data, 'xl/workbook.xml')).toContain('<calcPr fullCalcOnLoad="1"/>')
  })

  it('the same book gives the same bytes', async () => {
    const result = compute(book, { clock })
    expect((await writeXlsx(book, { result })).data).toEqual((await writeXlsx(book, { result })).data)
  })

  it('reports what a file cannot hold: meta, a circular reference, text past 32,767 characters, formulas and names it cannot write', async () => {
    const odd: BookInput = {
      sheets: [
        {
          name: 'S',
          cells: {
            A1: { value: 1, meta: { row: 7 } },
            A2: { formula: 'A2+1' },
            A3: { value: 'x'.repeat(40_000) },
            A4: { formula: "'Nope'!A1" },
            A5: { formula: 'SUM(' },
            A6: { formula: '[Book.xlsx]S!A1' },
          },
        },
      ],
      names: [{ name: 'A1', ref: "'S'!$A$1" }, { name: 'ok', ref: "'S'!$A$1" }, { name: 'ok', ref: "'S'!$A$2" }, { name: 'gone', ref: "'S'!$A$1", scope: 'Nope' }, { name: 'bad', ref: 'SUM(' }],
    }
    const result = compute(odd, { clock })
    const { data, issues } = await writeXlsx(odd, { result })
    expect(codes(issues)).toEqual(['meta A1', 'errorValue A2', 'textLength A3', 'formula A4 A5 A6', 'name', 'name', 'name', 'name'])
    expect(issues.filter((i) => i.code === 'name').map((i) => `${i.name}: ${i.message}`)).toEqual([
      'A1: a name Excel does not take is left out of the file',
      'ok: a name given twice is left out of the file',
      'gone: a name of a sheet the book does not have is left out of the file',
      'bad: a name whose reference does not parse is left out of the file',
    ])
    const r = await readXlsx(data)
    expect(r.book.sheets[0].cells).toMatchObject({ A2: { formula: 'A2+1' }, A3: { value: 'x'.repeat(32_767) }, A4: { formula: '#REF!' }, A5: { formula: '#NULL!' } })
    expect(r.cached[0].cells.A2).toBe(0)
  })

  it('checks its input: the book, the mode, a result of another book', async () => {
    await expect(writeXlsx({ sheets: [{ name: 'a:b', cells: {} }] })).rejects.toThrow(InputError)
    await expect(writeXlsx(book, { mode: 'values' })).rejects.toThrow(/values mode takes the values from the result/)
    await expect(writeXlsx(book, { mode: 'other' as 'values' })).rejects.toThrow(/expected "formulas" or "values"/)
    const other = compute({ sheets: [{ name: 'X', cells: {} }] }, { clock })
    await expect(writeXlsx(book, { result: other })).rejects.toThrow(/its sheets differ from the book's/)
  })
})
