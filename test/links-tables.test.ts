/**
 * Structured references to a table of another workbook: `[1]!Sales[[#This Row],[Qty]]`, as Excel writes them, are
 * `#REF!` while the book is closed and keep their text through checks, Excel's display, typing and files; written without
 * `!` (`[1]Sales[Qty]`) the reference is lost as Excel reads it, `#REF!`. Excel's answers: `test/links-excel.test.ts`
 * (the set `tables-held`).
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { compute, toInvariant, toLocal } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'

const externalBooks = [{ target: 'C:\\data\\Book2.xlsx', sheets: [{ name: 'Лист1', cells: { A1: 1 } }] }]
const formulas = { A2: '[1]!Таблица[[#This Row],[Столбец]]', A3: 'SUM([1]!Таблица[Столбец])*2', A4: 'IFERROR([1]!T[#All],"e")', A5: '[1]Таблица[Столбец]' }
const book = { sheets: [{ name: 'S', cells: Object.fromEntries(Object.entries(formulas).map(([a, f]) => [a, { formula: f }])) }], externalBooks }

describe('tables of linked books', () => {
  it('are #REF! and keep their text; without ! the reference is lost', () => {
    const cells = compute(book).sheets[0].cells
    expect([cells.A2.value, cells.A3.value, cells.A4.value, cells.A5.value]).toEqual([{ error: 'REF' }, { error: 'REF' }, 'e', { error: 'REF' }])
    expect([cells.A2.formula, cells.A3.formula, cells.A4.formula, cells.A5.formula]).toEqual([formulas.A2, formulas.A3, formulas.A4, '#REF!'])
  })

  it('show as Excel shows them and read back from what a person types', () => {
    const shown = toLocal(formulas.A2, 'en-US', { externalBooks })
    expect(shown).toEqual({ ok: true, formula: "'C:\\data\\Book2.xlsx'!Таблица[@Столбец]" })
    expect(toInvariant(shown.ok ? shown.formula : '', 'en-US', { externalBooks })).toEqual({ ok: true, formula: formulas.A2 })
  })

  it('go into a file and back as Excel writes them', async () => {
    const { data } = await writeXlsx({ book })
    const sheet = (await openArchive(data).text('xl/worksheets/sheet1.xml'))!
    expect(sheet).toContain('>[1]!Таблица[[#This Row],[Столбец]]</f>')
    const read = await readXlsx(data)
    expect(read.book.sheets[0].cells.A2.formula).toBe(formulas.A2)
  })

  it("read from the file Excel saved, with Excel's value", async () => {
    const saved = await readXlsx(readFileSync(new URL('./excel/links/tables-held/bang.xlsx', import.meta.url)))
    const cells = compute(saved.book).sheets[0].cells
    expect([cells.A2.formula, cells.A2.value]).toEqual([formulas.A2, { error: 'REF' }])
  })
})
