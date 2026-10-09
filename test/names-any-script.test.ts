/**
 * Defined names in letters of any script (`Ставка`, `НДС_2025`), by the rules Excel has for names and tables alike:
 * the check of a name, a book and a session computing them, formulas checked in en-US and ru-RU, names compared in
 * any case, one namespace with tables, and a name kept through structure operations and an xlsx file.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, openSession, restructure, ruRU, toLocal, validateFormula, validateLocalFormula, validateName, type BookInput, type NameInput } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'

/** Sheet S with A1 = 100 and the formulas given, and the names. */
const book = (formulas: Record<string, string>, names: NameInput[]): BookInput => ({
  sheets: [{ name: 'S', cells: { A1: { value: 100 }, ...Object.fromEntries(Object.entries(formulas).map(([a, f]) => [a, { formula: f }])) } }],
  names,
})

const values = (b: BookInput) => {
  const r = compute(b, { display: false })
  return { issues: r.nameIssues, cells: Object.fromEntries(Object.entries(r.sheets[0].cells).map(([a, c]) => [a, c.value])) }
}

describe('defined names in letters of any script', () => {
  it('a name of any script is valid; the rules of a name stay: no digit first, no cell address, no R1C1, no TRUE', () => {
    for (const ok of ['Ставка', 'НДС_2025', 'НДС.2025', '_Ставка', '\\Ставка', 'Ё', 'aБ', 'АБ12', 'К2С3', 'С', 'Сумма', 'Prix_é', 'Τιμή', '税率'])
      expect([ok, validateName(ok)]).toEqual([ok, { ok: true }])
    for (const bad of ['1Ставка', 'Став ка', 'Став-ка', 'AB12', 'R2C3', 'R', 'c', 'TRUE', 'false', 'Я'.repeat(256)]) expect([bad, validateName(bad).ok]).toEqual([bad, false])
  })

  it('a book computes a name of any script, used in any case', () => {
    const v = values(book({ B1: 'Ставка*2', B2: 'SUM(ставка,НДС_2025)', B3: 'СТАВКА+нДс_2025' }, [{ name: 'Ставка', ref: '0.2' }, { name: 'НДС_2025', ref: "'S'!$A$1" }]))
    expect(v.issues).toEqual([])
    expect([v.cells.B1, v.cells.B2, v.cells.B3]).toEqual([0.4, 100.2, 100.2])
  })

  it('a session takes a name of any script through setNames', () => {
    const s = openSession(book({ B1: 'Ставка*2' }, []))
    expect(s.result().sheets[0].cells.B1.value).toEqual({ error: 'NAME' })
    expect(s.setNames([{ name: 'Ставка', ref: '0.2' }]).map((c) => [c.cell, c.value])).toEqual([['B1', 0.4]])
    s.close()
  })

  it('a sheet\'s own name of any script, on a sheet named in Cyrillic', () => {
    const b: BookInput = {
      sheets: [
        { name: 'S', cells: { A1: { formula: "'Лист'!Ставка*2" }, A2: { formula: 'Ставка' } } },
        { name: 'Лист', cells: { A1: { formula: 'ставка' } } },
      ],
      names: [{ name: 'Ставка', ref: '0.2' }, { name: 'Ставка', ref: '0.3', scope: 'Лист' }],
    }
    const r = compute(b, { display: false })
    expect(r.nameIssues).toEqual([])
    expect([r.sheets[0].cells.A1.value, r.sheets[0].cells.A2.value, r.sheets[1].cells.A1.value]).toEqual([0.6, 0.2, 0.3])
  })

  it("a formula spells a name as it is defined, as Excel shows it (tables/names-held any-case): typed in any case, a sheet's own too", () => {
    const b: BookInput = {
      sheets: [
        { name: 'S', cells: { A1: { formula: 'ставка*2' }, A2: { formula: "SUM('Лист'!СТАВКА,1)" } } },
        { name: 'Лист', cells: { A1: { formula: 'ставка' } } },
      ],
      names: [{ name: 'Ставка', ref: '0.2' }, { name: 'Ставка', ref: '0.3', scope: 'Лист' }],
    }
    const r = compute(b, { display: false })
    expect([r.sheets[0].cells.A1.formula, r.sheets[0].cells.A2.formula, r.sheets[1].cells.A1.formula]).toEqual(['Ставка*2', "SUM('Лист'!Ставка,1)", 'Ставка'])
    const s = openSession(b)
    s.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'сТаВкА*10' } }])
    expect(s.result().sheets[0].cells.B1).toMatchObject({ value: 2, formula: 'Ставка*10' })
    s.close()
  })

  it('two names that differ only in case are one name: the second is a duplicate', () => {
    const v = values(book({ B1: 'ставка' }, [{ name: 'Ставка', ref: '1' }, { name: 'СТАВКА', ref: '2' }]))
    expect(v.issues).toEqual([{ name: 'СТАВКА', reason: 'duplicate name' }])
    expect(v.cells.B1).toBe(1)
  })

  it('formulas with such names check in en-US and in ru-RU; a Russian function name is not a name', () => {
    expect(validateFormula('Ставка*2', { names: ['СТАВКА'] })).toMatchObject({ ok: true, canonical: 'Ставка*2', names: ['Ставка'] })
    expect(validateFormula('SUM(Ставка)', { names: ['Ставка'] })).toMatchObject({ ok: true, functions: ['SUM'], names: ['Ставка'] })
    expect(validateFormula('Ставка*2', { names: ['Ставк'] })).toMatchObject({ ok: false, code: 'unknown_name' })
    expect(validateLocalFormula('=Ставка*2', ruRU, { names: ['Ставка'] })).toMatchObject({ ok: true, canonical: 'Ставка*2' })
    expect(validateLocalFormula('=СУММ(Ставка;1,5)', ruRU, { names: ['Ставка'] })).toMatchObject({ ok: true, canonical: 'SUM(Ставка,1.5)', names: ['Ставка'] })
    expect(validateLocalFormula('=СУММ(Сумма;1)', ruRU, { names: ['Сумма'] })).toMatchObject({ ok: true, canonical: 'SUM(Сумма,1)', functions: ['SUM'], names: ['Сумма'] })
    expect(toLocal('SUM(Сумма,НДС_2025)', ruRU)).toEqual({ ok: true, formula: 'СУММ(Сумма;НДС_2025)' })
  })

  it('names and tables share one namespace in any case, Cyrillic too', () => {
    const cells = { A1: { value: 'Кол' }, A2: { value: 1 } }
    const withTable = (name: string, names: NameInput[]): BookInput => ({ sheets: [{ name: 'S', cells, tables: [{ name, range: 'A1:A2', columns: [{ name: 'Кол' }] }] }], names })
    expect(() => compute(withTable('Ставка', [{ name: 'СТАВКА', ref: '1' }]))).toThrow(InputError)
    expect(() => compute(withTable('Ставка', [{ name: 'СТАВКА', ref: '1' }]))).toThrow(/also a defined name/)
    const r = restructure(withTable('Продажи', [{ name: 'Ставка', ref: '1' }]), { op: 'renameTable', table: 'Продажи', name: 'ставка' })
    expect(r).toMatchObject({ ok: false, code: 'table_name' })
    expect(restructure(withTable('Продажи', []), { op: 'renameTable', table: 'Продажи', name: 'АБ12' })).toMatchObject({ ok: true })
  })

  it('a sheet\'s own name follows a copy to another sheet, and a file keeps the names as they are', async () => {
    const b: BookInput = {
      sheets: [
        { name: 'S', cells: { A1: { formula: 'Ставка*2' } } },
        { name: 'T', cells: {} },
      ],
      names: [{ name: 'Ставка', ref: '0.5', scope: 'S' }],
    }
    const copied = restructure(b, { op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'T', range: 'B1' } })
    expect(copied.ok && copied.book.names).toEqual([{ name: 'Ставка', ref: '0.5', scope: 'S' }, { name: 'Ставка', ref: '0.5', scope: 'T' }])
    const { data, issues } = await writeXlsx({ ...b, names: [...b.names!, { name: 'НДС_2025', ref: "'S'!$A$1" }] })
    expect(issues).toEqual([])
    const read = await readXlsx(data)
    expect(read.book.names).toEqual([{ name: 'Ставка', ref: '0.5', scope: 'S' }, { name: 'НДС_2025', ref: "'S'!$A$1" }])
    expect(compute(read.book, { display: false }).sheets[0].cells.A1.value).toBe(1)
  })
})
