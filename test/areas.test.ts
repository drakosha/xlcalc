/**
 * The syntax of `@`, unions and references across sheets beyond Excel's answers (`test/excel/cases/areas.json`,
 * `test/excel/spill/at*.json`): parsing and canonical text, checks, ru-RU text, and how a reference across sheets
 * reads the sheets between its ends in the book's order.
 */
import { describe, expect, it } from 'vitest'
import { compute, editStructure, toInvariant, toLocal, validateFormula, type BookInput } from '../src/index.js'
import { ruRU } from './locales.js'

const canon = (f: string) => {
  const r = validateFormula(f)
  return r.ok ? r.canonical : `${r.code}: ${r.message}`
}

describe('the syntax of @, unions and references across sheets', () => {
  it('canonical text: @ before its operand, a union in its parentheses, the two sheets quoted together', () => {
    expect(canon('@ A1:A3 * 2')).toBe('@A1:A3*2')
    expect(canon('@(A1:A3*2)+@sum(A1)')).toBe('@(A1:A3*2)+@SUM(A1)')
    expect(canon('SUM((A1:A3, C1), 5)')).toBe('SUM((A1:A3,C1),5)')
    expect(canon('SUM(Jan:Mar!A1:B2)')).toBe("SUM('Jan:Mar'!A1:B2)")
    expect(canon("SUM('Jan 1:Mar 1'!A:A)")).toBe("SUM('Jan 1:Mar 1'!A:A)")
    // as Excel saved a reference across sheets inside a union
    expect(canon("COUNT((Jan:'Mar'!A1,D!B1))")).toBe("COUNT(('Jan:Mar'!A1,'D'!B1))")
    expect(canon('Data!rate*2')).toBe("'Data'!rate*2")
  })

  it('refuses a union outside parentheses, a spill reference across sheets, a function after a sheet', () => {
    expect(canon('A1,B1')).toMatch(/^syntax: comma outside/)
    expect(canon("'Jan:Mar'!A1#")).toMatch(/^syntax: a spill reference/)
    expect(canon('S!SUM(1)')).toMatch(/^syntax: reference expected after sheet name/)
  })

  it('checks both ends of a reference across sheets and the sheet of a qualified name', () => {
    expect(validateFormula("SUM('Jan:Mar'!A1)", { sheets: ['Jan', 'Feb'] })).toMatchObject({ ok: false, code: 'unknown_sheet' })
    expect(validateFormula("SUM('Jan:Mar'!A1)", { sheets: ['Jan', 'Mar'] })).toMatchObject({ ok: true, sheets: ['Jan', 'Mar'] })
  })

  it('ru-RU text: a union joined by ";", @ and the sheets as Excel shows them', () => {
    const local = toLocal("SUM((A1:A3,C1))+@B1:B3+SUM('Jan:Mar'!A1)+1.5", ruRU)
    expect(local).toEqual({ ok: true, formula: 'СУММ((A1:A3;C1))+@B1:B3+СУММ(Jan:Mar!A1)+1,5' })
    expect(toInvariant(local.ok ? local.formula : '', ruRU)).toEqual({ ok: true, formula: 'SUM((A1:A3,C1))+@B1:B3+SUM(Jan:Mar!A1)+1.5' })
  })

  it('a reference across sheets reads the sheets between its ends in the book\'s order, either way round', () => {
    const sheet = (name: string, v: number) => ({ name, cells: { A1: { value: v } } })
    const book: BookInput = { sheets: [sheet('Jan', 1), sheet('Feb', 10), sheet('Mar', 100), { name: 'S', cells: { A1: { formula: "SUM('Jan:Mar'!A1)" }, A2: { formula: "SUM('Mar:Feb'!A1)" } } }] }
    const cells = (b: BookInput) => compute(b, { display: false }).sheets.find((s) => s.name === 'S')!.cells
    expect([cells(book).A1.value, cells(book).A2.value]).toEqual([111, 110])
    // a sheet moved out from between the ends is not read
    const moved = editStructure(book, { op: 'moveSheet', sheet: 'Feb', before: 'S' })
    expect(moved.ok && cells(moved.book).A1.value).toBe(101)
    const inserted = editStructure(book, { op: 'insertSheet', name: 'New', before: 'Feb' })
    expect(inserted.ok && inserted.book.sheets.map((s) => s.name)).toEqual(['Jan', 'New', 'Feb', 'Mar', 'S'])
  })

  it('refuses a new or moved sheet whose name is taken or breaks Excel\'s rules', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: {} }, { name: 'T', cells: {} }] }
    expect(editStructure(book, { op: 'insertSheet', name: 't' })).toMatchObject({ ok: false, code: 'sheet_name' })
    expect(editStructure(book, { op: 'insertSheet', name: 'a/b' })).toMatchObject({ ok: false, code: 'sheet_name' })
    expect(editStructure(book, { op: 'moveSheet', sheet: 'S' })).toMatchObject({ ok: true, book: { sheets: [{ name: 'T' }, { name: 'S' }] } })
  })
})
