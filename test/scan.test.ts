/**
 * `scanReferences`: the references of formula text with their places, read by the engine's tokenizer, stored text and
 * text typed in ru-RU alike; a formula half typed never throws and gives the references the whole formula gives before
 * the place it is typed to, and a token that does not read is skipped, the text read on past it.
 */
import { describe, expect, it } from 'vitest'
import type { FormulaReference } from '../src/contract.js'
import { InputError, scanReferences, toLocal } from '../src/index.js'
import { ruRU } from './locales.js'

const brief = (refs: FormulaReference[]) => refs.map((r) => `${r.at}-${r.end} ${r.text}`)

describe('scanReferences', () => {
  it('cells, whole columns and rows, sheets, references across sheets, spills, with their rectangles', () => {
    const refs = scanReferences("=SUM(B5:a1, 'Q1 data'!$C$1, Jan:Mar!A1, b:B, 9:2) + E2# + 'O''Brien'!x1")
    expect(brief(refs)).toEqual(['5-10 B5:a1', "12-26 'Q1 data'!$C$1", '28-38 Jan:Mar!A1', '40-43 b:B', '45-48 9:2', '52-55 E2#', "58-71 'O''Brien'!x1"])
    expect(refs[0]).toEqual({ kind: 'cells', at: 5, end: 10, text: 'B5:a1', top: 1, left: 1, bottom: 5, right: 2 })
    expect(refs[1]).toMatchObject({ sheet: 'Q1 data', top: 1, left: 3 })
    expect(refs[2]).toMatchObject({ sheet: 'Jan', sheetLast: 'Mar' })
    expect(refs[3]).toMatchObject({ kind: 'columns', top: 1, left: 2, bottom: 1_048_576, right: 2 })
    expect(refs[4]).toMatchObject({ kind: 'rows', top: 2, left: 1, bottom: 9, right: 16_384 })
    expect(refs[5]).toMatchObject({ kind: 'cells', spill: true, top: 2, left: 5 })
    expect(refs[6]).toMatchObject({ sheet: "O'Brien", left: 24 })
  })

  it('skips strings, function names and defined names; reads stored text without `=`', () => {
    expect(brief(scanReferences('IF(A1>0,"B2 C3",LOG10(rate)+A1:A2 B1:B3)'))).toEqual(['3-5 A1', '28-33 A1:A2', '34-39 B1:B3'])
  })

  it('structured references with their table, columns and specifiers', () => {
    expect(scanReferences('=SUM(Sales[Qty])+[@Price]+Sales[[#Headers],[Qty]:[Price]]')).toEqual([
      { kind: 'table', at: 5, end: 15, text: 'Sales[Qty]', table: 'Sales', columns: ['Qty'] },
      { kind: 'table', at: 17, end: 25, text: '[@Price]', columns: ['Price'], items: ['#This Row'] },
      { kind: 'table', at: 26, end: 57, text: 'Sales[[#Headers],[Qty]:[Price]]', table: 'Sales', columns: ['Qty', 'Price'], items: ['#Headers'] },
    ])
  })

  it('text typed in a locale, with places in the typed text', () => {
    const text = '=СУММ(Лист1!A1:B3; 1,5; "C1"; Sales[[#Заголовки];[Qty]])'
    expect(brief(scanReferences(text, { locale: 'ru-RU' }))).toEqual(['6-17 Лист1!A1:B3', '30-55 Sales[[#Заголовки];[Qty]]'])
    // the same references as the stored text gives, at their places in it
    const stored = 'SUM(Лист1!A1:B3,1.5,"C1",Sales[[#Headers],[Qty]])'
    expect(scanReferences(stored).map((r) => r.text)).toEqual(['Лист1!A1:B3', 'Sales[[#Headers],[Qty]]'])
    expect(toLocal(stored, ruRU).ok).toBe(true)
  })

  it('a formula half typed: the references before the place it stops reading, as the whole formula gives them', () => {
    for (const [whole, locale] of [
      ["=SUM(A1:B3,'Q1 data'!C1,Sales[Qty],\"x\",Jan:Mar!$B$2)+LOG(E2#)*2:9", undefined],
      ['=ЕСЛИ(A1>0;СУММ(B:B;Лист1!C1:C9);"нет")+[@Qty]', 'ru-RU'],
    ] as const) {
      const all = scanReferences(whole, { ...(locale ? { locale } : {}) })
      for (let n = 0; n <= whole.length; n++) {
        const typed = whole.slice(0, n)
        const got = scanReferences(typed, { ...(locale ? { locale } : {}) })
        expect(got.every((r) => r.end <= n && typed.slice(r.at, r.end) === r.text), typed).toBe(true)
        // a reference the next character shows complete is found as the whole formula has it
        for (const r of all.filter((x) => x.end < n)) expect(got, `${typed}: ${r.text}`).toContainEqual(r)
      }
    }
  })

  it('skips a token that does not read and reads on past it', () => {
    expect(brief(scanReferences('=A1+)B2'))).toEqual(['1-3 A1', '5-7 B2'])
    expect(brief(scanReferences('=SUM(A1,'))).toEqual(['5-7 A1'])
    // a character the formula does not take, alone
    expect(brief(scanReferences('=A1~B2'))).toEqual(['1-3 A1', '4-6 B2'])
    // a word: a reference off the sheet, an error literal half typed, a quoted name that is no sheet
    expect(brief(scanReferences('=A1+$XFE$1+B2'))).toEqual(['1-3 A1', '11-13 B2'])
    expect(brief(scanReferences('=$A$0:B2+C1'))).toEqual(['9-11 C1'])
    expect(brief(scanReferences('=SUM(A1:B2,#N/+C1)'))).toEqual(['5-10 A1:B2', '15-17 C1'])
    expect(brief(scanReferences("='Q1 data'+B2"))).toEqual(['11-13 B2'])
    // a sheet with nothing after it, a structured reference Excel does not take, a reference to another book
    expect(brief(scanReferences('=SUM(Sheet1!)+B2'))).toEqual(['14-16 B2'])
    expect(brief(scanReferences('=Sales[#Foo]+B2'))).toEqual(['13-15 B2'])
    expect(brief(scanReferences('=[1]S!A1+B2'))).toEqual(['9-11 B2'])
    // a cell's shape past the sheet is a name, as Excel reads it
    expect(brief(scanReferences('=A1+XFE1+B2'))).toEqual(['1-3 A1', '9-11 B2'])
  })

  it('an open string, quoted name or bracket takes the rest of the text; a reference after a closed string reads', () => {
    expect(scanReferences('="A1')).toEqual([])
    expect(brief(scanReferences('=A1+"open B2'))).toEqual(['1-3 A1'])
    expect(brief(scanReferences("=A1+'Q1 data!B2"))).toEqual(['1-3 A1'])
    expect(brief(scanReferences("=A1+'Sheet 2'!"))).toEqual(['1-3 A1'])
    expect(brief(scanReferences('=A1+Sales[Qty+B2'))).toEqual(['1-3 A1'])
    expect(brief(scanReferences('="x"&A1+)B2'))).toEqual(['5-7 A1', '9-11 B2'])
    expect(brief(scanReferences('=A1+#X"B2" C3'))).toEqual(['1-3 A1', '11-13 C3'])
    expect(scanReferences('')).toEqual([])
  })

  it('in a locale: a piece that does not translate is skipped, positions and array constants kept', () => {
    expect(brief(scanReferences('=СУММ(A1;1.5)+B2', { locale: ruRU }))).toEqual(['6-8 A1', '14-16 B2'])
    expect(brief(scanReferences('=A1+~B2', { locale: ruRU }))).toEqual(['1-3 A1', '5-7 B2'])
    // `2:3` stays an array's rows after the piece skipped in it, not rows of the sheet
    expect(brief(scanReferences('={1.5;2:3}+B2', { locale: ruRU }))).toEqual(['11-13 B2'])
    expect(brief(scanReferences('=СУММ(A1;"x)+B2', { locale: ruRU }))).toEqual(['6-8 A1'])
    expect(brief(scanReferences('=ЕСЛИ(A1>0;B2;C3', { locale: ruRU }))).toEqual(['6-8 A1', '11-13 B2', '14-16 C3'])
  })

  it('wrong input is an InputError', () => {
    expect(() => scanReferences(1 as unknown as string)).toThrow(InputError)
    expect(() => scanReferences('A1', { locale: 'xx' })).toThrow(/unknown locale/)
    expect(() => scanReferences('A1', { lang: 'ru-RU' } as never)).toThrow(/unknown field "lang"/)
  })
})
