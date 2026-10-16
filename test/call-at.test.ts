/**
 * `callAt`: the innermost call a caret stands in and its argument, in en-US and ru-RU text as it is typed: strings, quoted
 * sheets' names, tables' brackets, array constants and groups hold no separator of the call; half typed text gives its
 * call; the name in English for `functionSignature`; the input checked. And `nameAt`, the name of a function being typed.
 */
import { describe, expect, it } from 'vitest'
import { callAt, functionSignature, InputError, nameAt } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { ruRUFunctionText } from '../src/function-text/ru-RU.js'

/** The call at the place of `|` in `text`, as `name index`, or null. */
function at(marked: string, locale?: typeof ruRU): string | null {
  const caret = marked.indexOf('|')
  const r = callAt(marked.replace('|', ''), caret, locale ? { locale } : undefined)
  return r && `${r.fn} ${r.index}`
}

describe('callAt', () => {
  it('finds the call and the argument the caret is in', () => {
    expect(at('=SUM(|')).toBe('SUM 0')
    expect(at('=SUM(A1,|')).toBe('SUM 1')
    expect(at('=SUM(A1,B|2,C3)')).toBe('SUM 1')
    expect(at('=SUMIFS(A:A,B:B,">1",C:C,|')).toBe('SUMIFS 4')
    expect(at('=IF(A1>0,ROUND(A1,|),0)')).toBe('ROUND 1')
    expect(at('=IF(A1>0,ROUND(A1,2),|0)')).toBe('IF 2')
    expect(at('=sum(1,|')).toBe('SUM 1')
  })

  it('answers null outside any call and for text that is no formula', () => {
    expect(at('=SUM|(A1)')).toBeNull()
    expect(at('=SUM(A1)|')).toBeNull()
    expect(at('=A1+|')).toBeNull()
    expect(at('SUM(A1,|')).toBeNull()
    expect(at('=(A1,|')).toBeNull()
  })

  it('counts no separator inside a string, a quoted sheet, a table, an array constant or a group', () => {
    expect(at('=CONCAT("a,b",|')).toBe('CONCAT 1')
    expect(at('=CONCAT("a,|')).toBe('CONCAT 0')
    expect(at("=SUM('a,b'!A1,|")).toBe('SUM 1')
    expect(at('=SUM(T[[#This Row],[x]],|')).toBe('SUM 1')
    expect(at('=SUM({1,2;3,4},|')).toBe('SUM 1')
    expect(at('=SUM((A1,B1),|')).toBe('SUM 1')
    expect(at('=SUM((A1,|')).toBe('SUM 0')
  })

  it("reads ru-RU text by its separator and gives the English name", () => {
    expect(at('=СУММЕСЛИ(A:A;">1";|', ruRU)).toBe('SUMIF 2')
    expect(at('=СУММ(1,5;|', ruRU)).toBe('SUM 1')
    expect(at('=СУММ({1;2:3;4};|', ruRU)).toBe('SUM 1')
    expect(at('=ABS(|', ruRU)).toBe('ABS 0')
    const r = callAt('=СуммЕсли(A:A;1)', 14, { locale: ruRU })!
    expect(r).toEqual({ name: 'СуммЕсли', fn: 'SUMIF', index: 1, open: 9, close: 15 })
    expect(functionSignature(r.fn, { ...ruRU, ...ruRUFunctionText })!.args[r.index].name).toBe('условие')
  })

  it('gives where the call opens and closes', () => {
    expect(callAt('=SUM(A1,(B1+1),C1)+1', 9)).toEqual({ name: 'SUM', fn: 'SUM', index: 1, open: 4, close: 17 })
    expect(callAt('=SUM(A1,"(",', 11)).toMatchObject({ open: 4, close: null })
  })

  it('checks its input', () => {
    expect(() => callAt(1 as unknown as string)).toThrow(InputError)
    expect(() => callAt('=SUM(', 9)).toThrow(InputError)
    expect(() => callAt('=SUM(', 1, { locale: 'de-XX' })).toThrow(InputError)
  })
})

describe('nameAt: the name of a function being typed at the caret', () => {
  it('where an operand may begin: after the first character, an operator, (, the separator, a space', () => {
    expect(nameAt('=SU')).toEqual({ from: 1, to: 3, typed: 'SU' })
    expect(nameAt('=SUM(1,VL')).toEqual({ from: 7, to: 9, typed: 'VL' })
    expect(nameAt('=A1+vl', 6)).toEqual({ from: 4, to: 6, typed: 'vl' })
    expect(nameAt('= SU')).toEqual({ from: 2, to: 4, typed: 'SU' })
    expect(nameAt('-SU')).toEqual({ from: 1, to: 3, typed: 'SU' })
    // the name goes on past the caret
    expect(nameAt('=SUM(1,VLOOK', 9)).toEqual({ from: 7, to: 12, typed: 'VL' })
  })

  it("the locale's list separator, and its names", () => {
    expect(nameAt('=СУММ(1;СУ', undefined, { locale: ruRU })).toEqual({ from: 8, to: 10, typed: 'СУ' })
    // the decimal separator of ru-RU, where no name begins
    expect(nameAt('=СУММ(1,СУ', undefined, { locale: ruRU })).toBeNull()
  })

  it('null where no function can stand', () => {
    for (const t of ['="SU', "='a b'!SU", '=T[SU', '=Sheet1!A', '=$A', '=1', 'SU', '='])
      expect(nameAt(t), t).toBeNull()
    expect(nameAt('=SUM(', 5)).toBeNull()
  })
})
