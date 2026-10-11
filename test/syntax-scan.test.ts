/**
 * The pieces the tokenizer, the locale lexer and the reader of a file's formulas scan alike (`src/syntax/scan.ts`), on
 * the edges where their copies used to agree: a run left open, a doubled quote at the end, an escape before the end.
 */
import { describe, expect, it } from 'vitest'
import { BOOK_SHEET, bracketEnd, callFollows, closingQuote, isSpace, numberSource } from '../src/syntax/scan.js'

describe('scanning formula text', () => {
  it('a quoted run ends at its closing quote, a doubled quote inside; -1 when it is not closed', () => {
    expect(['"a""b"+1', '"a""', '"', '""""', "'S''x'!A1", "'S"].map((s) => closingQuote(s, 0))).toEqual([5, -1, -1, 3, 5, -1])
    expect(closingQuote('x="a"', 2)).toBe(4)
  })

  it('a bracketed run ends past its last bracket, nested brackets and escapes inside; the text’s length when it is not closed', () => {
    const s = 'Sales[[#Data],[Qty]]!'
    expect(bracketEnd(s, 5)).toBe(s.length - 1)
    expect(["[a']b]+1", '[[x]', "[a'", '[1]S!A1'].map((r) => bracketEnd(r, 0))).toEqual([6, 4, 3, 3])
  })

  it('what follows a bracketed run makes it a reference to another workbook', () => {
    expect(['S!A1', '!name', "'S x'!A1", '+1', ''].map((t) => BOOK_SHEET.test(t))).toEqual([true, true, true, false, false])
  })

  it('a call follows a word past spaces only; whitespace is four characters', () => {
    expect(['SUM(', 'SUM  (', 'SUM\t(', 'SUM'].map((s) => callFollows(s, 3))).toEqual([true, true, false, false])
    expect([' ', '\t', '\n', '\r', ' ', undefined].map(isSpace)).toEqual([true, true, true, true, false, false])
  })

  it('a number with its decimal separator', () => {
    const num = new RegExp(`^${numberSource(',')}$`)
    expect(['1,5', ',5', '5,', '1,5E-07', '1.5', 'E5'].map((s) => num.test(s))).toEqual([true, true, true, true, false, false])
  })
})
