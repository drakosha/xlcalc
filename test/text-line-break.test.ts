/**
 * Text with a line break where a number is read (`A1*1` over a cell with an address on two lines): it is no number,
 * so the formula is `#VALUE!` and the rest of the book computes, in either locale, in a book and in a session.
 */
import { describe, expect, it } from 'vitest'
import { compute, openSession, type BookInput } from '../src/index.js'
import './locales.js'

const TEXTS = ['x\ny', '1\n2', 'x\r\ny', 'x y', '$1\n']
const book = (locale: string, text: string): BookInput => ({ locale, sheets: [{ name: 'S', cells: { A1: { value: text }, B1: { formula: 'A1*1' }, B2: { formula: 'IFERROR(A1+1,"no")' }, B3: { formula: '1+1' } } }] })

describe.each(['closures', 'eval'] as const)('text with a line break read as a number, %s', (execution) => {
  it.each(['en-US', 'ru-RU'])('%s: #VALUE! in the cell, the rest computed', (locale) => {
    for (const text of TEXTS) {
      const v = compute(book(locale, text), { execution }).sheets[0].cells
      expect([v.B1.value, v.B2.value, v.B3.value], JSON.stringify(text)).toEqual([{ error: 'VALUE' }, 'no', 2])
    }
  })

  it.each(['en-US', 'ru-RU'])('%s: a session edit that brings such text in', (locale) => {
    const s = openSession(book(locale, '2'), { execution })
    const changes = s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 'x\ny' } }]).map((c) => [c.cell, c.value])
    expect(changes.sort((a, b) => String(a[0]).localeCompare(String(b[0])))).toEqual([['A1', 'x\ny'], ['B1', { error: 'VALUE' }], ['B2', 'no']])
    s.close()
  })
})
