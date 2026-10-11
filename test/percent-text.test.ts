/**
 * Percent text read as Excel reads it: the decimal text is scaled by its exponent, so `1.1%` is the number 0.011 (as
 * `Number("1.1e-2")`), not 1.1/100 (0.011000000000000001). Typed into a cell (`parseEntry`, the entry sets), and in VALUE
 * and operators (the set `pct-text`), in both locales.
 */
import { describe, expect, it } from 'vitest'
import { compute, parseEntry, type BookInput } from '../src/index.js'
import './locales.js'

const value = (text: string, locale?: string) => {
  const r = parseEntry(text, { locale })
  return r.ok ? r.input?.value : r
}

describe('percent typed into a cell', () => {
  it('keeps the decimal text of the number', () => {
    expect(value('1.1%')).toBe(0.011)
    expect(value('0.7%')).toBe(0.007)
    expect(value('1.255%')).toBe(0.01255)
    expect(value('-0.7%')).toBe(-0.007)
    expect(value('(1.1%)')).toBe(-0.011)
    expect(value('%1.1')).toBe(0.011)
    expect(value('1,000.7%')).toBe(10.007)
    expect(value('1,1%', 'ru-RU')).toBe(0.011)
    expect(value('0,7%', 'ru-RU')).toBe(0.007)
    expect(value('99,9%', 'ru-RU')).toBe(0.999)
  })
})

describe.each(['closures', 'eval'] as const)('percent text in formulas (%s)', (execution) => {
  const run = (formula: string, locale?: string) => {
    const book: BookInput = { ...(locale ? { locale } : {}), sheets: [{ name: 'S', cells: { A1: { formula } } }] }
    return compute(book, { execution, clock: { now: '2026-10-10T12:00:00Z', seed: 1 } }).sheets[0].cells.A1.value
  }
  it('VALUE and operators read percent text as its decimal text', () => {
    expect(run('(VALUE("1.1%")-0.011)*1E+20')).toBe(0)
    expect(run('("0.7%"*1-0.007)*1E+20')).toBe(0)
    expect(run('(-"(1.255%)"-0.01255)*1E+20')).toBe(0)
    expect(run('(VALUE("1,1%")-0.011)*1E+20', 'ru-RU')).toBe(0)
    expect(run('("99,9%"+0-0.999)*1E+20', 'ru-RU')).toBe(0)
  })
})
