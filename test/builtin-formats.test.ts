/**
 * Built-in number formats by number beyond the Excel cases (`test/excel/win/display/builtin.json` and its ru-RU twin, as
 * Excel for Windows shows them): on rows and columns, through a session and `displayValue`, and the checks of the input.
 */
import { describe, expect, it } from 'vitest'
import { compute, displayValue, InputError, openSession, type BookInput } from '../src/index.js'

const book = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [{ name: 'S', cells: { A1: { value: 45688.75, format: 14 }, B1: { value: 45688.75 }, C2: { value: -1234.5 } }, rows: { '1': { format: 22 } }, columns: { 'A:B': { width: 20 }, C: { width: 20, format: 38 } } }],
})

describe('built-in formats by number', () => {
  it('show by the locale, on a cell, a row and a column', () => {
    const texts = (b: BookInput) => Object.values(compute(b).sheets[0].cells).map((c) => c.text)
    expect(texts(book())).toEqual(['1/31/2025', '1/31/2025 18:00', '(1,235)'])
    expect(texts(book('ru-RU'))).toEqual(['31.01.2025', '31.01.2025 18:00', '-1 235'])
    expect(compute(book()).sheets[0].cells.A1.format).toBe(14)
  })

  it('a session takes them in edits and layout', () => {
    const s = openSession(book('ru-RU'))
    expect(s.setCells([{ sheet: 'S', cell: 'B1', input: { value: 45688.75, format: 15 } }]).changes.map((c) => c.text)).toEqual(['31.янв.25'])
    expect(s.setLayout([{ sheet: 'S', columns: 'C', set: { format: 7 } }]).changes.map((c) => c.text)).toEqual(['-1 234,50 ₽'])
    s.close()
  })

  it('displayValue takes them', () => {
    expect(displayValue(0.75, { format: 20 })).toEqual({ text: '18:00' })
    expect(displayValue(0.25, { format: 20, locale: 'ru-RU' })).toEqual({ text: '6:00' })
  })

  it('are the numbers 0 to 49', () => {
    for (const bad of [50, -1, 1.5]) expect(() => compute({ sheets: [{ name: 'S', cells: { A1: { value: 1, format: bad } } }] })).toThrow(InputError)
    expect(() => displayValue(1, { format: 99 })).toThrow(/opts\.format: expected a format code or the number of a built-in format, 0 to 49/)
    expect(compute({ sheets: [{ name: 'S', cells: { A1: { value: 1, format: 3 } } }] }).inputHash).not.toBe(compute({ sheets: [{ name: 'S', cells: { A1: { value: 1, format: '3' } } }] }).inputHash)
  })
})
