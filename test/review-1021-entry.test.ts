/**
 * Reproductions from the review of 21.10 on typed entry: the text `entryText` writes for a percentage cell is the
 * in-place editor's text, and `parseEntry` with `inPlace` reads it back to the value, as `setCellText` does with the
 * option. Parsing runs no formula, so there is no execution mode to vary.
 */
import { describe, expect, it } from 'vitest'
import { entryText, parseEntry } from '../src/index.js'
import { parseEntry as bareParseEntry } from '../src/typed-entry.js'
import { ruRU } from '../src/locales/ru-RU.js'

describe('parseEntry edited in place', () => {
  const formats = [9, 10, '0.0%'] as const
  // a fixed series of values over many magnitudes
  const values = Array.from({ length: 200 }, (_, i) => Number(((((i * 7919) % 1000) / 997) * 10 ** ((i % 13) - 6)).toPrecision(15)))

  it.each([['en-US', undefined], ['ru-RU', ruRU]] as const)('reads back a percentage cell entryText writes, in %s', (_, locale) => {
    for (const numberFormat of formats) {
      for (const v of [0, ...values]) {
        const text = entryText(v, { numberFormat, ...(locale ? { locale } : {}) })
        const back = parseEntry(text, { cell: { numberFormat }, inPlace: true, ...(locale ? { locale } : {}) })
        expect(back.ok && back.input?.value, text).toBe(v)
      }
    }
  })

  it('reads a negative percentage as Excel edits it in place, a formula (the set ui/edit-held, sign-5)', () => {
    const text = entryText(-0.05, { numberFormat: 9 })
    expect(text).toBe('-5%')
    expect(parseEntry(text, { cell: { numberFormat: 9 }, inPlace: true })).toEqual({ ok: true, input: { formula: '-5%%', numberFormat: 9 } })
  })

  it('parses as typed over the cell without the option, and in the package without the engine alike', () => {
    expect(parseEntry('6%', { cell: { numberFormat: 9 } })).toEqual({ ok: true, input: { value: '6%%', numberFormat: 9 } })
    expect(parseEntry('6%', { cell: { numberFormat: 9 }, inPlace: false })).toEqual({ ok: true, input: { value: '6%%', numberFormat: 9 } })
    expect(parseEntry('6', { cell: { numberFormat: 9 }, inPlace: true })).toEqual({ ok: true, input: { value: 6, numberFormat: 9 } })
    expect(bareParseEntry('6%', { cell: { numberFormat: 9 }, inPlace: true })).toEqual({ ok: true, input: { value: 0.06, numberFormat: 9 } })
    expect(() => parseEntry('6', { inPlace: 'yes' as unknown as boolean })).toThrow(/opts\.inPlace/)
  })
})
