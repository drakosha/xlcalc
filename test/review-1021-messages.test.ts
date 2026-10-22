/**
 * Messages from the review of 21.10: a value in a locale's message is written as a formula of the locale writes it, a
 * quote in text doubled and an exponent in capitals, as ru-RU's explain shows values. With closures and eval.
 */
import { describe, expect, it } from 'vitest'
import { createSession, type CellValue } from '../src/index.js'
import { ownText } from '../src/locale/messages.js'
import { toLocal } from '../src/locale/index.js'
import { ruRU } from './locales.js'

const ru = (value: Exclude<CellValue, null>) => ownText('ru-RU', { key: 'explainValue', value }, (f) => (toLocal(f, ruRU) as { formula: string }).formula)

describe('a value in a message of ru-RU', () => {
  it('is written as a formula of the locale writes it', () => {
    expect([ru('a"b'), ru(1e21), ru(1e-7), ru(1.5), ru(0.1 + 0.2), ru(true), ru({ error: 'NA' })]).toEqual(['"a""b"', '1E+21', '1E-7', '1,5', '0,30000000000000004', 'ИСТИНА', '#Н/Д'])
  })

  it.each(['closures', 'eval'] as const)('shows so in explain (%s)', (execution) => {
    const s = createSession({ locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { value: 'a"b' }, B1: { formula: 'A1&"x"' }, A2: { value: 1e21 }, B2: { formula: 'A2*1' } } }] }, { execution })
    expect(s.explain('S', 'B1').text).toBe('S!B1 = "a""bx"  ← A1&"x"\n  S!A1 = "a""b"')
    expect(s.explain('S', 'B2').text).toBe('S!B2 = 1E+21  ← A2*1\n  S!A2 = 1E+21')
    s.close()
  })
})
