/**
 * The book's locale as input (`BookInput.locale`): the tag of a loaded locale is checked, and it takes part in the hashes only
 * when it is not en-US, so books without it hash as before (the synthetic book's hash is pinned in
 * `sha256.test.ts`).
 */
import { describe, expect, it } from 'vitest'
import { valuesHash } from '../src/canon.js'
import type { BookInput, CellValue, ComputeResult } from '../src/contract.js'
import { compute, CONFIG_VERSION, enUS, InputError, LOCALES as PUBLIC_LOCALES, createSession } from '../src/index.js'
import { ruRU } from './locales.js'
import { LOCALES } from '../src/locale/registry.js'
import { displayRules } from '../src/locale/rules.js'

const clock = { now: '2026-10-05T12:00:00Z', seed: 1 }
const book = (locale?: string): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*3' } } }], ...(locale === undefined ? {} : { locale }) })
const valuesOf = (r: ComputeResult) => new Map(r.sheets.map((s) => [s.name, new Map(Object.entries(s.cells).map(([a, c]): [string, CellValue] => [a, c.value]))]))

describe('the book locale', () => {
  it('takes the tags of the loaded locales; another tag or a non-string is an InputError naming book.locale and the loaded tags', () => {
    for (const tag of ['en-US', 'ru-RU']) expect(() => compute(book(tag), { clock })).not.toThrow()
    for (const bad of ['de-DE', 'ru-ru', 'ru', '', 'toString']) {
      expect(() => compute(book(bad), { clock })).toThrow(InputError)
      expect(() => compute(book(bad), { clock })).toThrow(`(loaded: ${[...LOCALES.keys()].join(', ')};`)
      expect(() => compute(book(bad), { clock })).toThrow(/^book\.locale: unknown locale .* \(loaded: [^()]*\)$/)
    }
    expect(() => compute(book('ru-ru'), { clock })).toThrow('tags are case-sensitive: ru-RU')
    expect(() => compute({ ...book(), locale: 7 } as unknown as BookInput, { clock })).toThrow('book.locale: expected a string, got number 7')
    expect(() => createSession(book('xx-XX'), { clock })).toThrow(InputError)
  })

  it('en-US hashes as a book without a locale; another locale changes both hashes', () => {
    const none = compute(book(), { clock })
    const en = compute(book('en-US'), { clock })
    const ru = compute(book('ru-RU'), { clock })
    expect(en.hash).toBe(none.hash)
    expect(en.inputHash).toBe(none.inputHash)
    expect(ru.inputHash).not.toBe(none.inputHash)
    expect(ru.hash).not.toBe(none.hash)
    expect(ru.configVersion).toBe(CONFIG_VERSION)
  })

  it('the values hash takes ;locale=<tag> after the configuration version', () => {
    const ru = compute(book('ru-RU'), { clock })
    expect(ru.hash).toBe(valuesHash(valuesOf(ru), `${CONFIG_VERSION};locale=ru-RU`))
    const none = compute(book(), { clock })
    expect(none.hash).toBe(valuesHash(valuesOf(none), CONFIG_VERSION))
  })

  it('LOCALES lists the loaded locales by tag: the tags a book takes, each the locale itself', () => {
    expect([...PUBLIC_LOCALES.keys()]).toEqual(expect.arrayContaining(['en-US', 'ru-RU']))
    expect(PUBLIC_LOCALES.get('en-US')).toBe(enUS)
    expect(PUBLIC_LOCALES.get('ru-RU')).toBe(ruRU)
    for (const tag of PUBLIC_LOCALES.keys()) expect(() => compute(book(tag), { clock })).not.toThrow()
  })

  it('every loaded locale other than en-US has its rules of computation', () => {
    for (const [tag, l] of LOCALES) expect(tag === 'en-US' || l.regional !== undefined, tag).toBe(true)
    for (const tag of LOCALES.keys()) expect(displayRules(tag).locale).toBe(LOCALES.get(tag))
  })

  it('a session keeps the locale through edits and structure operations', () => {
    const s = createSession(book('ru-RU'), { clock })
    s.setCells([{ sheet: 'S', cell: 'B1', input: { value: 1 } }])
    expect(s.getBook().locale).toBe('ru-RU')
    const r = s.editStructure({ op: 'insertRows', sheet: 'S', row: 1, count: 1 })
    expect(r.ok && s.getBook().locale).toBe('ru-RU')
    expect(s.getResult().hash).toBe(compute(s.getBook(), { clock: s.clock }).hash)
  })
})
