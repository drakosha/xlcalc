/**
 * `loadLocale(tag)` loads one of the package's locales by its tag, from the root and from `xlcalc/xlsx`: a workbook in
 * that locale is taken from then on, a file is written in it, and a pool started before it computes in it, as after
 * `registerLocale`. Loaded already, it is given back without importing its module again; another tag is an
 * `InputError`. Before it a workbook in the locale is refused with the hint that names it. The file does not import
 * `locales.ts`.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, loadLocale, LOCALES, type BookInput } from '../src/index.js'
import * as xlsx from '../src/xlsx.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { closedAfter } from './fixtures/pool.js'

const clock = { now: '2026-10-05T12:00:00Z', seed: 1 }
const book: BookInput = { locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { formula: 'VALUE("1,5")*2' } } }] }

const { made } = closedAfter()

describe('loadLocale', () => {
  it("loads a package's locale by its tag, for workbooks, files and a pool's threads, once", async () => {
    const pool = await made({ threads: 1 })
    expect(() => compute(book, { clock })).toThrow("book.locale: the locale ru-RU is not loaded: import { ruRU } from 'xlcalc/locales/ru-RU' and call registerLocale(ruRU), or await loadLocale('ru-RU') (loaded: en-US)")
    const [a, b] = await Promise.all([xlsx.loadLocale('ru-RU'), loadLocale('ru-RU')])
    expect([a, b, LOCALES.get('ru-RU')]).toEqual([ruRU, ruRU, ruRU])
    expect(a).toBe(b)
    expect(await loadLocale('ru-RU')).toBe(a)
    expect(compute(book, { clock }).sheets[0].cells.A1.value).toBe(3)
    expect((await pool.compute(book, { clock })).sheets[0].cells.A1.value).toBe(3)
    expect((await xlsx.readXlsx((await xlsx.writeXlsx({ book })).data)).book.locale).toBe('ru-RU')
    expect(await loadLocale('en-US')).toBe(LOCALES.get('en-US'))
  })

  it('refuses a tag that is not one of the package\'s locales, naming the tag of another case', async () => {
    await expect(loadLocale('ru-ru')).rejects.toThrow(new InputError('tag', `"ru-ru" is not one of the package's locales (en-US, ru-RU); tags are case-sensitive: ru-RU`))
    await expect(xlsx.loadLocale('de-DE')).rejects.toThrow(new InputError('tag', `"de-DE" is not one of the package's locales (en-US, ru-RU)`))
    await expect(loadLocale(5 as never)).rejects.toThrow(InputError)
  })
})
