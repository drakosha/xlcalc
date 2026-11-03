/**
 * `numberFormatSamples` against the samples of the Number Format list of Excel for Windows (`test/excel/win/samples/`,
 * the Windows step `samples` over the sets `basic` and `held` of `bench/coverage/samples/`, en-US and ru-RU): every
 * sample of every cell, the known difference by locale and item with its reason; the list's item names; the
 * predictions committed before each run (`bench/windows/samples/<set>.predicted.json`) with how many samples Excel
 * confirmed; the function's input checks.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredSamples } from '../bench/windows/samples.js'
import { InputError, numberFormatChoices, numberFormatSamples } from '../src/index.js'
import * as entry from '../src/typed-entry.js'
import { localeSets } from './stored-sets.js'
import { ruRU } from './locales.js'

/** The items whose samples knowingly differ from Excel for Windows', by `<tag>:<key>`: how the engine's text differs, and why. */
const KNOWN: Record<string, { engine: (excel: string) => RegExp; why: string }> = {
  'ru-RU:longDate': {
    engine: (excel) => new RegExp(`^${excel.startsWith('-') ? '-' : ''}[а-я]+, ${excel.replace(/^-/, '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`),
    why: "the engine keeps Excel for Mac's ru-RU long date [$-F800] with the weekday (понедельник, 18 мая 1903 г.), as a cell in Long Date shows it; Excel for Windows writes it without",
  },
}

/** How many samples of each run are the known difference: every long date ru-RU shows. */
const KNOWN_COUNT: Record<string, number> = { 'ru-RU:basic': 17, 'ru-RU:held': 8 }

/** How many samples of each set's prediction Excel for Windows confirmed, by `<tag>:<set>`. */
const CONFIRMED: Record<string, number> = { 'en-US:basic': 299, 'ru-RU:basic': 282, 'en-US:held': 303, 'ru-RU:held': 295 }

const RUNS = localeSets<StoredSamples>('win/samples')

describe('numberFormatSamples against Excel for Windows', () => {
  it('ran both sets in en-US and ru-RU', () => {
    expect(RUNS.map((r) => `${r.tag}:${r.set}`)).toEqual(['en-US:basic', 'en-US:held', 'ru-RU:basic', 'ru-RU:held'])
  })

  for (const { tag, set, stored } of RUNS) {
    it(`${tag} ${set}: every sample as Excel showed it, but the known difference`, () => {
      const names = numberFormatChoices(tag).map((c) => c.name)
      let known = 0
      for (const c of stored.cases) {
        const a = stored.answers[c.id]
        expect(a.items.map(([name]) => name), c.id).toEqual(names)
        const got = numberFormatSamples(c.value, { locale: tag, ...(c.dateSystem ? { dateSystem: c.dateSystem } : {}) })
        got.forEach(({ key, sample }, i) => {
          const excel = a.items[i][1]
          const k = KNOWN[`${tag}:${key}`]
          if (k && sample !== excel && k.engine(excel).test(sample)) known++
          else expect(sample, `${c.id} ${key}`).toBe(excel)
        })
      }
      expect(known).toBe(KNOWN_COUNT[`${tag}:${set}`] ?? 0)
    })

    it(`${tag} ${set}: the prediction committed before the run, ${CONFIRMED[`${tag}:${set}`]} samples confirmed`, () => {
      const predicted = (JSON.parse(readFileSync(new URL(`../bench/windows/samples/${set}.predicted.json`, import.meta.url), 'utf8')) as { locales: Record<string, Record<string, string[]>> }).locales[tag]
      const same = stored.cases.reduce((n, c) => n + stored.answers[c.id].items.filter(([, s], i) => s === predicted[c.id][i]).length, 0)
      expect(same).toBe(CONFIRMED[`${tag}:${set}`])
    })
  }
})

describe('numberFormatSamples', () => {
  it("gives each item of the list in its order, General's words for every value", () => {
    expect(numberFormatSamples(1234.5)).toEqual([
      { key: 'general', sample: 'No specific format' },
      { key: 'number', sample: '1234.50' },
      { key: 'currency', sample: '$1,234.50' },
      { key: 'accounting', sample: ' $1,234.50 ' },
      { key: 'shortDate', sample: '5/18/1903' },
      { key: 'longDate', sample: 'Monday, May 18, 1903' },
      { key: 'time', sample: '12:00:00 PM' },
      { key: 'percentage', sample: '123450.00%' },
      { key: 'fraction', sample: '1234 1/2' },
      { key: 'scientific', sample: '1.23E+03' },
      { key: 'text', sample: '1234.5' },
    ])
    expect(numberFormatSamples(null, { locale: ruRU }).map((s) => s.sample)).toEqual(['Формат не указан', ...Array(10).fill('')])
  })

  it('cuts a sample past 50 characters, and writes ten # for a number a format cannot show', () => {
    const s = (v: Parameters<typeof numberFormatSamples>[0], key: string) => numberFormatSamples(v).find((x) => x.key === key)!.sample
    expect(s('x'.repeat(50), 'text')).toBe('x'.repeat(50))
    expect(s('x'.repeat(51), 'text')).toBe(`${'x'.repeat(50)}...`)
    expect(s('x'.repeat(49), 'accounting')).toBe(` ${'x'.repeat(49)}...`)
    expect([s(-1, 'shortDate'), s(1e307, 'percentage'), s('#####', 'shortDate')]).toEqual(['##########', '##########', '#####'])
  })

  it('in the date system given, the same object in xlcalc/entry', () => {
    expect(numberFormatSamples(0, { dateSystem: 1904 })[4].sample).toBe('1/1/1904')
    expect(numberFormatSamples(-0.25, { dateSystem: 1904, locale: ruRU })[6].sample).toBe('-6:00:00')
    expect(entry.numberFormatSamples(46000, { locale: ruRU })).toEqual(numberFormatSamples(46000, { locale: 'ru-RU' }))
  })

  it('checks its input', () => {
    const fails = (f: () => unknown, message: string) => expect(f).toThrow(new InputError(message.split(':')[0], message.slice(message.indexOf(':') + 2)))
    fails(() => numberFormatSamples(1, { numberFormat: 2 } as never), 'opts: unknown field "numberFormat" (expected locale, dateSystem)')
    fails(() => numberFormatSamples(1, { dateSystem: 1901 } as never), 'opts.dateSystem: expected one of 1900, 1904, got number 1901')
    fails(() => numberFormatSamples(1, 5 as never), "opts: expected options such as { locale: 'ru-RU', dateSystem: 1904 }, got number 5")
    expect(() => numberFormatSamples(Symbol() as never)).toThrow(InputError)
  })
})
