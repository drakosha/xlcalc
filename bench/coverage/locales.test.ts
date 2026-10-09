/**
 * The harnesses' locales (`locales.ts`): what a command line names, where answers go, how Excel for Mac starts and what
 * Windows needs. en-US and ru-RU keep the names, directories and launch arguments their stored answers were taken with;
 * another tag takes the general form. Stored answers in a locale's directory belong to a locale of the package.
 */
import { existsSync, readdirSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { launchArgs, localeDir, localeMeta, measurementFile, parseTag, TAGS, windowsLocale } from './locales.js'

describe('the harnesses\' locales', () => {
  it('a command line names a locale by en, ru or a tag of a language and a region, en-US when it names none', () => {
    expect([undefined, 'en', 'ru', 'ru-RU', 'de-DE', 'xx-YY'].map((a) => parseTag(a))).toEqual(['en-US', 'en-US', 'ru-RU', 'ru-RU', 'de-DE', 'xx-YY'])
    expect(parseTag(undefined, 'ru-RU')).toBe('ru-RU')
    for (const bad of ['de', 'ru_RU', 'RU-ru', 'en-us', 'russian', '']) expect(() => parseTag(bad)).toThrow(/^locale: en, ru or a tag/)
  })

  it('answers go to the top of their directory for en-US and to a directory of the tag for another locale', () => {
    expect(['en-US', 'ru-RU', 'xx-YY'].map(localeDir)).toEqual(['', 'ru-RU/', 'xx-YY/'])
    expect(['en-US', 'ru-RU', 'xx-YY'].map(measurementFile)).toEqual(['excel-en.json', 'excel-ru.json', 'excel-xx-YY.json'])
  })

  it('Excel for Mac starts as the stored answers were taken: en-US with the English interface and locale, ru-RU without arguments', () => {
    expect(launchArgs('en-US')).toEqual(['-AppleLanguages', '(en)', '-AppleLocale', 'en_US'])
    expect(launchArgs('ru-RU')).toEqual([])
    expect(launchArgs('de-DE')).toEqual(['-AppleLanguages', '(de)', '-AppleLocale', 'de_DE'])
    expect(['en-US', 'ru-RU'].map(localeMeta)).toEqual(['en-US (-AppleLanguages (en) -AppleLocale en_US)', 'ru-RU (the system language and locale)'])
  })

  it("Windows takes a locale's language id and code page from its data, and only the package's locales", () => {
    expect(windowsLocale('en-US')).toEqual({ lcid: 1033, codepage: 1252 })
    expect(windowsLocale('ru-RU')).toEqual({ lcid: 1049, codepage: 1251 })
    for (const tag of TAGS) expect(Object.values(windowsLocale(tag)).every(Number.isInteger), tag).toBe(true)
    expect(() => windowsLocale('de-DE')).toThrow(/de-DE is none of the package's locales .*src\/locales/)
  })

  it('every locale with stored answers is one of the package\'s', () => {
    const root = new URL('../../test/excel/', import.meta.url)
    const dirs = (rel: string) => (existsSync(new URL(rel, root)) ? readdirSync(new URL(rel, root), { withFileTypes: true }).filter((d) => d.isDirectory()).map((d) => d.name) : [])
    const kinds = ['', 'win/'].flatMap((w) => dirs(w || '.').map((k) => `${w}${k}/`))
    const tags = kinds.flatMap((k) => dirs(k).filter((d) => /^[a-z]{2,3}-[A-Z]{2}$/.test(d)).map((d) => `${k}${d}`))
    expect(tags.filter((t) => !TAGS.includes(t.split('/').pop()!))).toEqual([])
    const measured = ['coverage/', 'win/coverage/'].flatMap((d) => (existsSync(new URL(d, root)) ? readdirSync(new URL(d, root)).filter((f) => f.startsWith('excel-')) : []))
    expect(measured.filter((f) => !TAGS.some((t) => measurementFile(t) === f))).toEqual([])
  })
})
