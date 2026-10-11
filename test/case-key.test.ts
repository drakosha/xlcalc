/**
 * The case of names and words. Names (`caseKey`: defined names, tables, sheets, the specifiers of structured
 * references) compare by Unicode's default upper case, which no locale changes, a letter's whole (`straße` is
 * `STRASSE`), but for the dotless ı and the long ſ, which Excel for Windows keeps apart from I and S (the table sets
 * `name-case-held`, `name-case-sheets-held`). Words of formula text (`wordKey`: function names, TRUE and FALSE, LET's
 * names, a locale's words) take each letter's own upper case where it is one letter, ı and ſ kept (`gauß` is no GAUSS).
 * Turkish rules differ and are not followed (Excel tr-TR not measured).
 */
import { describe, expect, it } from 'vitest'
import { caseKey, wordKey } from '../src/collate.js'
import { compute } from '../src/index.js'
import './locales.js'

/** The characters of the code points from `from`, below `to`, without lone surrogates. */
function chars(from: number, to: number): string {
  let s = ''
  for (let cp = from; cp < to; cp++) if (cp < 0xd800 || cp > 0xdfff) s += String.fromCodePoint(cp)
  return s
}

describe('the case of names', () => {
  it('is Russian and English rules’ upper case for every code point but ı and ſ', () => {
    const off: string[] = []
    // upper case maps a character alone in these rules, so a block of characters at a time is the same check
    for (let from = 0; from < 0x110000; from += 0x1000) {
      const block = chars(from, from + 0x1000)
      const key = caseKey(block)
      const upper = (tag: string) => block.replace(/[^ıſ]+/g, (run) => run.toLocaleUpperCase(tag))
      if (key !== upper('ru-RU') || key !== upper('en-US')) off.push(from.toString(16))
    }
    expect(off).toEqual([])
  })

  it('a word takes the upper case of each letter that has a letter for it, the name key elsewhere', () => {
    const off: string[] = []
    for (let cp = 0; cp < 0x110000; cp++) {
      if (cp >= 0xd800 && cp <= 0xdfff) continue
      const ch = String.fromCodePoint(cp)
      const one = [...caseKey(ch)].length === 1
      if (wordKey(ch) !== (one ? caseKey(ch) : ch)) off.push(cp.toString(16))
    }
    expect(off).toEqual([])
    expect(['gauß', 'ﬁnd', 'falſe', 'ıf', 'İf', 'сумм', 'straße'].map(wordKey)).toEqual(['GAUß', 'ﬁND', 'FALſE', 'ıF', 'İF', 'СУММ', 'STRAßE'])
  })

  it('compares Cyrillic and Latin names in any case, ё and й apart from е and и', () => {
    expect(['лист', 'Данные', 'ёлка', 'йод', 'Sheet', 'straße'].map(caseKey)).toEqual(['ЛИСТ', 'ДАННЫЕ', 'ЁЛКА', 'ЙОД', 'SHEET', 'STRASSE'])
    expect(caseKey('ёлка')).not.toBe(caseKey('елка'))
  })

  it('does not follow Turkish rules: i is I, ı and İ stay apart', () => {
    expect([caseKey('i'), caseKey('ı'), caseKey('İ'), caseKey('ſ'), caseKey('ﬁ')]).toEqual(['I', 'ı', 'İ', 'ſ', 'FI'])
    expect('i'.toLocaleUpperCase('tr-TR')).toBe('İ')
  })

  it('a ru-RU book finds a sheet and a table typed in another case', () => {
    const r = compute({
      locale: 'ru-RU',
      sheets: [
        { name: 'Данные', cells: { A1: { value: 'Кол' }, A2: { value: 2 }, A3: { value: 3 } }, tables: [{ name: 'Продажи', range: 'A1:A3', columns: [{ name: 'Кол' }] }] },
        { name: 'S', cells: { A1: { formula: "SUM('ДАННЫЕ'!A2:A3)" }, A2: { formula: 'SUM(продажи[кол])' }, A3: { formula: 'ROWS(ПРОДАЖИ[#data])' } } },
      ],
    })
    const cells = r.sheets[1].cells
    expect([cells.A1.value, cells.A2.value, cells.A3.value]).toEqual([5, 5, 2])
  })
})
