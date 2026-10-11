/**
 * The case of names (`caseKey`): sheets, tables, the specifiers of structured references and the words of typed formula
 * text compare by Unicode's default upper case, which no locale changes. It was Russian rules' upper case
 * (`toLocaleUpperCase('ru-RU')`); both give the same key for every character, so nothing compares otherwise in ru-RU or
 * en-US books. Turkish rules differ and are not followed (Excel tr-TR not measured).
 */
import { describe, expect, it } from 'vitest'
import { caseKey } from '../src/collate.js'
import { compute } from '../src/index.js'
import './locales.js'

/** The characters of the code points from `from`, below `to`, without lone surrogates. */
function chars(from: number, to: number): string {
  let s = ''
  for (let cp = from; cp < to; cp++) if (cp < 0xd800 || cp > 0xdfff) s += String.fromCodePoint(cp)
  return s
}

describe('the case of names', () => {
  it('is Russian and English rules’ upper case for every code point', () => {
    const off: string[] = []
    // upper case maps a character alone in these rules, so a block of characters at a time is the same check
    for (let from = 0; from < 0x110000; from += 0x1000) {
      const block = chars(from, from + 0x1000)
      const key = caseKey(block)
      if (key !== block.toLocaleUpperCase('ru-RU') || key !== block.toLocaleUpperCase('en-US')) off.push(from.toString(16))
    }
    expect(off).toEqual([])
  })

  it('compares Cyrillic and Latin names in any case, ё and й apart from е and и', () => {
    expect(['лист', 'Данные', 'ёлка', 'йод', 'Sheet', 'straße'].map(caseKey)).toEqual(['ЛИСТ', 'ДАННЫЕ', 'ЁЛКА', 'ЙОД', 'SHEET', 'STRASSE'])
    expect(caseKey('ёлка')).not.toBe(caseKey('елка'))
  })

  it('does not follow Turkish rules: i is I, ı is I, İ stays apart', () => {
    expect([caseKey('i'), caseKey('ı'), caseKey('İ')]).toEqual(['I', 'I', 'İ'])
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
