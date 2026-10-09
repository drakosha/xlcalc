/**
 * ru-RU against Excel: `test/excel/locale/ru-RU.json` holds Excel 16.113's own text of 6,381 formulas (the
 * coverage measurement's cases, the probes, a call of every listed function the measurement lacks, literals)
 * in English (`formula`) and Russian (`formula local`), and its international settings
 * (`bench/locale/snapshot.ts`). The locale data equals what that text gives (`bench/locale/derive.ts`, a
 * reading independent of the library's lexer), and translation reproduces Excel's text both ways.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { derive, loadSnapshot } from '../bench/locale/derive.js'
import { formulaText, withPart } from '../bench/locale/generate.js'
import { ruRU, toInvariant, toLocal } from '../src/index.js'

const snap = loadSnapshot('ru-RU')
const d = derive(snap)
const formulas = Object.entries(snap.formulas)

describe('ru-RU locale data against Excel', () => {
  it('the snapshot reads without contradictions and covers the measurement and the probes', () => {
    expect(d.problems).toEqual([])
    const groups = (p: string) => formulas.filter(([id]) => id.startsWith(p)).length
    expect([groups('m:'), groups('p:'), groups('fn:'), groups('lit:'), groups('err:')]).toEqual([5275, 959, 117, 22, 8])
  })

  it('separators, TRUE/FALSE and error literals are what Excel writes', () => {
    const { tag: _tag, functions: _f, tableItems: _t, ...rest } = ruRU
    expect(rest).toEqual({
      decimalSeparator: d.decimalSeparator,
      listSeparator: d.listSeparator,
      arrayColumnSeparator: d.arrayColumnSeparator,
      arrayRowSeparator: d.arrayRowSeparator,
      booleans: d.booleans,
      errors: d.errors,
    })
    // array constants as Excel writes them, not only as its settings say
    expect(snap.formulas['lit:8']).toEqual(['SUM({1,2;3,4})', 'СУММ({1;2:3;4})'])
  })

  it('every function Excel names has its Russian name in the locale, and no other entries', () => {
    const named = Object.fromEntries(Object.keys(d.functions).map((f) => [f, ruRU.functions[f] ?? f]))
    expect(named).toEqual(d.functions)
    expect(Object.keys(ruRU.functions).filter((f) => !(f in d.functions))).toEqual([])
    expect(Object.keys(d.functions)).toHaveLength(513)
  })

  it('functions without a Russian name from Excel: unknown to Excel 16.113 for Mac, or deleted from books on load', () => {
    expect(d.unknown).toEqual(['ENCODEURL', 'FILTERXML', 'REGEXEXTRACT', 'REGEXREPLACE', 'REGEXTEST', 'WEBSERVICE'])
    expect(d.dropped).toEqual(['CALL', 'REGISTER.ID'])
    expect(snap.rejected).toEqual(['#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!', '#PYTHON!'].map((e) => `err:${e}`))
  })

  it('the locale file holds what bench/locale/generate.ts writes between its marks', () => {
    const source = readFileSync(new URL('../src/locales/ru-RU.ts', import.meta.url), 'utf8')
    expect(withPart(source, formulaText(d, snap.meta.excel))).toBe(source)
  })

  it('no Russian name is the English name of another function (English names are read in Russian text too)', () => {
    const english = new Set(Object.keys(d.functions))
    expect(Object.entries(ruRU.functions).filter(([en, l]) => english.has(l.toUpperCase()) && l.toUpperCase() !== en)).toEqual([])
  })
})

describe('translation reproduces Excel', () => {
  it('toLocal of Excel’s English text is its Russian text, for every formula', () => {
    const off = formulas.filter(([, [en, loc]]) => {
      const r = toLocal(en, ruRU)
      return !r.ok || r.formula !== loc
    })
    expect(off.slice(0, 10)).toEqual([])
  })

  it('toInvariant of Excel’s Russian text is its English text, for every formula', () => {
    const off = formulas.filter(([, [en, loc]]) => {
      const r = toInvariant(loc, ruRU)
      return !r.ok || r.formula !== en
    })
    expect(off.slice(0, 10)).toEqual([])
  })
})
