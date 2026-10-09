/**
 * Each locale's formula text against Excel: `test/excel/locale/<tag>.json` holds Excel 16.113's own text of the
 * probes of `bench/locale/probes.ts` (the coverage measurement's cases, the probes, a call of every listed function
 * the measurement lacks, literals) in English (`formula`) and in the locale's language (`formula local`), and its
 * international settings (`bench/locale/snapshot.ts`). The locale's data equals what that text gives
 * (`bench/locale/derive.ts`, a reading independent of the library's lexer), and translation reproduces Excel's text
 * both ways. A locale without a snapshot is a skipped test that names the missing file.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { derive, loadSnapshot, snapshotFile } from '../bench/locale/derive.js'
import { formulaText, withPart } from '../bench/locale/generate.js'
import { LOCALES, toInvariant, toLocal } from '../src/index.js'

for (const [tag, loc] of LOCALES) {
  if (!existsSync(snapshotFile(tag))) {
    it.skip(`${tag} locale data against Excel: no snapshot of Excel's formula text (test/excel/locale/${tag}.json)`)
    continue
  }
  const snap = loadSnapshot(tag)
  const d = derive(snap)
  const formulas = Object.entries(snap.formulas)

  describe(`${tag} locale data against Excel`, () => {
    it('the snapshot reads without contradictions', () => {
      expect(d.problems).toEqual([])
    })

    it('separators, TRUE/FALSE and error literals are what Excel writes', () => {
      const { tag: _tag, functions: _f, tableItems: _t, ...rest } = loc
      expect(rest).toEqual({
        decimalSeparator: d.decimalSeparator,
        listSeparator: d.listSeparator,
        arrayColumnSeparator: d.arrayColumnSeparator,
        arrayRowSeparator: d.arrayRowSeparator,
        booleans: d.booleans,
        errors: d.errors,
      })
      // array constants as Excel writes them, not only as its settings say
      const [col, row] = [loc.arrayColumnSeparator, loc.arrayRowSeparator]
      expect(snap.formulas['lit:8']).toEqual(['SUM({1,2;3,4})', `${loc.functions.SUM ?? 'SUM'}({1${col}2${row}3${col}4})`])
    })

    it('every function Excel names has its local name in the locale, and no other entries', () => {
      const named = Object.fromEntries(Object.keys(d.functions).map((f) => [f, loc.functions[f] ?? f]))
      expect(named).toEqual(d.functions)
      expect(Object.keys(loc.functions).filter((f) => !(f in d.functions))).toEqual([])
    })

    it('the locale file holds what bench/locale/generate.ts writes between its marks', () => {
      const source = readFileSync(new URL(`../src/locales/${tag}.ts`, import.meta.url), 'utf8')
      expect(withPart(source, formulaText(d, snap.meta.excel))).toBe(source)
    })

    it('no local name is the English name of another function (English names are read in local text too)', () => {
      const english = new Set(Object.keys(d.functions))
      expect(Object.entries(loc.functions).filter(([en, l]) => english.has(l.toUpperCase()) && l.toUpperCase() !== en)).toEqual([])
    })
  })

  describe(`${tag}: translation reproduces Excel`, () => {
    it('toLocal of Excel’s English text is its local text, for every formula', () => {
      const off = formulas.filter(([, [en, local]]) => {
        const r = toLocal(en, loc)
        return !r.ok || r.formula !== local
      })
      expect(off.slice(0, 10)).toEqual([])
    })

    it('toInvariant of Excel’s local text is its English text, for every formula', () => {
      const off = formulas.filter(([, [en, local]]) => {
        const r = toInvariant(local, loc)
        return !r.ok || r.formula !== en
      })
      expect(off.slice(0, 10)).toEqual([])
    })
  })
}

describe('the ru-RU snapshot: Excel 16.113 for Mac', () => {
  const snap = loadSnapshot('ru-RU')
  const d = derive(snap)

  it('covers the measurement and the probes; 513 functions with a Russian name, the rest unknown to this Excel or deleted from books on load', () => {
    const groups = (p: string) => Object.keys(snap.formulas).filter((id) => id.startsWith(p)).length
    expect([groups('m:'), groups('p:'), groups('fn:'), groups('lit:'), groups('err:')]).toEqual([5275, 959, 117, 22, 8])
    expect(Object.keys(d.functions)).toHaveLength(513)
    expect(d.unknown).toEqual(['ENCODEURL', 'FILTERXML', 'REGEXEXTRACT', 'REGEXREPLACE', 'REGEXTEST', 'WEBSERVICE'])
    expect(d.dropped).toEqual(['CALL', 'REGISTER.ID'])
    expect(snap.rejected).toEqual(['#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!', '#PYTHON!'].map((e) => `err:${e}`))
  })
})
