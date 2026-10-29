/**
 * REGEXTEST, REGEXEXTRACT and REGEXREPLACE on each level of regular expressions, registered in turn (built in,
 * `pcreSyntax`, `pcreEngine`): the same answers where the levels read the pattern alike, the groups of a pattern that
 * `pcreSyntax` translates with groups of its own, empty matches, case sensitivity, replacements, the match limit.
 * Neither Excel the harness drives has these functions (`regex-functions-held`: #NAME? in Excel for Windows 16.0.20430).
 */
import { describe, expect, it } from 'vitest'
import { compute, registerRegex } from '../src/index.js'
import { pcreEngine, pcreSyntax } from '../src/regex/index.js'

const values = (formulas: string[]) => {
  const r = compute({ sheets: [{ name: 'S', cells: Object.fromEntries(formulas.map((f, i) => [`A${i + 1}`, { formula: f }])) }] })
  return formulas.map((_, i) => r.sheets[0].cells[`A${i + 1}`].value)
}

/** Formulas every level answers alike, and their answers. */
const COMMON: [string, unknown][] = [
  ['REGEXTEST("Excel 2025","\\d+")', true],
  ['REGEXTEST("Excel","excel")', false],
  ['REGEXTEST("Excel","excel",1)', true],
  ['REGEXTEST("Excel","x",2)', { error: 'VALUE' }],
  ['REGEXTEST("abc","[")', { error: 'VALUE' }],
  ['REGEXTEST(123,"^\\d+$")', true],
  ['REGEXEXTRACT("tel 555-1234 or 555-9876","\\d{3}-\\d{4}")', '555-1234'],
  ['REGEXEXTRACT("abc","\\d")', { error: 'NA' }],
  ['TEXTJOIN(";",,REGEXEXTRACT("a1b22c333","\\d+",1))', '1;22;333'],
  ['ROWS(REGEXEXTRACT("a1b22c333","\\d+",1))', 1],
  ['REGEXEXTRACT("abc","\\d",1)', { error: 'NA' }],
  ['TEXTJOIN(";",FALSE,REGEXEXTRACT("John Smith","(\\w+) (x)?(\\w+)",2))', 'John;;Smith'],
  ['REGEXEXTRACT("ABC","b",0,1)', 'B'],
  ['REGEXEXTRACT("abc","b",3)', { error: 'VALUE' }],
  ['TEXTJOIN(";",FALSE,REGEXEXTRACT("ab","x*",1))', ';;'],
  ['TEXTJOIN(";",FALSE,REGEXEXTRACT("привет мир","\\w+",1))', 'привет;мир'],
  ['REGEXREPLACE("a1b2c3","\\d","#")', 'a#b#c#'],
  ['REGEXREPLACE("a1b2c3","\\d","#",2)', 'a1b#c3'],
  ['REGEXREPLACE("a1b2c3","\\d","#",-1)', 'a1b2c#'],
  ['REGEXREPLACE("a1b2c3","\\d","#",4)', 'a1b2c3'],
  ['REGEXREPLACE("a1b2c3","\\d","#",-4)', 'a1b2c3'],
  ['REGEXREPLACE("SoniaBrown","([A-Z][a-z]+)([A-Z][a-z]+)","$2, $1")', 'Brown, Sonia'],
  ['REGEXREPLACE("ab","(a)","${1}1$$$9$12")', 'a1$$9a2b'],
  ['REGEXREPLACE("abc","B","x",0,1)', 'axc'],
  ['REGEXREPLACE("abc","","-")', '-a-b-c-'],
  ['REGEXREPLACE("a😀b","","-")', '-a-😀-b-'],
  ['LEN(REGEXREPLACE(REPT("a",20000),"a","aa"))', { error: 'VALUE' }],
]

const levels: [string, () => void][] = [
  ['built in', () => {}],
  ['pcreSyntax', () => registerRegex(pcreSyntax)],
  ['pcreEngine', () => registerRegex(pcreEngine)],
]

for (const [name, load] of levels)
  describe(`the REGEX functions, ${name}`, () => {
    it('answer as their description says', () => {
      load()
      expect(values(COMMON.map(([f]) => f))).toEqual(COMMON.map(([, v]) => v))
    })

    // the built-in level takes JavaScript's syntax: an atomic group is no pattern there
    it.skipIf(name === 'built in')("groups are the pattern's own, whatever the level writes besides (an atomic group, a possessive repeat)", () => {
      expect(values(['TEXTJOIN(";",FALSE,REGEXEXTRACT("aaab","(?>a+)(b)",2))', 'REGEXREPLACE("aaab-c","a++(b)-(c)","$2$1")'])).toEqual(['b', 'cb'])
    })
  })

describe('the REGEX functions, pcreEngine: the match limit', () => {
  it('is #VALUE!', () => {
    registerRegex(pcreEngine)
    const heavy = `"${'a'.repeat(40)}b","(a+)+$"`
    expect(values([`REGEXTEST(${heavy})`, `REGEXEXTRACT(${heavy})`, `REGEXREPLACE(${heavy},"x")`])).toEqual([{ error: 'VALUE' }, { error: 'VALUE' }, { error: 'VALUE' }])
  })
})
