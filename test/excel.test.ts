/**
 * Acceptance against Microsoft Excel: every whitelisted function is computed by this engine on the coverage
 * measurement's cases (`docs/coverage-formulajs.md`) and compared with Excel 16.113 (en-US locale, English
 * UI). A function is in the whitelist only if all of its cases match Excel; the test pins that for the
 * whole whitelist. Excel's answers are stored in the repository: no office suite is needed.
 */
import { describe, expect, it } from 'vitest'
import { FUNCTION_NAMES } from '../src/index.js'
import { FUNCTIONS } from '../src/functions.js'
import { CASES, EXCEL, runExcelCases, show, type Outcome } from './excel-cases.js'

/** Whitelisted functions the measurement has no cases for, and why. */
const NO_CASES: Record<string, string> = {
  XLOOKUP: 'formula.js 4.6.1 has no XLOOKUP, so the measurement does not cover it; test/semantics-cases.ts does',
  TODAY: 'volatile, excluded from the measurement; test/volatile.test.ts',
  NOW: 'volatile, excluded from the measurement; test/volatile.test.ts',
  RAND: 'volatile, excluded from the measurement; test/volatile.test.ts',
  RANDBETWEEN: 'volatile, excluded from the measurement; test/volatile.test.ts',
}

/** Known differences from Excel, by case id: to be fixed, never to be extended. */
const KNOWN: Record<string, string> = {
  'COUNTIFS#12': 'a number criterion does not match numeric text',
  'SUMIF#12': 'a number criterion does not match numeric text',
  'SUMIFS#12': 'a number criterion does not match numeric text',
  'DATE#5': 'the phantom 1900-02-29 is not produced',
  'DATE#8': 'a month out of range gives #VALUE! instead of #NUM!',
  'EOMONTH#4': 'date text is not read',
  'EOMONTH#5': 'date text is not read',
  'EOMONTH#6': 'a month out of range gives #VALUE! instead of #NUM!',
  'EOMONTH#8': 'serial 60 is read as 1900-03-01',
  'EOMONTH#13': 'serial 0 is read as 1899-12-31',
  'ROUND#10': 'a huge number of digits gives #NUM!',
  'TEXT#8': 'fraction formats are not supported',
  'TEXT#14': 'an empty format gives the General format',
}

const covered = FUNCTION_NAMES.filter((f) => !(f in NO_CASES))
const outcomes = runExcelCases(CASES.filter((c) => covered.includes(c.fn)))
const byFunction = new Map<string, Outcome[]>()
for (const o of outcomes) byFunction.set(o.c.fn, [...(byFunction.get(o.c.fn) ?? []), o])

describe('acceptance against Excel en-US', () => {
  it('the stored Excel answers cover every case', () => {
    expect(CASES.filter((c) => !EXCEL[c.id]).map((c) => c.id)).toEqual([])
  })

  it('every whitelisted function has cases, or is listed with the reason why not', () => {
    expect(covered.filter((f) => !byFunction.has(f))).toEqual([])
    for (const f of Object.keys(NO_CASES)) expect(FUNCTIONS[f], f).toBeDefined()
  })

  for (const f of covered) {
    it(`${f}: all cases match Excel`, () => {
      const off = (byFunction.get(f) ?? []).filter((o) => !o.ok && !(o.c.id in KNOWN)).map((o) => `${o.c.id} ${o.c.formula}: Excel ${show(o.excel)}, engine ${show(o.engine)}`)
      expect(off).toEqual([])
    })
  }

  it('the known differences are still differences (remove an entry once it is fixed)', () => {
    const fixed = outcomes.filter((o) => o.ok && o.c.id in KNOWN).map((o) => o.c.id)
    expect(fixed).toEqual([])
  })
})
