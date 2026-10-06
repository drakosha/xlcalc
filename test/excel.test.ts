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
import { SET_FUNCTIONS } from './excel-sets.js'

/** Whitelisted functions without cases in the measurement or a case set, and why. */
const NO_CASES: Record<string, string> = {
  TODAY: 'volatile, excluded from the measurement; test/volatile.test.ts',
  NOW: 'volatile, excluded from the measurement; test/volatile.test.ts',
  RAND: 'volatile, excluded from the measurement; test/volatile.test.ts',
  RANDBETWEEN: 'volatile, excluded from the measurement; test/volatile.test.ts',
}

const covered = FUNCTION_NAMES.filter((f) => !(f in NO_CASES))
const outcomes = runExcelCases(CASES.filter((c) => covered.includes(c.fn)))
const byFunction = new Map<string, Outcome[]>()
for (const o of outcomes) byFunction.set(o.c.fn, [...(byFunction.get(o.c.fn) ?? []), o])

describe('acceptance against Excel en-US', () => {
  it('the stored Excel answers cover every case', () => {
    expect(CASES.filter((c) => !EXCEL[c.id]).map((c) => c.id)).toEqual([])
  })

  it('every whitelisted function has cases (measurement or a case set), or is listed with the reason why not', () => {
    expect(covered.filter((f) => !byFunction.has(f) && !SET_FUNCTIONS.has(f))).toEqual([])
    for (const f of Object.keys(NO_CASES)) expect(FUNCTIONS[f], f).toBeDefined()
    // a function that has cases now leaves the list, so the cases count for it
    expect(Object.keys(NO_CASES).filter((f) => CASES.some((c) => c.fn === f) || SET_FUNCTIONS.has(f))).toEqual([])
  })

  for (const f of covered) {
    it(`${f}: all cases match Excel`, () => {
      const off = (byFunction.get(f) ?? []).filter((o) => !o.ok).map((o) => `${o.c.id} ${o.c.formula}: Excel ${show(o.excel)}, engine ${show(o.engine)}`)
      expect(off).toEqual([])
    })
  }
})
