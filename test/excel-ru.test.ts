/**
 * Acceptance against Excel ru-RU on the coverage measurement: every whitelisted function's cases computed in a
 * ru-RU book and compared with `docs/coverage-formulajs/excel-ru.json`, as `excel.test.ts` does for en-US.
 * Known differences are pinned in `RU_KNOWN` (`excel-ru-cases.ts`).
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { FUNCTION_NAMES } from '../src/index.js'
import { CASES } from './excel-cases.js'
import { describeOff, reasonText, RU_EXCEL, RU_KNOWN, RU_SETS, runRuCases, type RuOutcome } from './excel-ru-cases.js'

const outcomes = runRuCases(CASES.filter((c) => FUNCTION_NAMES.includes(c.fn)))
const byFunction = new Map<string, RuOutcome[]>()
for (const o of outcomes) byFunction.set(o.fn!, [...(byFunction.get(o.fn!) ?? []), o])

describe('acceptance against Excel ru-RU: the coverage measurement', () => {
  it('the stored Excel ru-RU answers cover every case', () => {
    expect(CASES.filter((c) => !RU_EXCEL[c.id]).map((c) => c.id)).toEqual([])
  })

  for (const [f, os] of byFunction)
    it(`${f}: all cases match Excel ru-RU`, () => {
      expect(os.filter((o) => !o.ok).map(describeOff)).toEqual([])
    })
})

describe('known differences from Excel ru-RU', () => {
  it('every RU_KNOWN entry names an existing case', () => {
    const keys = new Set([...CASES.map((c) => `measurement:${c.id}`), ...RU_SETS.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`))])
    expect(Object.keys(RU_KNOWN).filter((k) => !keys.has(k))).toEqual([])
  })

  it("the README's known differences have a line for every reason in RU_KNOWN", () => {
    const readme = readFileSync(new URL('../README.md', import.meta.url), 'utf8')
    const section = readme.slice(readme.indexOf('- **Known differences from Excel:**'), readme.indexOf('- **Syntax.**'))
    const used = [...new Set(Object.values(RU_KNOWN).map((k) => k.reason))]
    expect(used.filter((r) => !section.includes(`  - ${reasonText(r)}\n`))).toEqual([])
  })
})
