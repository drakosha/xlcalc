/**
 * Acceptance against Excel ru-RU on the case sets computed in ru-RU (`test/excel/cases/ru-RU/`): the locale set
 * (text as numbers, dates and booleans, numbers as text, TEXT codes, DOLLAR, FIXED, NUMBERVALUE, complex text,
 * criteria), the order of text, criteria over text cells, AND/OR/XOR with words as text. Each set's stored cases
 * equal its generator's; known differences are pinned in `RU_KNOWN` (`excel-ru-cases.ts`).
 */
import { describe, expect, it } from 'vitest'
import { describeOff, RU_SETS, runRuSet } from './excel-ru-cases.js'

for (const set of RU_SETS) {
  const { current, outcomes } = await runRuSet(set)

  describe(`acceptance against Excel ru-RU: set ${set.name}`, () => {
    it('the stored Excel answers are for the current cases', () => expect(current).toBe(true))

    for (const group of [...new Set(outcomes.map((o) => o.group))])
      it(`${group}: all cases match Excel ru-RU`, () => {
        expect(outcomes.filter((o) => o.group === group && !o.ok).map(describeOff)).toEqual([])
      })
  })
}

describe('ru-RU case sets', () => {
  it('the locale set reaches every mechanism of docs/locale.md', () => {
    const groups = new Set(RU_SETS.find((s) => s.name === 'locale')!.cases.map((c) => c.group))
    expect([...groups].sort()).toEqual(['boolean', 'complex', 'criteria', 'currency', 'date', 'format', 'month', 'number', 'numbervalue', 'text'])
  })
})
