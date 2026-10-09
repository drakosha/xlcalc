/**
 * Acceptance against Excel in each locale but en-US on the case sets computed in it (`test/excel/cases/<tag>/`). In
 * ru-RU: the locale set (text as numbers, dates and booleans, numbers as text, TEXT codes, DOLLAR, FIXED, NUMBERVALUE,
 * complex text, criteria), the order of text, criteria over text cells, AND/OR/XOR with words as text, and en-US sets
 * run again in ru-RU (unions, intersections, sheets named like cells, one value for a table, financial and date edges,
 * formulas inside the ranges they make). Each set's stored cases equal its generator's; known differences are pinned in
 * `KNOWN_BY_LOCALE` (`excel-locale-cases.ts`). A locale without sets is a skipped test that names the missing directory.
 */
import { describe, expect, it } from 'vitest'
import { describeOff, LOCALE_SETS, LOCALE_TAGS, runLocaleSet } from './excel-locale-cases.js'

for (const tag of LOCALE_TAGS) {
  const sets = LOCALE_SETS.get(tag)!
  if (!sets.length) {
    it.skip(`acceptance against Excel ${tag}: no case sets computed in Excel ${tag} (test/excel/cases/${tag}/)`)
    continue
  }

  for (const set of sets) {
    const { current, outcomes } = await runLocaleSet(tag, set)

    describe(`acceptance against Excel ${tag}: set ${set.name}`, () => {
      it('the stored Excel answers are for the current cases', () => expect(current).toBe(true))

      for (const group of [...new Set(outcomes.map((o) => o.group))])
        it(`${group}: all cases match Excel ${tag}`, () => {
          expect(outcomes.filter((o) => o.group === group && !o.ok).map(describeOff)).toEqual([])
        })
    })
  }

  describe(`${tag} case sets`, () => {
    const locale = sets.find((s) => s.name === 'locale')
    it.skipIf(!locale)('the locale set reaches every mechanism of docs/locale.md', () => {
      const groups = new Set(locale!.cases.map((c) => c.group))
      expect([...groups].sort()).toEqual(['boolean', 'complex', 'criteria', 'currency', 'date', 'format', 'month', 'number', 'numbervalue', 'text'])
    })
  })
}
