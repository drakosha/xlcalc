/**
 * Spilling against Excel for Windows (`test/excel/win/spill/`, `bench/windows`): every cell of a case's rectangles holds
 * what Excel for Windows showed, as `spill-excel.test.ts` checks the Mac's. Where the engine knowingly differs,
 * `WIN_SPILL` gives the reason; such a case must still differ.
 */
import { describe, expect, it } from 'vitest'
import type { StoredSpill } from '../bench/coverage/spill.js'
import type { SpillCase } from '../bench/coverage/spill/cases.js'
import { winFiles } from './excel-win-cases.js'
import { computed, differences } from './spill-cases.js'

const SETS = winFiles<StoredSpill>('spill/')

/** `<set>:<case>` → why the engine's answer differs from Excel for Windows's. */
const WIN_SPILL: Readonly<Record<string, string>> = {}

for (const set of SETS) {
  const built = ((await import(`../bench/coverage/spill/${set.name}.ts`)) as { buildSpillCases: () => SpillCase[] }).buildSpillCases()

  describe(`spilling against Excel for Windows: set ${set.name}`, () => {
    it('the stored cases are the current ones, all answered', () => {
      expect(set.cases).toEqual(built)
      expect(built.filter((c) => !set.answers[c.id])).toEqual([])
    })

    for (const c of set.cases.filter((x) => !set.answers[x.id]?.rejected))
      it(`${c.id}: ${c.about}`, () => {
        const diffs = differences(c, set.answers[c.id], computed(c))
        const known = WIN_SPILL[`${set.name}:${c.id}`]
        if (known) expect(diffs.length, known).toBeGreaterThan(0)
        else expect(diffs).toEqual([])
      })
  })
}

describe('known differences from Excel for Windows: spilling', () => {
  it('every WIN_SPILL entry names a stored case', () => {
    const keys = new Set(SETS.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`)))
    expect(Object.keys(WIN_SPILL).filter((k) => !keys.has(k))).toEqual([])
  })
})
