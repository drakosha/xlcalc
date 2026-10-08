/**
 * Case sets whose answers were predicted before Excel ran them (`PREDICTED` in `bench/coverage/sets/<set>.ts`): the
 * predictions that missed Excel's answer are listed in the set's `MISSED`, so the record of what was predicted stays
 * with the set. A set without stored answers yet is skipped.
 */
import { describe, expect, it } from 'vitest'
import type { Val } from '../bench/coverage/formulajs.js'
import { matches, show } from './excel-cases.js'
import { SETS } from './excel-sets.js'

const PREDICTING = ['areas-each', 'findate-edges', 'intersect', 'sheet-cells', 'single-values']

describe('case sets predicted before Excel ran them', async () => {
  for (const name of PREDICTING) {
    const mod = (await import(`../bench/coverage/sets/${name}.ts`)) as { PREDICTED: Readonly<Record<string, Val>>; MISSED?: readonly string[] }
    const set = SETS.find((s) => s.name === name)
    it.skipIf(!set)(`${name}: the predictions Excel did not confirm are the ones in MISSED`, () => {
      const missed = Object.entries(mod.PREDICTED)
        .filter(([id, v]) => !matches(set!.cases.find((c) => c.id === id)?.fn, set!.values[id], v))
        .map(([id, v]) => `${id}: predicted ${show(v)}, Excel ${show(set!.values[id])}`)
      expect(missed.map((m) => m.slice(0, m.indexOf(':')))).toEqual([...(mod.MISSED ?? [])])
      expect(Object.keys(mod.PREDICTED).length).toBeGreaterThan(0)
    })
  }
})
