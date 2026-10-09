/**
 * Case sets whose answers were predicted before Excel ran them (`PREDICTED` in `bench/coverage/sets/<set>.ts`, or in
 * `bench/coverage/legacy/<set>.ts` for formulas of the old kind): the predictions that missed Excel's answer are
 * listed in the set's `MISSED`, so the record of what was predicted stays with the set. A set without stored answers
 * yet is skipped.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { Val } from '../bench/coverage/formulajs.js'
import { matches, show } from './excel-cases.js'
import { SETS } from './excel-sets.js'

const PREDICTING = ['areas-each', 'findate-edges', 'intersect', 'known-held', 'ref-deps', 'sheet-cells', 'single-values']
const LEGACY_PREDICTING = ['ifs-text']

/** A legacy set's stored answers (the converted formula's value is predicted), if Excel ran it. */
function legacySet(name: string): { cases: { id: string; fn?: string }[]; values: Record<string, Val> } | undefined {
  const file = new URL(`./excel/legacy/${name}.json`, import.meta.url)
  return existsSync(file) ? JSON.parse(readFileSync(file, 'utf8')) : undefined
}

describe('case sets predicted before Excel ran them', async () => {
  for (const [name, dir] of [...PREDICTING.map((n) => [n, 'sets']), ...LEGACY_PREDICTING.map((n) => [n, 'legacy'])]) {
    const mod = (await import(`../bench/coverage/${dir}/${name}.ts`)) as { PREDICTED: Readonly<Record<string, Val>>; MISSED?: readonly string[] }
    const set = dir === 'sets' ? SETS.find((s) => s.name === name) : legacySet(name)
    it.skipIf(!set)(`${name}: the predictions Excel did not confirm are the ones in MISSED`, () => {
      const missed = Object.entries(mod.PREDICTED)
        .filter(([id, v]) => !matches(set!.cases.find((c) => c.id === id)?.fn, set!.values[id], v))
        .map(([id, v]) => `${id}: predicted ${show(v)}, Excel ${show(set!.values[id])}`)
      expect(missed.map((m) => m.slice(0, m.indexOf(':')))).toEqual([...(mod.MISSED ?? [])])
      expect(Object.keys(mod.PREDICTED).length).toBeGreaterThan(0)
    })
  }
})
