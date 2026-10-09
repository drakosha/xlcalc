/**
 * Case sets whose answers were predicted before Excel ran them (`PREDICTED` in `bench/coverage/sets/<set>.ts`, or in
 * `bench/coverage/legacy/<set>.ts` for formulas of the old kind; in ru-RU `sets/<set>.ru.predicted.json`, written by
 * `sets/predict-ru.ts`): the predictions that missed Excel's answer are listed in the set's `MISSED` (`RU_MISSED`), so
 * the record of what was predicted stays with the set. A set without stored answers yet is skipped.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { Val } from '../bench/coverage/formulajs.js'
import { matches, show } from './excel-cases.js'
import { RU_SETS } from './excel-ru-cases.js'
import { SETS } from './excel-sets.js'

const PREDICTING = ['areas-each', 'findate-edges', 'intersect', 'known-held', 'ref-deps', 'sheet-cells', 'single-values']
const LEGACY_PREDICTING = ['ifs-text']
const RU_PREDICTING = ['areas-each', 'findate-edges', 'intersect', 'ref-deps', 'sheet-cells', 'single-values']

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

describe('case sets predicted in ru-RU before Excel ru-RU ran them', async () => {
  for (const name of RU_PREDICTING) {
    const predicted = JSON.parse(readFileSync(new URL(`../bench/coverage/sets/${name}.ru.predicted.json`, import.meta.url), 'utf8')) as Record<string, Val>
    const { RU_MISSED = [] } = (await import(`../bench/coverage/sets/${name}.ts`)) as { RU_MISSED?: readonly string[] }
    const set = RU_SETS.find((s) => s.name === name)
    it.skipIf(!set)(`${name}: the predictions Excel ru-RU did not confirm are the ones in RU_MISSED`, () => {
      expect(Object.keys(predicted)).toEqual(set!.cases.map((c) => c.id))
      const missed = Object.entries(predicted).filter(([id, v]) => !matches(set!.cases.find((c) => c.id === id)?.fn, set!.values[id], v))
      expect(missed.map(([id]) => id)).toEqual([...RU_MISSED])
    })
  }
})
