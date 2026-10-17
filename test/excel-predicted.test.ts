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
import { LOCALE_SETS } from './excel-locale-cases.js'
import { SETS } from './excel-sets.js'

const PREDICTING = ['accrint-ironcalc-held', 'areas-each', 'binom-sumx-held', 'countif-ironcalc-held', 'currency-held', 'database-ironcalc-held', 'date-ironcalc-held', 'euro-ironcalc-held', 'euro-ironcalc-held-2', 'euro-ironcalc-held-3', 'finance-ironcalc-held', 'finance-ironcalc-held-2', 'findate-edges', 'fit-offset-held', 'intersect', 'known-held', 'lookup-held', 'lookup-own-held', 'math-ironcalc-held', 'name-case-held', 'near-zero-held', 'pct-text', 'power-bound-held', 'power-edges-held', 'power-neg-held', 'ref-deps', 'refs-ironcalc-held', 'refs-ironcalc-held-2', 'rounding-ironcalc-held', 'seconds-held', 'sheet-cells', 'single-values', 'sortby-one-held', 'stat-edges-held', 'sum-zero-held', 'sumproduct-held', 'sumx-cell-held', 'tests-cell-held', 'text-date-ironcalc-held', 'text-date-ironcalc-held-2', 'text-date-ironcalc-held-3', 'text-date-ironcalc-held-4', 'textafter-ironcalc-held', 'textafter-ironcalc-held-2', 'xlookup-ironcalc-held', 'xlookup-ironcalc-held-2', 'xmatch-regex-held', 'xmatch-regex-held-2']
const LEGACY_PREDICTING = ['ifs-text']
const RU_PREDICTING = ['areas-each', 'findate-edges', 'intersect', 'pct-text', 'ref-deps', 'sheet-cells', 'single-values']
const RU_SETS = LOCALE_SETS.get('ru-RU') ?? []

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
      expect(missed.map((m) => m.slice(0, m.indexOf(': predicted')))).toEqual([...(mod.MISSED ?? [])])
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
