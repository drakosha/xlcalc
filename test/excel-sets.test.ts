/**
 * Acceptance against Excel for the case sets in `docs/excel-cases/` (see `test/excel-sets.ts`): each set's
 * stored cases equal what its generator builds now, and the engine matches Excel on every case. Comparison
 * as in the coverage measurement.
 */
import { describe, expect, it } from 'vitest'
import type { Probe } from '../bench/coverage/probes.js'
import { compute, type CellInput } from '../src/index.js'
import { dataSheet, engineFormula, matches, show, toVal } from './excel-cases.js'
import { SETS } from './excel-sets.js'

type Sheets = Record<string, Record<string, number | string | boolean>>

for (const set of SETS) {
  const mod = (await import(`../bench/coverage/sets/${set.name}.ts`)) as { buildProbes: () => Probe[]; SHEETS?: Sheets }
  const probes = mod.buildProbes()
  const sheets = Object.entries(mod.SHEETS ?? {}).map(([name, cells]) => ({ name, cells: Object.fromEntries(Object.entries(cells).map(([a, v]): [string, CellInput] => [a, { value: v }])) }))
  const cases = { name: 'C', cells: Object.fromEntries(probes.map((p, i): [string, CellInput] => [`A${i + 1}`, { formula: engineFormula(p.formula) }])) }
  // as the book sent to Excel: the measurement's data sheet D, the set's own sheets, the cases on C
  const result = compute({ sheets: [dataSheet(), ...sheets, cases] }, { clock: { now: '2026-10-05T12:00:00Z', seed: 1 } }).sheets.at(-1)!.cells

  describe(`acceptance against Excel en-US: set ${set.name}`, () => {
    it('the stored Excel answers are for the current cases', () => {
      expect(set.cases).toEqual(probes)
      expect(probes.filter((p) => !set.values[p.id])).toEqual([])
    })

    for (const group of [...new Set(probes.map((p) => p.fn ?? p.group))]) {
      it(`${group}: all cases match Excel`, () => {
        const off = probes
          .map((p, i) => ({ p, engine: toVal(result[`A${i + 1}`].value), excel: set.values[p.id] }))
          .filter((o) => (o.p.fn ?? o.p.group) === group && !matches(o.p.fn, o.excel, o.engine))
          .map((o) => `${o.p.id} ${o.p.formula}: Excel ${show(o.excel)}, engine ${show(o.engine)}`)
        expect(off).toEqual([])
      })
    }
  })
}

describe('excel case sets', () => {
  it('every set has a generator and unique case ids', () => {
    for (const set of SETS) expect(new Set(set.cases.map((c) => c.id)).size, set.name).toBe(set.cases.length)
  })
})
