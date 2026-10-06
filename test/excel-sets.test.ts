/**
 * Acceptance against Excel for the case sets in `docs/excel-cases/` (see `test/excel-sets.ts`): each set's
 * stored cases equal what its generator builds now, and the engine matches Excel on every case or gives the
 * answer pinned in KNOWN (`test/excel-cases.ts`). Comparison as in the coverage measurement.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { compute } from '../src/index.js'
import { CASES, CLOCK, KNOWN, REASONS, show, toVal, verdict } from './excel-cases.js'
import { OPEN } from './excel-open.js'
import { caseBook, loadSet, SETS } from './excel-sets.js'

for (const set of SETS) {
  const mod = await loadSet(set.name)
  const probes = mod.buildProbes()
  const result = compute(caseBook(probes, mod.SHEETS), { clock: CLOCK }).sheets.at(-1)!.cells

  describe(`acceptance against Excel en-US: set ${set.name}`, () => {
    it('the stored Excel answers are for the current cases', () => {
      expect(set.cases).toEqual(probes)
      expect(probes.filter((p) => !set.values[p.id])).toEqual([])
    })

    for (const group of [...new Set(probes.map((p) => p.fn ?? p.group))]) {
      it(`${group}: all cases match Excel`, () => {
        const off = probes
          .map((p, i) => ({ p, engine: toVal(result[`A${i + 1}`].value), excel: set.values[p.id] }))
          .filter((o) => (o.p.fn ?? o.p.group) === group)
          .map((o) => ({ ...o, v: verdict(`${set.name}:${o.p.id}`, o.p.fn, o.excel, o.engine) }))
          .filter((o) => !o.v.ok)
          .map((o) => `${o.p.id} ${o.p.formula}: Excel ${show(o.excel)}, engine ${show(o.engine)}${o.v.why ? ` (${o.v.why})` : ''}`)
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

describe('known differences from Excel', () => {
  it('every KNOWN and OPEN entry names an existing case', () => {
    const keys = new Set([...CASES.map((c) => `measurement:${c.id}`), ...SETS.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`))])
    expect([...Object.keys(KNOWN), ...Object.keys(OPEN)].filter((k) => !keys.has(k))).toEqual([])
  })

  it("the README's known differences have a line for every reason in KNOWN", () => {
    const readme = readFileSync(new URL('../README.md', import.meta.url), 'utf8')
    const section = readme.slice(readme.indexOf('- **Known differences from Excel:**'), readme.indexOf('- **Syntax.**'))
    const used = [...new Set(Object.values(KNOWN).map((k) => k.reason))]
    expect(used.filter((r) => !section.includes(`  - ${REASONS[r]}\n`))).toEqual([])
    expect(Object.keys(REASONS).filter((r) => !used.includes(r as keyof typeof REASONS))).toEqual([])
  })
})
