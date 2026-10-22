/**
 * Spilling against Excel en-US (`test/excel/spill/<set>.json`, from `run.ts spill <set>`): each stored case is
 * its generator's current case, and every cell of its rectangles holds what Excel showed (blank or a value;
 * numbers as in the coverage measurement, a volatile case's numbers by type). Two arrays over one cell that do
 * not depend on each other: Excel's choice varies between runs (`order-runs.json` keeps every answer seen),
 * and any of them is accepted. Held-out cases (`held`) also keep the answer predicted before Excel ran.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { RectsAnswer } from '../bench/coverage/case-book.js'
import type { StoredSpill } from '../bench/coverage/spill.js'
import type { SpillCase } from '../bench/coverage/spill/cases.js'
import { computed, differences } from './spill-cases.js'
import { compare } from '../bench/coverage/compare.js'
import type { Val } from '../bench/coverage/formulajs.js'
import { currentCases } from './stored-cases.js'

const dir = new URL('./excel/spill/', import.meta.url)
const SETS = existsSync(dir)
  ? readdirSync(dir)
      .filter((f) => f.endsWith('.json') && !f.endsWith('-runs.json'))
      .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredSpill) }))
  : []
const RUNS = JSON.parse(readFileSync(new URL('order-runs.json', dir), 'utf8')) as { answers: Record<string, RectsAnswer[]> }

/** `<set>:<case>` → why the engine's answer differs from the one Excel run of the case. */
const KNOWN: Readonly<Record<string, string>> = {
  'base:cross': 'two independent arrays over one cell: in the set order the same pair (cross-C1-A2-row-first) let either spill',
  'base:cross-rect': 'two independent arrays over one cell: Excel lets either spill, the set order shows it for 11 pairs',
  'more:volatile-size': 'a size drawn by RANDBETWEEN: Excel spilled its own draw (five runs of five), the engine spills the draw of its clock',
  'more:volatile-sizes': 'ten sizes drawn in one book: Excel gave #SPILL! 47 times of 50 in five runs (the sizes changed between its passes), the engine draws once and spills',
  'more:volatile-cell': 'a size from a cell holding a draw: both spill that cell\'s value (Excel five runs of five), which Excel and the engine draw differently',
  'legacy-shown:params': "an error as NETWORKDAYS.INTL's weekend is #VALUE! in Excel and #NUM! here, XIRR's guess given a range #VALUE! there and lifted here",
  'legacy-shown:nesting': 'inside SUMPRODUCT Excel does not lift ACCRINT, CUMIPMT, PRICE, XIRR, GROWTH and TREND over an array in a one-value argument; the engine lifts them',
}

for (const set of SETS) {
  const built = ((await import(`../bench/coverage/spill/${set.name}.ts`)) as { buildSpillCases: () => SpillCase[] }).buildSpillCases()

  describe(`spilling against Excel en-US: set ${set.name}`, () => {
    currentCases(set, built)

    // a book Excel does not load is checked where its reason is (`test/names-excel.test.ts`)
    for (const c of set.cases.filter((x) => !set.answers[x.id]?.rejected)) {
      it(`${c.id}: ${c.about}`, () => {
        const r = computed(c)
        const seen = RUNS.answers[c.id] ?? [set.answers[c.id]]
        const diffs = seen.map((a) => differences(c, a, r))
        const best = diffs.find((d) => !d.length) ?? diffs[0]
        const known = KNOWN[`${set.name}:${c.id}`]
        if (known) expect(best.length, known).toBeGreaterThan(0)
        else expect(best).toEqual([])
      })
    }
  })
}

describe('held-out spill cases: the engine gives the answer predicted before Excel ran', async () => {
  for (const set of ['held', 'names-held', 'legacy-held', 'cse-held', 'at-held', 'cube-held']) {
    const held = ((await import(`../bench/coverage/spill/${set}.ts`)) as { buildSpillCases: () => SpillCase[] }).buildSpillCases()
    for (const c of held) it(`${set}: ${c.id}`, () => expect(differences(c, { values: c.predicted!, formulas: {} }, computed(c))).toEqual([]))
  }
})

describe('error-areas-held: the predictions Excel did not confirm are the ones in MISSED', async () => {
  const { buildSpillCases, MISSED } = (await import('../bench/coverage/spill/error-areas-held.js')) as { buildSpillCases: () => SpillCase[]; MISSED: readonly string[] }
  const set = SETS.find((s) => s.name === 'error-areas-held')
  it.skipIf(!set)('each case', () => {
    // a book Excel did not load misses whatever was predicted for it
    const same = (a: Record<string, Val>, b: Record<string, Val> = {}) => Object.keys({ ...a, ...b }).every((k) => a[k] !== undefined && b[k] !== undefined && compare(a[k], b[k]).ok)
    const missed = buildSpillCases().filter((c) => set!.answers[c.id].rejected || !same(set!.answers[c.id].values, c.predicted))
    expect(missed.map((c) => c.id)).toEqual([...MISSED])
  })
})
