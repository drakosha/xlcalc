/**
 * Spilling against Excel en-US (`docs/excel-spill/<set>.json`, from `run.ts spill <set>`): each stored case is
 * its generator's current case, and every cell of its rectangles holds what Excel showed (blank or a value;
 * numbers as in the coverage measurement, a volatile case's numbers by type). Two arrays over one cell that do
 * not depend on each other: Excel's choice varies between runs (`order-runs.json` keeps every answer seen),
 * and any of them is accepted. Held-out cases (`held`) also keep the answer predicted before Excel ran.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { compare } from '../bench/coverage/compare.js'
import type { Val } from '../bench/coverage/formulajs.js'
import type { SpillAnswer, StoredSpill } from '../bench/coverage/spill.js'
import type { SpillCase } from '../bench/coverage/spill/cases.js'
import { compute, type ComputeResult } from '../src/index.js'
import { formatAddress, parseAddress } from '../src/model.js'
import { toVal } from './excel-cases.js'

const dir = new URL('../docs/excel-spill/', import.meta.url)
const SETS = existsSync(dir)
  ? readdirSync(dir)
      .filter((f) => f.endsWith('.json') && !f.endsWith('-runs.json'))
      .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredSpill) }))
  : []
const RUNS = JSON.parse(readFileSync(new URL('order-runs.json', dir), 'utf8')) as { answers: Record<string, SpillAnswer[]> }

/** `<set>:<case>` → why the engine's answer differs from the one Excel run of the case. */
const KNOWN: Readonly<Record<string, string>> = {
  'base:cross': 'two independent arrays over one cell: in the set order the same pair (cross-C1-A2-row-first) let either spill',
  'base:cross-rect': 'two independent arrays over one cell: Excel lets either spill, the set order shows it for 11 pairs',
}

/** What a cell of the result shows: its value, a large array's unstored cells their fill, else blank. */
export function shownValue(r: ComputeResult, sheet: string, addr: string): Val | null {
  const cells = r.sheets.find((s) => s.name === sheet)!.cells
  const c = cells[addr]
  if (c) return c.value === null ? null : toVal(c.value)
  const { row, col } = parseAddress(addr)
  for (const x of Object.values(cells)) {
    if (x.spillFill === undefined) continue
    const [p, q] = x.spill!.split(':').map(parseAddress)
    if (row >= p.row && row <= q.row && col >= p.col && col <= q.col) return toVal(x.spillFill)
  }
  return null
}

/** Cells of the case's rectangles where the engine does not show Excel's answer. */
function differences(c: SpillCase, a: SpillAnswer, r: ComputeResult): string[] {
  const out: string[] = []
  for (const g of c.read) {
    const [p, q = p] = g.range.split(':').map(parseAddress)
    for (let row = p.row; row <= q.row; row++)
      for (let col = p.col; col <= q.col; col++) {
        const addr = formatAddress(row, col)
        const excel = a.values[`${g.sheet}!${addr}`] ?? null
        const engine = shownValue(r, g.sheet, addr)
        const ok = excel === null || engine === null ? excel === engine : c.volatile && excel.t === 'n' ? engine.t === 'n' : compare(excel, engine).ok
        if (!ok) out.push(`${g.sheet}!${addr}: Excel ${JSON.stringify(excel)}, engine ${JSON.stringify(engine)}`)
      }
  }
  return out
}

for (const set of SETS) {
  const built = ((await import(`../bench/coverage/spill/${set.name}.ts`)) as { buildSpillCases: () => SpillCase[] }).buildSpillCases()

  describe(`spilling against Excel en-US: set ${set.name}`, () => {
    it('the stored cases are the current ones, all answered', () => {
      expect(set.cases).toEqual(built)
      expect(built.filter((c) => !set.answers[c.id])).toEqual([])
    })

    for (const c of set.cases) {
      it(`${c.id}: ${c.about}`, () => {
        const r = compute(c.book, { clock: { now: '2026-10-06T12:00:00Z', seed: 1 } })
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
  const held = ((await import('../bench/coverage/spill/held.js')) as { buildSpillCases: () => SpillCase[] }).buildSpillCases()
  for (const c of held) it(c.id, () => expect(differences(c, { values: c.predicted!, formulas: {} }, compute(c.book))).toEqual([]))
})
