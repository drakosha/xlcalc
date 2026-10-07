/**
 * Spilling against Excel en-US (`test/excel/spill/<set>.json`, from `run.ts spill <set>`): each stored case is
 * its generator's current case, and every cell of its rectangles holds what Excel showed (blank or a value;
 * numbers as in the coverage measurement, a volatile case's numbers by type). Two arrays over one cell that do
 * not depend on each other: Excel's choice varies between runs (`order-runs.json` keeps every answer seen),
 * and any of them is accepted. Held-out cases (`held`) also keep the answer predicted before Excel ran.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { compare } from '../bench/coverage/compare.js'
import type { SpillAnswer, StoredSpill } from '../bench/coverage/spill.js'
import type { SpillCase } from '../bench/coverage/spill/cases.js'
import { compute, fromLegacy, InputError, openSession, type BookInput, type ComputeResult } from '../src/index.js'
import { formatAddress, parseAddress } from '../src/model.js'
import { shownValue } from './spill-shown.js'

const dir = new URL('./excel/spill/', import.meta.url)
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
  'more:volatile-size': 'a size drawn by RANDBETWEEN: Excel spilled its own draw (five runs of five), the engine spills the draw of its clock',
  'more:volatile-sizes': 'ten sizes drawn in one book: Excel gave #SPILL! 47 times of 50 in five runs (the sizes changed between its passes), the engine draws once and spills',
  'more:volatile-cell': 'a size from a cell holding a draw: both spill that cell\'s value (Excel five runs of five), which Excel and the engine draw differently',
}

const CLOCK = { now: '2026-10-06T12:00:00Z', seed: 1 }

/** The case's book as a reader loads it: the formulas a file stores as of the old kind converted (`fromLegacy`). */
function loaded(c: SpillCase): BookInput {
  if (!c.legacy) return c.book
  const old = new Set(c.legacy)
  return {
    ...c.book,
    sheets: c.book.sheets.map((sh) => ({
      ...sh,
      cells: Object.fromEntries(Object.entries(sh.cells).map(([a, x]) => [a, old.has(`${sh.name}!${a}`) && x.formula !== undefined ? { ...x, formula: fromLegacy(x.formula, { names: c.book.names, sheet: sh.name }) } : x])),
    })),
  }
}

/**
 * The case computed, and its edits applied one by one in a session, as Excel had them typed after computing: an edit
 * of a range edits each of its cells, an edit Excel refuses (part of an old-kind array formula) leaves the book as it was.
 */
function computed(c: SpillCase): ComputeResult {
  if (!c.edits) return compute(loaded(c), { clock: CLOCK })
  const s = openSession(loaded(c), { clock: CLOCK })
  for (const e of c.edits) {
    const [p, q = p] = e.cell.split(':').map(parseAddress)
    const input = e.value === null ? null : typeof e.value === 'object' ? e.value : { value: e.value }
    const edits = []
    for (let row = p.row; row <= q.row; row++) for (let col = p.col; col <= q.col; col++) edits.push({ sheet: e.sheet, cell: formatAddress(row, col), input })
    try {
      s.setCells(edits)
    } catch (err) {
      if (!(err instanceof InputError)) throw err
    }
  }
  return s.result()
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
  const held = ((await import('../bench/coverage/spill/held.js')) as { buildSpillCases: () => SpillCase[] }).buildSpillCases()
  for (const c of held) it(c.id, () => expect(differences(c, { values: c.predicted!, formulas: {} }, compute(c.book))).toEqual([]))
})
