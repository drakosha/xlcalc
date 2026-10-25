import { expect, it } from 'vitest'
import { createState } from '../src/engine.js'
import { readersOf } from '../src/graph.js'
import { MAX_COLS, MAX_ROWS, formatAddress } from '../src/model.js'
import { RangeCache } from '../src/range-cache.js'
import { mulberry32 } from '../src/synth.js'
import { Rng } from '../src/values.js'

it('graph readers and cache invalidation agree with a full scan across narrow, wide and whole-sheet areas', () => {
  const rnd = mulberry32(51)
  const ranges = Array.from({ length: 180 }, (_, i) => {
    const r = Math.floor(rnd() * 600), c = Math.floor(rnd() * 120)
    const width = [1, 2, 64, 65, 200][i % 5]
    return new Rng(i % 2, r, c, r + Math.floor(rnd() * 100), c + width - 1)
  })
  ranges.push(new Rng(0, 0, 0, MAX_ROWS - 1, 0), new Rng(1, 0, 0, 0, MAX_COLS - 1))
  const cells = Object.fromEntries(ranges.map((r, i) => [`A${i + 1}`, { formula: `SUM(S${r.s}!${formatAddress(r.r1, r.c1)}:${formatAddress(r.r2, r.c2)})` }]))
  const state = createState({ sheets: [{ name: 'S0', cells: {} }, { name: 'S1', cells: {} }, { name: 'Q', cells }] }, { now: 0, seed: 1 })
  const cache = new RangeCache<number>()
  for (const [i, r] of ranges.entries()) cache.set(r.key, i)
  const remaining = new Map(ranges.map((r, i) => [r.key, i]))
  const queries = [
    new Rng(0, 800, 0, 900, 1), new Rng(1, 0, 800, 1, 900),
    ...Array.from({ length: 60 }, (_, i) => {
      const r = Math.floor(rnd() * 700), c = Math.floor(rnd() * 150)
      return new Rng(i % 2, r, c, r + 30, c + (i % 3 === 0 ? 300 : 1))
    }),
    new Rng(0, 0, 0, MAX_ROWS - 1, MAX_COLS - 1), new Rng(1, 0, 0, MAX_ROWS - 1, MAX_COLS - 1),
  ]
  for (const q of queries) {
    const hit = (r: Rng) => r.s === q.s && r.r1 <= q.r2 && r.r2 >= q.r1 && r.c1 <= q.c2 && r.c2 >= q.c1
    const found = new Set<number>()
    readersOf(state, q, found)
    const expected = state.formulas.filter((f) => f?.refs.some(hit)).map((f) => f!.id)
    expect([...found].sort((a, b) => a - b)).toEqual(expected.sort((a, b) => a - b))
    for (const r of ranges) if (hit(r)) remaining.delete(r.key)
    cache.invalidate(q)
    expect([...cache.keys()]).toEqual([...remaining.keys()])
  }
})
