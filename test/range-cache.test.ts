import { expect, it } from 'vitest'
import { createState } from '../src/engine.js'
import { invalidate } from '../src/evaluate.js'
import { MAX_COLS, MAX_ROWS } from '../src/model.js'
import { Rng } from '../src/values.js'

it('invalidates intersecting ranges across sheets, narrow columns and wide rectangles', () => {
  const s = createState({ sheets: [{ name: 'S', cells: {} }, { name: 'T', cells: {} }] }, { now: 0, seed: 1 })
  const ranges = [new Rng(0, 0, 0, 599, 0), new Rng(0, 0, 0, MAX_ROWS - 1, 0), new Rng(0, 0, 1, 599, 2), new Rng(1, 0, 0, 599, 0), new Rng(0, 9, 0, 19, MAX_COLS - 1)]
  for (const r of ranges) { s.ctx.flat.set(r.key, [1]); s.ctx.index.set(r.key, new Map()) }
  invalidate(s.ctx, { s: 0, r1: 999, r2: 999, c1: 0, c2: 0 })
  expect([...s.ctx.flat.keys()]).toEqual(ranges.filter((_, i) => i !== 1).map(r => r.key))
  invalidate(s.ctx, { s: 0, r1: 10, r2: 10, c1: 0, c2: 2 })
  expect([...s.ctx.flat.keys()]).toEqual([ranges[3].key])
  expect([...s.ctx.index.keys()]).toEqual([ranges[3].key])
  invalidate(s.ctx, { s: 1, r1: 0, r2: MAX_ROWS - 1, c1: 0, c2: MAX_COLS - 1 })
  expect(s.ctx.flat.size).toBe(0)
  expect(s.ctx.index.size).toBe(0)
  s.ctx.flat.set(ranges[0].key, [1]); s.ctx.flat.clear(); s.ctx.flat.set(ranges[1].key, [2])
  invalidate(s.ctx, { s: 0, r1: 800, r2: 800, c1: 0, c2: 0 })
  expect(s.ctx.flat.size).toBe(0)
})
