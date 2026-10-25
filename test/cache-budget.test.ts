/**
 * The budget of the range caches (`docs/cache.md`): a range's values and criteria index are kept from its second read
 * on, the books of one heap share what the caches hold, and past the cap the least recently read entries go between
 * formulas. What goes is read again with the same values: sessions under a small budget give the values and hash of a
 * full computation through edits, sheet growth and structure operations, in both execution modes.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput, type CellInput, type Session } from '../src/index.js'
import { createState } from '../src/engine.js'
import { CacheBudget, heapCaches, RangeCache } from '../src/range-cache.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const SMALL = 32 * 1024

/** Runs `f` with this heap's budget cut to `SMALL`, so that each few formulas drop what the ones before read. */
function tight<T>(f: () => T): T {
  const cap = heapCaches.cap
  heapCaches.cap = SMALL
  try {
    return f()
  } finally {
    heapCaches.cap = cap
  }
}

/** Running means read twice by each formula, criteria and lookups over the column, whole columns, a rectangle wider than its data. */
function book(n: number, own = 0, materialized = false): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= n; r++) {
    cells[`A${r}`] = { value: r % 7 === 0 ? `t${(r + own) % 5}` : (r + own) % 11 }
    cells[`B${r}`] = { formula: `AVERAGE($A$1:A${r}${materialized ? `,$A$1:A${r}` : ''})+COUNTA($A$1:A${r})` }
  }
  for (let k = 1; k <= 12; k++) {
    cells[`D${k}`] = { formula: `COUNTIF(A:A,${k % 11})+SUMIFS($A$1:$A$${n},$A$1:$A$${n},${k % 11})+COUNTIFS($A$1:$A$${n},"t${k % 5}")` }
    cells[`E${k}`] = { formula: `IFERROR(MATCH(${k},$A$1:$A$${n},0),0)+INDEX($A$1:$J$${n}*1,${k * 17},1)+XLOOKUP(${k % 11},A:A,B:B,0)` }
  }
  return { sheets: [{ name: 'S', cells }] }
}

it('keeps a value from its second read and drops the least recently read past the cap, never what the last formula read', () => {
  const budget = new CacheBudget(10_000)
  const c = new RangeCache<number[]>(budget)
  const [a, b, d, e] = ['0:0:0:299:0', '0:0:1:299:1', '0:0:2:299:2', '0:0:3:299:3']
  expect(c.admits(a)).toBe(false)
  expect(c.get(a)).toBeUndefined()
  expect(c.admits(a)).toBe(true)
  for (const k of [a, b, d]) c.set(k, [1], 3000)
  budget.fit()
  expect(c.get(a)).toEqual([1])
  budget.fit()
  c.set(e, [2], 3000)
  expect(budget.held).toBeGreaterThan(10_000)
  budget.fit()
  expect([...c.keys()].sort()).toEqual([a, e].sort())
  expect(budget.held).toBe(c.held)
  expect(budget.held).toBeLessThanOrEqual(7_500)
  // one read past the cap stays while the formula that read it is the last one
  c.set(b, [3], 20_000)
  budget.fit()
  expect(c.get(b)).toEqual([3])
  c.clear()
  expect(budget.held).toBe(0)
})

it('a range read once fills no cache, a range read twice is kept', () => {
  const once: Record<string, CellInput> = {}
  const twice: Record<string, CellInput> = {}
  for (let r = 1; r <= 600; r++) {
    once[`A${r}`] = twice[`A${r}`] = { value: r }
    once[`B${r}`] = { formula: `SUMPRODUCT($A$1:A${r})` }
    twice[`B${r}`] = { formula: `SUMPRODUCT($A$1:A${r})+MEDIAN($A$1:A${r})` }
  }
  const cache = new CacheBudget(Infinity)
  expect(createState({ sheets: [{ name: 'S', cells: once }] }, { now: 0, seed: 1, cache }).ctx.flat.size).toBe(0)
  expect(createState({ sheets: [{ name: 'S', cells: twice }] }, { now: 0, seed: 1, cache }).ctx.flat.size).toBe(600 - 255)
})

it('stored keys preserve Map order and stored undefined while first-read marks stay out of iteration', () => {
  const c = new RangeCache<number | undefined>()
  const [a, b, mark] = ['0:0:0:9:0', '0:0:1:9:1', '0:0:2:9:2']
  const expected = new Map<string, number | undefined>()
  c.admits(a); c.admits(mark)
  expect(c.size).toBe(0)
  expect(c.has(a)).toBe(false)
  for (const [k, value] of [[b, undefined], [a, 3], [b, 4]] as const) { c.set(k, value); expected.set(k, value) }
  expect(c.size).toBe(expected.size)
  expect(c.has(b)).toBe(true)
  expect([...c.keys()]).toEqual([...expected.keys()])
  expect(c.get(b)).toBe(4)
  expect(c.get(a)).toBe(3)
  // a stored undefined is a value, not a miss, as Map.set kept it
  c.set(mark, undefined)
  expect(c.has(mark)).toBe(true)
  expected.set(mark, undefined)
  expect([...c.keys()]).toEqual([...expected.keys()])
  expect(c.delete(mark)).toBe(true)
  expected.delete(mark)
  expect(c.delete(b)).toBe(true)
  expected.delete(b)
  expect([...c.keys()]).toEqual([...expected.keys()])
  c.clear()
  expect(c.size).toBe(0)
  expect(c.held).toBe(0)
})

describe.each(['closures', 'eval'] as const)('cache budget, %s', (execution) => {
  const opts = { clock, execution, display: false }
  const same = (s: Session) => expect(s.getHash()).toBe(compute(s.getBook(), opts).hash)

  it('a book whose caches hold many times the cap computes as without it', () => {
    const b = book(600, 0, true)
    const free = createState(b, { now: 46000, seed: 7, execution, cache: new CacheBudget(Infinity) })
    const held = createState(b, { now: 46000, seed: 7, execution, cache: new CacheBudget(SMALL) })
    expect(free.ctx.cache.held).toBeGreaterThan(20 * SMALL)
    expect(held.ctx.cache.held).toBeLessThanOrEqual(SMALL)
    expect(tight(() => compute(b, opts)).hash).toBe(compute(b, opts).hash)
  })

  it('a session under a small budget gives the full computation’s hash through edits, growth and structure operations', () => {
    const n = 600
    const s = tight(() => createSession(book(n), opts))
    try {
      expect(heapCaches.held).toBeLessThanOrEqual(SMALL)
      same(s)
      const steps: (() => unknown)[] = [
        () => s.setCells([{ sheet: 'S', cell: 'A5', input: { value: 100 } }]),
        // the sheet grows wider and taller outside the ranges: a range read again after it was dropped takes the new shape
        () => s.setCells([{ sheet: 'S', cell: `I${n + 500}`, input: { value: 1 } }]),
        () => s.setCells([{ sheet: 'S', cell: 'D1', input: { value: 0 } }, { sheet: 'S', cell: 'D1', input: { formula: `COUNTIF(A:A,3)+SUMIFS($A$1:$A$${n},$A$1:$A$${n},3)` } }]),
        () => s.setCells([{ sheet: 'S', cell: `A${n / 2}`, input: { value: 't1' } }]),
        () => s.editStructure({ op: 'insertRows', sheet: 'S', row: 10, count: 2 }),
        () => s.setCells([{ sheet: 'S', cell: `A${n + 1}`, input: { value: 3 } }]),
        () => s.setCells([{ sheet: 'S', cell: 'A1', input: null }]),
      ]
      for (const step of steps) {
        tight(step)
        same(s)
      }
    } finally {
      s.close()
    }
  })

  it('books share the budget: a session’s formulas drop the entries of another, which reads them again', () => {
    const [s, t] = tight(() => [createSession(book(400), opts), createSession(book(400, 3), opts)])
    try {
      for (let i = 0; i < 6; i++) {
        const x = i % 2 ? s : t
        tight(() => x.setCells([{ sheet: 'S', cell: `A${1 + i * 37}`, input: { value: i } }]))
        same(x)
      }
      same(s)
      same(t)
    } finally {
      s.close()
      t.close()
    }
  })
})
