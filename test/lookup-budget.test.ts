/**
 * Exact lookups count their indexes in the cache budget (`docs/cache.md`). A book of 34,000 unique text keys and 1,000
 * exact lookups of them by XLOOKUP, MATCH or XMATCH over the column: once it is computed, the budget counts at least
 * the estimate of the column's index (`storedBytes`), which goes with the column's kept values. Under a budget of 4 KB,
 * which keeps neither, the lookups give the same values.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, CellInput } from '../src/index.js'
import { createState } from '../src/engine.js'
import { storedBytes } from '../src/functions/lookup/exact.js'
import { CacheBudget } from '../src/range-cache.js'

const key = (r: number) => `Ключ номер ${(r * 7919) % 100003}`

/** Sheet S: `n` unique keys in A with numbers in B, `lookups` keys in D and their lookups by `fn` in E. */
function book(fn: string, n: number, lookups: number): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= n; r++) {
    cells[`A${r}`] = { value: key(r) }
    cells[`B${r}`] = { value: r }
  }
  for (let r = 1; r <= lookups; r++) {
    cells[`D${r}`] = { value: key(((r * 31) % n) + 1) }
    cells[`E${r}`] = { formula: fn === 'XLOOKUP' ? `XLOOKUP(D${r},$A$1:$A$${n},$B$1:$B$${n})` : `${fn}(D${r},$A$1:$A$${n},0)` }
  }
  return { sheets: [{ name: 'S', cells }] }
}

describe.each(['XLOOKUP', 'MATCH', 'XMATCH'])('%s', (fn) => {
  it("counts its index over the column's kept values in the cache budget", () => {
    const n = 34_000
    const cache = new CacheBudget()
    createState(book(fn, n, 1000), { now: 0, seed: 1, cache })
    const text = Array.from({ length: n }, (_, i) => ({ value: key(i + 1), first: i, last: i }))
    expect(cache.held).toBeGreaterThanOrEqual(storedBytes({ text, scalarWork: 0, textWork: 0 }))
  })

  it('gives the same values under a budget of 4 KB', () => {
    const b = book(fn, 3400, 100)
    const [own, small] = [new CacheBudget(), new CacheBudget(4096)].map((cache) => createState(b, { now: 0, seed: 1, cache }))
    const values = (s: typeof own) => Array.from({ length: 100 }, (_, r) => s.rt.cell(0, r, 4))
    expect(values(small)).toEqual(values(own))
    expect(values(own).every((v) => typeof v === 'number')).toBe(true)
  })
})
