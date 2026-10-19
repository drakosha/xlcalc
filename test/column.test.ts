/**
 * A sheet's column of values (`src/column.ts`): chunks of 1,024 rows that start at their first row written, read and
 * written in any order against a plain map of the same writes, holes kept apart from written nulls; and a sparse sheet
 * whose chunks hold one cell each keeps a heap of its cells, not of their chunks' empty rows (in a fresh process,
 * `fixtures/column-heap-child.ts`).
 */
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { Column } from '../src/column.js'
import type { V } from '../src/values.js'

/** The column after `writes`, checked against a map of them over rows 0 to `rows`. */
function same(writes: [number, V][], rows: number) {
  const c = new Column(), want = new Map<number, V>()
  for (const [r, v] of writes) (c.set(r, v), want.set(r, v))
  const got: [number, V][] = []
  c.forEach((v, r) => got.push([r, v]))
  expect(got).toEqual([...want].sort((a, b) => a[0] - b[0]))
  const wrong = []
  for (let r = 0; r <= rows; r++) if (c.has(r) !== want.has(r) || c.get(r) !== want.get(r)) wrong.push(r)
  expect(wrong).toEqual([])
  expect(c.length).toBe(Math.max(0, ...want.keys()) + (want.size ? 1 : 0))
}

describe('a column of values', () => {
  it('a cell alone in the middle of its chunk, and rows written above it later, keep the rows between unwritten', () => {
    same([[500, 'a'], [10, 2], [300, null], [499, true], [0, 1]], 1100)
  })

  it('rows written upwards across chunks, and every other row of a chunk', () => {
    same(Array.from({ length: 3001 }, (_, i) => [3000 - i, 3000 - i]), 3100)
    same(Array.from({ length: 512 }, (_, i) => [2 * i + 1, i % 3 ? i : null] as [number, V]), 1100)
  })

  it('rows written in a random order, with values written again', () => {
    let seed = 7
    const next = () => ((seed = (seed * 1103515245 + 12345) % 2 ** 31), seed)
    same(Array.from({ length: 4000 }, () => [next() % 6000, next() % 5 ? next() % 100 : null] as [number, V]), 6100)
  })
})

it("a sparse sheet keeps its cells' heap, not its chunks' empty rows", () => {
  const child = fileURLToPath(new URL('./fixtures/column-heap-child.ts', import.meta.url))
  const out = JSON.parse(execFileSync(process.execPath, ['--expose-gc', '--import', 'tsx', child], { encoding: 'utf8' })) as { heap: number; sum: number }
  expect(out.sum).toBe(3)
  // chunks that held their rows from the first of 1,024 kept about 8.6 MB here; their cells and formulas take about 2
  expect(out.heap).toBeLessThan(5)
}, 60_000)
