/**
 * A session lets go of the state it built before: an operation that builds its book again (a table's header renamed, a
 * row inserted) keeps the heap where it was, where each one once kept the state before it, about 9 MB a time over a
 * table of 3,400 rows (the formulas a circular reference stops looked up what they showed through the old state, which
 * the new one held on to). In a fresh process with `--expose-gc` (`fixtures/rebuild-heap-child.ts`).
 */
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { expect, it } from 'vitest'

it('eight operations that build the book again over a table of 3,400 rows keep the heap from growing at each', () => {
  const child = fileURLToPath(new URL('./fixtures/rebuild-heap-child.ts', import.meta.url))
  const out = JSON.parse(execFileSync(process.execPath, ['--expose-gc', '--import', 'tsx', child, '3400'], { encoding: 'utf8' })) as { heap: number[]; sum: string }
  expect(out.sum).toBe('SUM(T[Double])')
  // from the first operation to the last: a state kept adds about 9 MB at each (63 MB in all), a cache of the process's
  // ranges, bounded, may fill once on the way (9 MB once under the full run)
  expect(out.heap[8] - out.heap[1], JSON.stringify(out.heap.map(Math.round))).toBeLessThan(20)
}, 120_000)
