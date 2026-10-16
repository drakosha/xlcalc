/**
 * A session lets go of the state it built before: an operation that builds its book again (a table's header renamed, a
 * row inserted) keeps the heap where it was, where each one once kept the state before it, about 9 MB a time over a
 * table of 3,400 rows (the formulas a circular reference stops looked up what they showed through the old state, which
 * the new one held on to). In a fresh process with `--expose-gc` (`fixtures/rebuild-heap-child.ts`).
 */
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { expect, it } from 'vitest'

it('eight operations that build the book again over a table of 3,400 rows keep the heap within 5 MB of the session opened', () => {
  const child = fileURLToPath(new URL('./fixtures/rebuild-heap-child.ts', import.meta.url))
  const out = JSON.parse(execFileSync(process.execPath, ['--expose-gc', '--import', 'tsx', child, '3400'], { encoding: 'utf8' })) as { heap: number[]; sum: string }
  expect(out.sum).toBe('SUM(T[Double])')
  expect(Math.max(...out.heap) - out.heap[0], JSON.stringify(out.heap.map(Math.round))).toBeLessThan(5)
}, 120_000)
