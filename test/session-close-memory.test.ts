/**
 * A closed session gives its heap back: after each close the heap returns to within 3 MB of where it was before the
 * session opened (a live session of the chain holds about 8 MB) and the closed session's object is collected, over ten
 * sessions opened and closed one after another, after one that was read, and over ten of a calculator that stays open, in
 * the same thread (it once kept every session it opened until it was closed itself) and in a pool's thread of 64 MB, with
 * a listener and a waiting call each. The JavaScript engine's background compilation of a book's formulas may hold the
 * closed state for some milliseconds, so the child waits for the heap, up to 500 ms a close, rather than reading it at
 * once. In a fresh process with `--expose-gc` (`fixtures/close-heap-child.ts`).
 */
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { expect, it } from 'vitest'

const child = fileURLToPath(new URL('./fixtures/close-heap-child.ts', import.meta.url))

it.each([
  ['ten sessions opened and closed one after another', 'series'],
  ['a session opened, read and closed', 'read'],
  ['ten sessions of an open calculator in the same thread, each closed', 'inline'],
  ['ten sessions of an open pool, each closed', 'pool'],
])('%s give the heap back after each close', (_, mode) => {
  const closes = JSON.parse(execFileSync(process.execPath, ['--expose-gc', '--import', 'tsx', child, mode], { encoding: 'utf8' })) as { ms: number; above: number; held: boolean }[]
  expect(closes.length).toBe(mode === 'read' ? 1 : 10)
  expect(closes.filter((c) => c.above > 3 || c.held), JSON.stringify(closes)).toEqual([])
}, 120_000)
