/**
 * Reproductions from the review of 21.10 on the pool of threads: a book whose text cannot be written, a restore that
 * cannot be sent, a thread that cannot start, and the deadline and signal of reads in a thread.
 */
import { describe, expect, it } from 'vitest'
import { type BookInput, type CellInput } from '../src/index.js'
import { chainBook } from './async-books.js'
import { closedAfter } from './fixtures/pool.js'

const { made } = closedAfter()

/** What `p` settles to, or 'hung' when it has not within `ms`. */
const settled = (p: Promise<unknown>, ms = 3000): Promise<unknown> =>
  Promise.race([p.then((v) => ['resolved', v], (e) => ['rejected', e]), new Promise((resolve) => setTimeout(() => resolve('hung'), ms))])

/** A book of `n` cells whose first cell's meta throws when it is read, as a cyclic one overflows the stack. */
function throwingBook(n: number): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= n; r++) cells[`A${r}`] = { value: r }
  cells.A1 = {
    value: 1,
    meta: {
      get x(): never {
        throw new Error('getter boom')
      },
    },
  } as unknown as CellInput
  return { sheets: [{ name: 'S', cells }] }
}

/** Over the 16,384 cells written at once: the book's text is written between turns of the event loop. */
const LARGE = 16_385

describe.each(['closures', 'eval'] as const)('the pool, %s', (execution) => {
  it('a book whose text cannot be written fails its call, and the tasks behind it run', async () => {
    const p = await made({ threads: 1 })
    expect(await settled(p.compute(throwingBook(10), { execution }))).toEqual(['rejected', expect.objectContaining({ message: 'getter boom' })])
    const bad = p.compute(throwingBook(LARGE), { execution })
    const next = p.compute(chainBook(2), { execution })
    expect(await settled(bad)).toEqual(['rejected', expect.objectContaining({ message: 'getter boom' })])
    expect(await settled(next)).toEqual(['resolved', expect.anything()])
    const opening = p.createSession(throwingBook(LARGE), { execution })
    const after = p.createSession(chainBook(2), { execution })
    const job = p.compute(chainBook(2), { execution })
    expect(await settled(opening)).toEqual(['rejected', expect.objectContaining({ message: 'getter boom' })])
    expect(await settled(after)).toEqual(['resolved', expect.anything()])
    expect(await settled(job)).toEqual(['resolved', expect.anything()])
  }, 30_000)

})
