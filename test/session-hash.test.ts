/**
 * `Session.hash()`: the values hash of `result()` without building the result, in a synchronous session, in the same
 * thread and in a thread, with closures and eval: the case books, the table books, a large spill whose unstored cells
 * hash as one line, a book in ru-RU and in the 1904 system; after edits too.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { buildTableCases } from '../bench/coverage/tables/ops.js'
import { createInlineCalc, openSession, type BookInput, type Execution } from '../src/index.js'
import { synthBook } from '../src/synth.js'
import { BOOKS, CLOCK } from './async-books.js'
import { pool } from './fixtures/pool.js'
import { edits } from './session-edits.js'

const fills: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, B1: { formula: 'A:A*2' }, C1: { formula: 'SEQUENCE(3,2)' } } }] }
const BOOKS_HERE: [string, BookInput][] = [...BOOKS, ['a whole column spilled', fills], ...buildTableCases().slice(0, 6).map((c): [string, BookInput] => [`the table case ${c.id}`, c.book])]

const pools: { close(): Promise<void> }[] = []
afterAll(() => Promise.all(pools.map((p) => p.close())))

describe.each(['closures', 'eval'] as const)('hash(), execution %s', (execution: Execution) => {
  it('a synchronous session', () => {
    for (const [name, book] of BOOKS_HERE) {
      const s = openSession(book, { clock: CLOCK, execution })
      expect(s.hash(), name).toBe(s.result().hash)
    }
  })

  it('in the same thread and in a thread', async () => {
    const calc = createInlineCalc({ sliceMs: 2 })
    const p = await pool({ threads: 2 })
    pools.push(p)
    for (const [name, book] of BOOKS_HERE) {
      const expected = openSession(book, { clock: CLOCK, execution }).result().hash
      const [a, b] = [await calc.openSession(book, { clock: CLOCK, execution }), await p.openSession(book, { clock: CLOCK, execution })]
      expect([await a.hash(), await b.hash()], name).toEqual([expected, expected])
    }
  }, 120_000)
})

describe('hash()', () => {
  it('after edits, as result() after them', async () => {
    const synth = synthBook({ seed: 1, rows: 340 })
    const s = openSession(synth.book, { clock: CLOCK })
    const p = await pool({ threads: 1 })
    pools.push(p)
    const t = await p.openSession(synth.book, { clock: CLOCK })
    for (const batch of edits(8, 20, synth)) {
      s.setCells(batch)
      await t.setCells(batch)
      expect(await t.hash()).toBe(s.hash())
      expect(s.hash()).toBe(s.result().hash)
    }
  }, 120_000)
})
