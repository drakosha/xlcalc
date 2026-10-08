/**
 * A book crossing to a thread as text (`src/worker/pack.ts`): it reads back as given, large or small; JSON cannot hold
 * some books (a non-finite number, -0, a date or a function in `meta`, undefined inside `meta`), which go as objects
 * and compute as before; the book is taken at once, so the caller's changes afterwards do not reach the session; the
 * pool takes its tasks in order while a large book is written; opening a large book stops the event loop for a
 * fraction of what a structured clone of it takes.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, openSession, type BookInput } from '../src/index.js'
import { synthBook } from '../src/synth.js'
import { pack, unpack, unpackSlowly, type Sent } from '../src/worker/pack.js'
import { BOOKS, chainBook, CLOCK } from './async-books.js'
import { pool } from './fixtures/pool.js'

const pools: { close(): Promise<void> }[] = []
afterAll(() => Promise.all(pools.map((p) => p.close())))

const isText = (s: Sent) => (s as { packed?: number }).packed === 1

const withMeta: BookInput = {
  styles: [{ font: { bold: true } }],
  names: [{ name: 'rate', ref: '0.2' }],
  sheets: [{ name: 'S', cells: { A1: { value: 1, meta: { ids: [1, 2], note: 'x', nested: { ok: true } } }, A2: { formula: 'A1*rate', style: 0, format: undefined }, B1: { value: { date: '2025-01-31' } } }, rows: { '2': { hidden: true } }, columns: { B: { width: 12 } } }],
}

describe('a book as text', () => {
  it('reads back as given, at once in a thread and a chunk at a time on the calling side', async () => {
    for (const [name, book] of [...BOOKS, ['a book with meta', withMeta] as [string, BookInput], ['the synthetic book of 3,400 rows', synthBook({ seed: 1, rows: 3400 }).book] as [string, BookInput]]) {
      const sent = await pack(book)
      expect(isText(sent), name).toBe(true)
      expect(unpack(sent), name).toEqual(book)
      expect(await unpackSlowly(sent), name).toEqual(book)
    }
  }, 60_000)

  it('a book JSON cannot hold goes as objects', async () => {
    const odd: [string, BookInput][] = [
      ['NaN', { sheets: [{ name: 'S', cells: { A1: { value: Number.NaN } } }] }],
      ['-0', { sheets: [{ name: 'S', cells: { A1: { value: -0 } } }] }],
      ['a date in meta', { sheets: [{ name: 'S', cells: { A1: { value: 1, meta: new Date(0) as never } } }] }],
      ['undefined inside meta', { sheets: [{ name: 'S', cells: { A1: { value: 1, meta: { a: undefined } as never } } }] }],
      ['a function in meta', { sheets: [{ name: 'S', cells: { A1: { value: 1, meta: (() => 1) as never } } }] }],
    ]
    for (const [name, book] of odd) {
      const sent = await pack(book)
      expect(isText(sent), name).toBe(false)
      expect(sent, name).toEqual(book)
    }
    // and they compute in a thread as here
    const p = await pool({ threads: 1 })
    pools.push(p)
    for (const [name, book] of odd.slice(0, 4)) expect((await p.compute(book, { clock: CLOCK })).hash, name).toBe(compute(book, { clock: CLOCK }).hash)
  })

  it('is taken at once: what the caller changes afterwards does not reach the session', async () => {
    const book = synthBook({ seed: 1, rows: 2000 }).book
    const before = structuredClone(book)
    const p = await pool({ threads: 1 })
    pools.push(p)
    const opening = p.openSession(book, { clock: CLOCK })
    // the large book is written between turns: change it right away
    const data = book.sheets[0]
    const [first, second] = Object.keys(data.cells)
    delete data.cells[first]
    data.cells[second] = { value: 'changed' }
    data.cells.ZZ9 = { value: 1 }
    book.sheets[1] = { name: 'Other', cells: {} }
    const s = await opening
    expect(await s.book()).toEqual(before)
    expect(await s.hash()).toBe(compute(before, { clock: CLOCK }).hash)
  }, 60_000)

  it('tasks keep their order while a large book is written', async () => {
    const p = await pool({ threads: 1 })
    pools.push(p)
    const order: string[] = []
    const big = p.compute(synthBook({ seed: 1, rows: 3400 }).book).then(() => order.push('large'))
    const small = p.compute(chainBook(3)).then(() => order.push('small'))
    await Promise.all([big, small])
    expect(order).toEqual(['large', 'small'])
  }, 60_000)

  it('opening a large book stops the event loop for a fraction of a structured clone of it', async () => {
    const book = synthBook({ seed: 1, rows: 8000 }).book
    const t0 = performance.now()
    structuredClone(book)
    const clone = performance.now() - t0
    const p = await pool({ threads: 1 })
    pools.push(p)
    await p.compute(chainBook(2))
    let last = performance.now()
    let worst = 0
    let done = false
    const tick = () => {
      const now = performance.now()
      worst = Math.max(worst, now - last)
      last = now
      if (!done) setImmediate(tick)
    }
    setImmediate(tick)
    const s = await p.openSession(book, { clock: CLOCK })
    done = true
    // the pool copied the book and cloned it to the thread in one stretch: more than half a clone
    expect(worst).toBeLessThan(clone / 3)
    expect(await s.hash()).toBe(openSession(book, { clock: CLOCK }).hash())
  }, 60_000)
})
