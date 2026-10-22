/**
 * Reproductions from the review of 21.10 on the pool of threads: a session's options are its own from the call on, so a
 * restore does not pick up what the caller changes in them later; the journal stays within `snapshotEvery` calls when
 * calls are sent without waiting for each other.
 */
import { describe, expect, it } from 'vitest'
import { createSession, type BookInput } from '../src/index.js'
import { closedAfter, inside } from './fixtures/pool.js'

const { keep } = closedAfter()
const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const edit = (v: number) => [{ sheet: 'S', cell: 'A2', input: { value: v } }]

describe.each(['closures', 'eval'] as const)('the pool, %s', (execution) => {
  it("a restore opens the options the session was given, not what the caller changed in them since (fix-u368)", async () => {
    const book = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { formula: 'BALANCE(1)+A2' }, A2: { value: 1 } } }] })
    const { calc } = await inside({ threads: 1, maxSessionsPerThread: 1 })
    keep(calc)
    const externals = { functions: ['BALANCE'], values: { 'BALANCE(1)': 5 } }
    const a = await calc.createSession(book(), { clock, execution, externals })
    externals.values['BALANCE(1)'] = 777
    await a.setCells(edit(2))
    const stayed = [(await a.getCell('S', 'A1'))?.value, await a.getHash()]
    expect(stayed[0]).toBe(7)
    // another session takes the full thread: a leaves it and comes back from its snapshot
    await calc.createSession(book(), { clock, execution })
    expect([(await a.getCell('S', 'A1'))?.value, await a.getHash()]).toEqual(stayed)
  })

  it('options a thread cannot take are refused before the session opens', async () => {
    const { calc } = await inside({ threads: 1 })
    keep(calc)
    await expect(calc.createSession(book2(), { clock, execution, externals: { values: { f: () => 1 } } as never })).rejects.toThrow(/^opts: cannot be sent to a thread/)
  })

  it('pipelined changing calls keep the journal within snapshotEvery, and the session its hash (fix-u371)', async () => {
    const { calc, inside: x } = await inside({ threads: 1, snapshotEvery: 4 })
    keep(calc)
    const a = await calc.createSession(book2(), { clock, execution })
    const r = [...x.recs][0]
    let longest = 0
    await Promise.all(Array.from({ length: 200 }, (_, i) => a.setCells(edit(i)).then(() => (longest = Math.max(longest, r.journal.length)))))
    // a snapshot at the end of the queue waited for the whole burst: the journal grew to 200
    expect(longest).toBeLessThanOrEqual(4)
    const sync = createSession(book2(), { clock, execution })
    for (let i = 0; i < 200; i++) sync.setCells(edit(i))
    expect(await a.getHash()).toBe(sync.getHash())
  })
})

const book2 = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { formula: 'A2*2' }, A2: { value: 1 } } }] })
