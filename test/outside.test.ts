/**
 * The answers to the calls a workbook makes outside itself (`Externals.resolve`, `src/outside.ts`): the host's
 * functions with any arguments and the functions of connections, a cube member passed by its cell, one call asked once
 * per calculation and kept by a session until `refreshConnections`, an answer by a promise waiting as #BUSY! or
 * #GETTING_DATA (`pending`), which a session takes by `settleConnections` and the calculators wait for within the call's
 * `timeoutMs`, a call's abort reaching the host. In both execution modes, the same thread and a pool.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createInlineCalc, createSession, type AbortSignalLike, type BookInput, type Calc, type CellValue, type ExternalAnswer, type ExternalCall, type Externals } from '../src/index.js'
import { pool as testPool } from './fixtures/pool.js'

const CLOCK = { now: '2026-10-07T12:00:00Z', seed: 1 }

const book: BookInput = {
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 2 },
        A2: { value: 3 },
        B1: { formula: 'BALANCE(A1:A2)*2' },
        C1: { formula: 'CUBEMEMBER("Sales","[Time].[2025]")' },
        D1: { formula: 'CUBEVALUE("Sales",C1)' },
        E1: { formula: 'CUBESET("Sales","{[A],[B]}")' },
        F1: { formula: 'CUBESETCOUNT(E1)' },
        G1: { formula: 'WEBSERVICE("http://x")' },
        H1: { formula: 'RTD("prog",,"x")+BALANCE(A1:A2)' },
        I1: { formula: 'SERIES()' },
      },
    },
  ],
}

const CELLS = ['B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1']

/** The host's answers by function. */
function answer(c: ExternalCall): ExternalAnswer {
  switch (c.name) {
    case 'BALANCE':
      return (c.args[0] as CellValue[][]).flat().reduce<number>((a, v) => a + (v as number), 0)
    case 'CUBEMEMBER':
      return { value: '2025', member: c.args[1] as string }
    case 'CUBEVALUE':
      return c.args[1] === '[Time].[2025]' ? 77 : 0
    case 'CUBESET':
      return { value: 'A and B', member: c.args[1] as string }
    case 'CUBESETCOUNT':
      return c.args[0] === '{[A],[B]}' ? 2 : 0
    case 'RTD':
      return 100
    case 'SERIES':
      return [[1, 2]]
    case 'WEBSERVICE':
      throw new Error('offline')
  }
}

const ANSWERED = [10, '2025', 77, 'A and B', 2, { error: 'VALUE' }, 105, 1]
const WAITING = [{ error: 'BUSY' }, { error: 'GETTING_DATA' }, { error: 'GETTING_DATA' }, { error: 'GETTING_DATA' }, { error: 'GETTING_DATA' }, { error: 'BUSY' }, { error: 'BUSY' }, { error: 'BUSY' }]

/** A host that answers at once, or by promises it settles when `release` is called; its calls listed. */
function host(later = false) {
  const calls: string[] = []
  const held: (() => void)[] = []
  const resolve = (c: ExternalCall) => {
    calls.push(c.key)
    if (!later) return answer(c)
    return new Promise<ExternalAnswer>((ok, no) =>
      held.push(() => {
        try {
          ok(answer(c))
        } catch (e) {
          no(e)
        }
      }),
    )
  }
  const externals: Externals = { functions: ['BALANCE', 'SERIES'], resolve }
  return { calls, externals, release: () => held.splice(0).forEach((f) => f()) }
}

/** The cells asking the host while the members are on their way: the cube functions over them wait without asking. */
const ASKING = ['B1', 'C1', 'E1', 'G1', 'H1', 'I1']

/** Every call of the book. */
const CALLS = ['BALANCE({2;3})', 'CUBEMEMBER("Sales","[Time].[2025]")', 'CUBESET("Sales","{[A],[B]}")', 'CUBESETCOUNT("{[A],[B]}")', 'CUBEVALUE("Sales","[Time].[2025]")', 'RTD("prog",,"x")', 'SERIES()', 'WEBSERVICE("http://x")']

const values = (cells: Record<string, { value: unknown }>) => CELLS.map((a) => cells[a].value)

describe.each(['closures', 'eval'] as const)('answers from outside the workbook, execution %s', (execution) => {
  it("the host's functions with any arguments, the functions of connections, a cube member passed by its cell, an array spilled; a call asked once", () => {
    const h = host()
    const r = compute(book, { clock: CLOCK, externals: h.externals, execution })
    expect(values(r.sheets[0].cells)).toEqual(ANSWERED)
    expect(r.sheets[0].cells.J1.value).toBe(2)
    expect(r.pending).toBeUndefined()
    expect([...h.calls].sort()).toEqual(CALLS)
  })

  it('values given in advance go first; answers are values of cells: hash includes them, inputHash does not', () => {
    const h = host()
    const r = compute(book, { clock: CLOCK, externals: { ...h.externals, values: { 'BALANCE({2;3})': 50 } }, execution })
    expect([r.sheets[0].cells.B1.value, r.sheets[0].cells.H1.value]).toEqual([100, 150])
    expect(h.calls).not.toContain('BALANCE({2;3})')
    const plain = compute(book, { clock: CLOCK, externals: host().externals, execution })
    expect([r.hash === plain.hash, r.inputHash === plain.inputHash]).toEqual([false, true])
  })

  it('without an answer: the functions of connections as without the host, a host function #N/A', () => {
    const r = compute(book, { clock: CLOCK, externals: { functions: ['BALANCE', 'SERIES'], resolve: () => undefined }, execution })
    expect(values(r.sheets[0].cells)).toEqual([{ error: 'NA' }, { error: 'NA' }, { error: 'NA' }, { error: 'NA' }, { error: 'NA' }, { error: 'VALUE' }, { error: 'NA' }, { error: 'NA' }])
  })

  it('answers by promises: the synchronous engine leaves the cells waiting, those asking listed (pending), those over them waiting with them', () => {
    const r = compute(book, { clock: CLOCK, externals: host(true).externals, execution })
    expect(values(r.sheets[0].cells)).toEqual(WAITING)
    expect(r.pending?.map((c) => c.cell)).toEqual(ASKING)
  })
})

describe('a session keeps the answers', () => {
  it('until refreshConnections, of the functions named or of all; settleConnections takes the answers on their way, the calls they lead to too', async () => {
    const h = host(true)
    const s = createSession(book, { clock: CLOCK, externals: h.externals })
    const told: string[] = []
    s.subscribe((c) => told.push(c.method))
    const settled = s.settleConnections()
    // the members' answers lead to calls of the cube functions over them, answered in turn
    const releasing = setInterval(h.release, 1)
    const e = await settled
    clearInterval(releasing)
    expect(values(s.getResult().sheets[0].cells)).toEqual(ANSWERED)
    expect([s.getResult().pending, told, e.changes.map((c) => c.cell).sort()]).toEqual([undefined, ['settleConnections'], CELLS.concat('J1').sort()])
    h.calls.length = 0
    s.setCells([{ sheet: 'S', cell: 'K1', input: { formula: 'BALANCE(A1:A2)' } }])
    expect([s.getCell('S', 'K1')?.value, h.calls]).toEqual([5, []])
    s.refreshConnections(['BALANCE'])
    expect(h.calls).toEqual(['BALANCE({2;3})'])
    expect(s.getCell('S', 'B1')?.value).toEqual({ error: 'BUSY' })
    h.release()
    await s.settleConnections()
    h.calls.length = 0
    s.refreshConnections()
    expect(s.getResult().pending?.map((c) => c.cell)).toEqual([...ASKING, 'K1'])
    const again = setInterval(h.release, 1)
    await s.settleConnections()
    clearInterval(again)
    expect([[...h.calls].sort(), values(s.getResult().sheets[0].cells)]).toEqual([CALLS, ANSWERED])
  })

  it('refreshConnections calculates the formulas keeping the values their file stores', () => {
    const s = createSession(book, { clock: CLOCK, externals: host().externals, cached: { sheets: [{ name: 'S', cells: { D1: { value: 5 } } }] } })
    expect(s.getCell('S', 'D1')?.value).toBe(5)
    expect(s.refreshConnections(['CUBEVALUE']).changes.map((c) => `${c.cell} ${JSON.stringify(c.value)}`)).toEqual(['D1 77'])
  })

  it("refreshConnections takes names as an array of strings", () => {
    const s = createSession(book, { clock: CLOCK, externals: host().externals })
    expect(() => s.refreshConnections('BALANCE' as never)).toThrow('names: expected an array of function names')
    expect(() => s.refreshConnections([1] as never)).toThrow('names[0]')
  })
})

const calcs: Calc[] = []
afterAll(() => Promise.all(calcs.map((c) => c.close())))

describe.each([
  ['in the same thread', async () => createInlineCalc()],
  ['in a pool', async () => testPool({ threads: 1 })],
] as const)('the calculators wait for the answers, %s', (_, make) => {
  const calc = async () => {
    const c = await make()
    calcs.push(c)
    return c
  }

  it('a computation and a session wait for answers by promises, the rounds of members too', async () => {
    const c = await calc()
    const h = host(true)
    const releasing = setInterval(h.release, 1)
    try {
      const r = await c.compute(book, { clock: CLOCK, externals: h.externals })
      expect([values(r.sheets[0].cells), r.pending]).toEqual([ANSWERED, undefined])
      const s = await c.createSession(book, { clock: CLOCK, externals: h.externals })
      expect(values((await s.getResult()).sheets[0].cells)).toEqual(ANSWERED)
      h.calls.length = 0
      const e = await s.refreshConnections(['BALANCE'])
      expect([h.calls, e.changes]).toEqual([['BALANCE({2;3})'], []])
      await s.close()
    } finally {
      clearInterval(releasing)
    }
  })

  it('an answer at once, as the synchronous engine', async () => {
    const c = await calc()
    const r = await c.compute(book, { clock: CLOCK, externals: host().externals })
    expect([values(r.sheets[0].cells), r.hash]).toEqual([ANSWERED, compute(book, { clock: CLOCK, externals: host().externals }).hash])
  })

  it('past timeoutMs the call succeeds with the cells waiting; the session takes the answers later', async () => {
    const c = await calc()
    const h = host(true)
    const t0 = Date.now()
    const r = await c.compute(book, { clock: CLOCK, externals: h.externals, timeoutMs: 100 })
    expect([values(r.sheets[0].cells), r.pending?.map((x) => x.cell)]).toEqual([WAITING, ASKING])
    expect(Date.now() - t0).toBeGreaterThanOrEqual(90)
    const s = await c.createSession(book, { clock: CLOCK, externals: h.externals, timeoutMs: 50 })
    expect((await s.getCell('S', 'B1'))?.value).toEqual({ error: 'BUSY' })
    const releasing = setInterval(h.release, 1)
    try {
      await s.settleConnections()
      expect([values((await s.getResult()).sheets[0].cells), (await s.getResult()).pending]).toEqual([ANSWERED, undefined])
    } finally {
      clearInterval(releasing)
    }
    await s.close()
  })

  it("a call's abort ends the wait and reaches the host's calls", async () => {
    const c = await calc()
    const signals: AbortSignalLike[] = []
    const ac = new AbortController()
    const p = c.compute(book, {
      clock: CLOCK,
      signal: ac.signal,
      externals: {
        functions: ['BALANCE', 'SERIES'],
        resolve: (x) => {
          // aborted once the calculation is under way
          if (!signals.length) setTimeout(() => ac.abort(), 20)
          signals.push(x.signal)
          return new Promise(() => {})
        },
      },
    })
    await expect(p).rejects.toThrow()
    await new Promise((r) => setTimeout(r, 20))
    expect(signals.length).toBeGreaterThan(0)
    expect(signals.every((s) => s.aborted)).toBe(true)
  })
})
