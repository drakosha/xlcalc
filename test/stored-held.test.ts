/**
 * Formulas that keep the value their file stores (`held.ts`): over a function the engine does not know under
 * `unknownFunctions: 'held'`, and over a function of connections whatever the option. In every kind of session and both
 * execution modes: the cells that read them are computed from the value, diagnostics mark them `fromFile`, an edit that
 * reaches them computes them (`#NAME?`, the connection's answer), an unrelated one does not; without a stored value a
 * formula over a function the engine does not know is #NAME?, as without the option.
 */
import { describe, expect, it } from 'vitest'
import { compute, createInlineCalc, createSession, type BookInput, type CachedValues } from '../src/index.js'
import { pool as testPool } from './fixtures/pool.js'

const CLOCK = { now: '2026-10-07T12:00:00Z', seed: 1 }

const book: BookInput = {
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 2 },
        B1: { formula: 'MYADDIN(A1)*2' },
        C1: { formula: 'B1+1' },
        D1: { formula: 'CUBEVALUE("Sales","[Measures].[Amount]")' },
        E1: { formula: 'D1*10' },
        F1: { formula: 'OTHER()' },
        G1: { formula: 'IFERROR(MYADDIN(A1),"none")' },
      },
    },
  ],
}

/** What Excel stored: the add-in's 40, the cube's 7, values the engine computes again for C1 and E1, nothing for F1 and G1. */
const cached: CachedValues = { sheets: [{ name: 'S', cells: { B1: { value: 40 }, C1: { value: 99 }, D1: { value: 7 }, E1: { value: 0 } } }] }

const values = (cells: Record<string, { value: unknown }>) => ['B1', 'C1', 'D1', 'E1', 'F1', 'G1'].map((a) => cells[a].value)
const marks = (d: { kind: string; cell?: string; error?: string; fromFile?: true; check?: string }[]) => d.map((x) => `${x.cell} ${x.error}${x.fromFile ? ' file' : ''}${x.check ? ` ${x.check}` : ''}`)

describe.each(['closures', 'eval'] as const)('formulas keeping their file values, execution %s', (execution) => {
  it('held: the file values, the cells reading them computed from them; the rest as without the option', () => {
    const r = compute(book, { clock: CLOCK, cached, unknownFunctions: 'held', execution })
    expect(values(r.sheets[0].cells)).toEqual([40, 41, 7, 70, { error: 'NAME' }, { error: 'NAME' }])
    expect(marks(r.diagnostics)).toEqual(['B1 NAME file unknown_function', 'D1 CONNECT file', 'F1 NAME unknown_function', 'G1 NAME unknown_function'])
  })

  it('without the option a function the engine does not know is #NAME?; a function of connections keeps its value all the same', () => {
    const r = compute(book, { clock: CLOCK, cached, execution })
    expect(values(r.sheets[0].cells)).toEqual([{ error: 'NAME' }, { error: 'NAME' }, 7, 70, { error: 'NAME' }, { error: 'NAME' }])
    expect(compute(book, { clock: CLOCK, unknownFunctions: 'held', execution }).sheets[0].cells.D1.value).toEqual({ error: 'NA' })
  })

  it('a session: an edit away keeps the values, one that reaches them computes them; getBook holds them as held', () => {
    const s = createSession(book, { clock: CLOCK, cached, unknownFunctions: 'held', execution })
    expect(s.getBook().sheets[0].cells.B1).toEqual({ formula: 'MYADDIN(A1)*2', held: 40 })
    s.setCells([{ sheet: 'S', cell: 'K9', input: { value: 1 } }])
    expect(values(s.getResult().sheets[0].cells)).toEqual([40, 41, 7, 70, { error: 'NAME' }, { error: 'NAME' }])
    const { changes } = s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
    expect(changes.map((c) => `${c.cell} ${JSON.stringify(c.value)}`)).toEqual(['A1 3', 'B1 {"error":"NAME"}', 'C1 {"error":"NAME"}'])
    expect(s.getBook().sheets[0].cells.B1).toEqual({ formula: 'MYADDIN(A1)*2' })
    expect(marks(s.getResult().diagnostics).slice(0, 2)).toEqual(['B1 NAME unknown_function', 'C1 NAME'])
    // a formula entered again computes: the cube without its connection
    s.setCells([{ sheet: 'S', cell: 'D1', input: { formula: 'CUBEVALUE("Sales","[Measures].[Amount]")' } }])
    expect([s.getCell('S', 'D1')?.value, s.getCell('S', 'E1')?.value]).toEqual([{ error: 'NA' }, { error: 'NA' }])
    // the book a session gives computes as the session shows
    const again = compute(s.getBook(), { clock: CLOCK, unknownFunctions: 'held', execution })
    expect(values(again.sheets[0].cells)).toEqual(values(s.getResult().sheets[0].cells))
  })

  it('the calculators: in the same thread and in a pool, as the synchronous engine', async () => {
    const sync = compute(book, { clock: CLOCK, cached, unknownFunctions: 'held', execution })
    const calc = createInlineCalc()
    const pool = await testPool({ threads: 1 })
    try {
      for (const c of [calc, pool]) {
        const r = await c.compute(book, { clock: CLOCK, cached, unknownFunctions: 'held', execution })
        expect([values(r.sheets[0].cells), r.hash, marks(r.diagnostics)]).toEqual([values(sync.sheets[0].cells), sync.hash, marks(sync.diagnostics)])
        const s = await c.createSession(book, { clock: CLOCK, cached, unknownFunctions: 'held', execution })
        expect((await s.getCell('S', 'C1'))?.value).toBe(41)
        await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
        expect((await s.getCell('S', 'C1'))?.value).toEqual({ error: 'NAME' })
        expect((await s.getCell('S', 'E1'))?.value).toBe(70)
      }
    } finally {
      await Promise.all([calc.close(), pool.close()])
    }
  })
})

describe('formulas keeping their file values beside dynamic references (OFFSET)', () => {
  const offsets: BookInput = {
    sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 5 }, K1: { value: 0 }, B1: { formula: 'OFFSET(A1,0,0)*5', held: 20 }, D1: { formula: 'CUBEVALUE("Sales",OFFSET(A1,K1,0))' }, E1: { formula: 'D1*10' } } }],
  }
  const file: CachedValues = { sheets: [{ name: 'S', cells: { B1: { value: 20 }, D1: { value: 7 } } }] }

  it('a formula over OFFSET, a function the engine knows, is computed, whatever value it held while OFFSET was unknown', () => {
    const r = compute(offsets, { clock: CLOCK, cached: file, unknownFunctions: 'held' })
    expect([r.sheets[0].cells.B1.value, r.sheets[0].cells.D1.value, r.sheets[0].cells.E1.value]).toEqual([5, 7, 70])
    expect(marks(r.diagnostics)).toEqual(['D1 CONNECT file'])
  })

  it("a cube's formula keeps its value until a cell its OFFSET reads changes, also one OFFSET comes to read", () => {
    const s = createSession(offsets, { clock: CLOCK, cached: file, unknownFunctions: 'held' })
    s.setCells([{ sheet: 'S', cell: 'A2', input: { value: 6 } }])
    expect(s.getCell('S', 'D1')?.value).toBe(7)
    s.setCells([{ sheet: 'S', cell: 'K1', input: { value: 1 } }])
    expect([s.getCell('S', 'D1')?.value, s.getCell('S', 'E1')?.value]).toEqual([{ error: 'NA' }, { error: 'NA' }])
  })
})

describe('formulas keeping their file values beside names the book defines as LAMBDA', () => {
  it("a name defined as a LAMBDA is called, not taken for a function the engine does not know; a formula with both keeps the file's value", () => {
    const named: BookInput = {
      sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { formula: 'DOUBLE(A1)' }, C1: { formula: 'DOUBLE(A1)+MYADDIN(A1)' }, D1: { formula: 'C1+1' } } }],
      names: [{ name: 'DOUBLE', ref: 'LAMBDA(x,x*2)' }],
    }
    const file: CachedValues = { sheets: [{ name: 'S', cells: { B1: { value: 99 }, C1: { value: 10 } } }] }
    const r = compute(named, { clock: CLOCK, cached: file, unknownFunctions: 'held' })
    expect(['B1', 'C1', 'D1'].map((a) => r.sheets[0].cells[a].value)).toEqual([4, 10, 11])
    expect(marks(r.diagnostics)).toEqual(['C1 NAME file unknown_function'])
  })
})

describe('unknownFunctions: the option checked', () => {
  it('takes name or held only', () => {
    expect(() => compute(book, { unknownFunctions: 'keep' as 'held' })).toThrow('opts.unknownFunctions: expected "name" or "held", got the string "keep"')
    expect(() => createSession(book, { unknownFunctions: 1 as unknown as 'held' })).toThrow('opts.unknownFunctions: expected "name" or "held", got number 1')
  })
})
