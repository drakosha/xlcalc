/**
 * Iterative calculation beyond the Excel cases (`test/excel/workbook/iterate.json`): a session gives what a full
 * computation gives after every edit, both execution modes agree, the settings are checked and hashed.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, openSession, type BookInput, type CellReplace, type CellInput } from '../src/index.js'
import { mulberry32 } from '../src/synth.js'

const clock = { now: '2026-10-07T12:00:00Z', seed: 1 }

/** S: interest on the average balance (B3, B4), a counter (D1), a converging cycle over two sheets (S!E1, T!A1), dependents. */
function book(iteration: BookInput['iteration'] | null = {}): BookInput {
  const cells: Record<string, CellInput> = {
    B1: { value: 1000 },
    B2: { value: 0.1 },
    B3: { formula: 'B2*(B1+B4)/2' },
    B4: { formula: 'B1+B3-100' },
    C1: { formula: 'B4-B1' },
    D1: { formula: 'D1+B2' },
    E1: { formula: "'T'!A1/2+B2" },
    F1: { formula: 'SUM(B3:E1)' },
  }
  return { sheets: [{ name: 'S', cells }, { name: 'T', cells: { A1: { formula: "'S'!E1" } } }], ...(iteration ? { iteration } : {}) }
}

const values = (b: BookInput) => Object.fromEntries(Object.entries(compute(b, { clock, display: false }).sheets[0].cells).map(([k, c]) => [k, c.value]))

describe('iterative calculation', () => {
  it('computes cycles by iterations, and without the setting they are #CYCLE!', () => {
    const v = values(book({ maxIterations: 10 }))
    expect(v.D1).toBeCloseTo(1, 12)
    expect(v.B3 as number).toBeCloseTo(100, 6)
    expect(values(book(null)).B3).toEqual({ error: 'CYCLE' })
  })

  it('stops when no value changes by the maximum change', () => {
    const coarse = values(book({ maxChange: 1 }))
    const fine = values(book({ maxChange: 1e-12 }))
    expect(Math.abs((coarse.B3 as number) - 100)).toBeGreaterThan(1e-6)
    expect(fine.B3).toBe(100)
    // the counter runs every round: 100 by default, stopped early by the converging cycles in the coarse book
    expect(fine.D1).toBeCloseTo(10, 9)
    expect(coarse.D1 as number).toBeLessThan(1)
  })

  it('a session gives the values and the hash of a full computation after every edit, in both execution modes', () => {
    const rnd = mulberry32(3)
    const inputs: (CellInput | null)[] = [{ value: 500 }, { value: 0.05 }, { formula: 'B1*2' }, { formula: 'G1+1' }, null]
    const at = ['B1', 'B2', 'G1', 'D1', 'B4']
    for (const execution of ['closures', 'eval'] as const) {
      let b = book()
      const s = openSession(b, { clock, execution })
      for (let i = 0; i < 150; i++) {
        const e: CellReplace = { sheet: 'S', cell: at[Math.floor(rnd() * at.length)], input: inputs[Math.floor(rnd() * inputs.length)] }
        s.setCells([e])
        const cells = { ...b.sheets[0].cells }
        if (e.input === null) delete cells[e.cell]
        else cells[e.cell] = e.input
        b = { ...b, sheets: [{ ...b.sheets[0], cells }, b.sheets[1]] }
        const full = compute(b, { clock, execution: 'closures' })
        expect(s.result().sheets, `${execution} edit ${i}: ${JSON.stringify(e)}`).toEqual(full.sheets)
        expect(s.result().hash).toBe(full.hash)
      }
      s.close()
    }
  })

  it('the settings are checked and go into the input hash', () => {
    for (const bad of [{ maxIterations: 0 }, { maxIterations: 32_768 }, { maxIterations: 1.5 }, { maxChange: -1 }, { steps: 3 }])
      expect(() => compute(book(bad as BookInput['iteration']))).toThrow(InputError)
    const [a, b, c] = [compute(book(null), { clock }), compute(book(), { clock }), compute(book({ maxIterations: 7 }), { clock })]
    expect(new Set([a.inputHash, b.inputHash, c.inputHash]).size).toBe(3)
  })
})
