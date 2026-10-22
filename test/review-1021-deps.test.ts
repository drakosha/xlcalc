/** Reproductions from the review of 21.10 on dependencies: a session recomputes every formula whose read cells an edit changes. */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createInlineCalc, createSession, type BookInput, type Calc, type CellInput } from '../src/index.js'
import { pool } from './fixtures/pool.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** SUMIF and AVERAGEIF read their sum range resized to the criteria range: A1, then an edit of a cell the resized range takes in. */
const CASES: [string, string, CellInput][] = [
  ['SUMIF(C1:C3,"",B1)', 'B2', { value: 5 }],
  ['AVERAGEIF(C1:C3,"",B1)', 'B2', { value: 5 }],
  ['SUMIF(C1:C3,"",B1:B2)', 'B3', { value: 5 }],
  ['SUMIF(C1:C3,"",B1)', 'B2', { formula: 'SEQUENCE(2)' }],
  ['SUMIF(C1:C3,"",B1)+B1', 'B2', { value: 5 }],
  ['SUMIF(C1:C3,"",(B1))', 'B3', { value: 5 }],
  ['LET(x,B1,SUMIF(C1:C3,"",x))', 'B3', { value: 5 }],
  ['SUMIF(C1:C3,"",one)', 'B3', { value: 5 }],
  ['SUMIF(C1:C3,"",INDEX(B1:B1,1))', 'B2', { value: 5 }],
  ['SUMIF(C1:C3,"",E1#)', 'E3', { value: 5 }],
  ['SUMIF(INDEX(C1:C3,0,1),"",B1)', 'B3', { value: 5 }],
  ['SUMIF(C:C,"",B1)', 'B500', { value: 5 }],
  ['SUMIF(C1:D2,"",B1)', 'C2', { value: 5 }],
]
const book = (formula: string): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { formula }, E1: { formula: '1' } } }], names: [{ name: 'one', ref: 'S!$B$1' }] })

let threads: Calc | undefined
afterAll(() => threads?.close())

describe.each(['closures', 'eval'] as const)('execution %s', (execution) => {
  it.each(CASES)('%s follows an edit of %s (x-fuzz-modes-5)', async (formula, cell, input) => {
    const edit = [{ sheet: 'S', cell, input }]
    const s = createSession(book(formula), { clock, execution })
    s.setCells(edit)
    const full = compute(s.getBook(), { clock, execution })
    expect(s.getCell('S', 'A1')?.value).toEqual(full.sheets[0].cells.A1.value)
    expect(s.getHash()).toBe(full.hash)
    const calc = createInlineCalc({ sliceMs: 1 })
    threads ??= await pool({ threads: 1 })
    for (const c of [calc, threads]) {
      const t = await c.createSession(book(formula), { clock, execution })
      await t.setCells(edit)
      expect(await t.getHash()).toBe(full.hash)
      await t.close()
    }
  })
})
