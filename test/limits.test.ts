/**
 * Limits (`ComputeOptions.limits`, `SessionOptions.limits`): a book over them throws `CalcError` with code `limit`,
 * naming the limit and, for `cells`, the formula's cell, before the values are allocated, so an array that would
 * bring the process down stops the call instead; a session call over them leaves the session as it was. Without
 * limits nothing changes.
 */
import { describe, expect, it } from 'vitest'
import { CalcError, compute, openSession, type BookInput } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const book = (cells: BookInput['sheets'][0]['cells']): BookInput => ({ sheets: [{ name: 'S', cells }] })

function stopped(f: () => unknown): CalcError {
  try {
    f()
  } catch (e) {
    expect(e).toBeInstanceOf(CalcError)
    return e as CalcError
  }
  throw new Error('expected CalcError')
}

describe('limits', () => {
  it('formulas: a book with more formulas than the limit is not computed', () => {
    const b = book({ A1: { value: 1 }, B1: { formula: 'A1+1' }, B2: { formula: 'A1+2' }, B3: { formula: 'SUM(' } })
    const e = stopped(() => compute(b, { clock, limits: { formulas: 2 } }))
    expect([e.code, e.limit, e.max, e.size]).toEqual(['limit', 'formulas', 2, 3])
    expect(compute(b, { clock, limits: { formulas: 3 } }).hash).toBe(compute(b, { clock }).hash)
  })

  it('cells: arrays too large stop the call before they are made, with the formula cell', () => {
    for (const formula of ['SUM(SEQUENCE(20000,20000))', 'SUM(MUNIT(30000))', 'SUM(ROW(A1:A2000)*COLUMN(A1:ZZ1))', 'ROWS(TOCOL(C:E))', 'MMULT(SEQUENCE(3000,1),SEQUENCE(1,3000))']) {
      const e = stopped(() => compute(book({ A1: { value: 1 }, B7: { formula } }), { clock, limits: { cells: 1_000_000 } }))
      expect([e.code, e.limit, e.max], formula).toEqual(['limit', 'cells', 1_000_000])
      expect(e.size, formula).toBeGreaterThan(1_000_000)
      expect(e.cell, formula).toEqual({ sheet: 'S', cell: 'B7' })
      expect(e.message, formula).toMatch(/over the limit of 1000000 \(S!B7\)$/)
    }
  })

  it('cells: what fits computes as without limits, whole columns by their occupied rows', () => {
    const cells: BookInput['sheets'][0]['cells'] = {}
    for (let r = 1; r <= 500; r++) cells[`A${r}`] = { value: r }
    Object.assign(cells, { C1: { formula: 'SUM(A:A*2)' }, C2: { formula: 'SUMIFS(A:A,A:A,">10")' }, C3: { formula: 'SEQUENCE(10,10)' }, C20: { formula: 'B:B' } })
    const b = book(cells)
    expect(compute(b, { clock, limits: { cells: 2_000 } }).hash).toBe(compute(b, { clock }).hash)
  })

  it('a spill over the limit stops the call', () => {
    const e = stopped(() => compute(book({ A1: { formula: 'SEQUENCE(3000)' } }), { clock, limits: { cells: 2_000 } }))
    expect([e.limit, e.cell]).toEqual(['cells', { sheet: 'S', cell: 'A1' }])
  })

  it('a session call over the limits changes nothing, and the session goes on', () => {
    const b = book({ A1: { value: 2 }, A2: { formula: 'A1*10' }, A3: { formula: 'SUM(A1:A2)' } })
    const s = openSession(b, { clock, limits: { formulas: 3, cells: 10_000 } })
    const hash = s.result().hash
    const big = stopped(() => s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }, { sheet: 'S', cell: 'B1', input: { formula: 'SUM(SEQUENCE(1000,1000))' } }]))
    expect(big.limit).toBe('cells')
    expect(s.book().sheets[0].cells).toEqual(b.sheets[0].cells)
    expect(s.result().hash).toBe(hash)
    const many = stopped(() => s.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'A1+1' } }, { sheet: 'S', cell: 'B2', input: { formula: 'A1+2' } }]))
    expect(many.limit).toBe('formulas')
    expect(s.result().hash).toBe(hash)
    expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }]).map((c) => [c.cell, c.value])).toEqual([
      ['A1', 3],
      ['A2', 30],
      ['A3', 33],
    ])
    expect(s.result().hash).toBe(compute(book({ A1: { value: 3 }, A2: { formula: 'A1*10' }, A3: { formula: 'SUM(A1:A2)' } }), { clock }).hash)
    expect(() => openSession(b, { clock, limits: { formulas: 1 } })).toThrow(CalcError)
  })

  it('limits are checked as input', () => {
    expect(() => compute(book({}), { limits: { cells: 0 } })).toThrow('opts.limits.cells: expected an integer from 1')
    expect(() => compute(book({}), { limits: { rows: 1 } as never })).toThrow('opts.limits: unknown field "rows"')
  })
})
