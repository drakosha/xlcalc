/**
 * Limits (`ComputeOptions.limits`, `SessionOptions.limits`): a book over them throws `CalcError` with code `limit`,
 * naming the limit and, for `cells` and `spill`, the formula's cell, before the values are allocated, so an array that
 * would bring the process down stops the call instead; a session call over them leaves the session as it was. Without
 * limits nothing changes.
 */
import { describe, expect, it } from 'vitest'
import { CalcError, compute, createInlineCalc, openSession, type BookInput } from '../src/index.js'

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

  it('spill: a spill of more stored values stops the call with the cell, an array that does not spill is not bounded by it', () => {
    const limits = { spill: 2_000 }
    for (const input of [{ formula: 'SEQUENCE(1000,3)' }, { formula: 'SEQUENCE(3000)', array: 'A1:A3000' }, { formula: '1', array: 'A1:C1000' }]) {
      const e = stopped(() => compute(book({ A1: input }), { clock, limits }))
      expect([e.code, e.limit, e.max, e.size, e.cell], input.formula).toEqual(['limit', 'spill', 2_000, 3_000, { sheet: 'S', cell: 'A1' }])
      expect(e.message).toMatch(/^3000 values in one spill, over the limit of 2000 \(S!A1\)$/)
    }
    const fits = book({ A1: { formula: 'SUM(SEQUENCE(3000))' }, B1: { formula: 'SEQUENCE(2000)' } })
    expect(compute(fits, { clock, limits }).hash).toBe(compute(fits, { clock }).hash)
  })

  it('spill: a large array counts its stored cells, not the cells past its last distinct value', () => {
    const cells: BookInput['sheets'][0]['cells'] = {}
    for (let r = 1; r <= 500; r++) cells[`A${r}`] = { value: r }
    Object.assign(cells, { C1: { formula: 'A:A*2' }, D1: { formula: 'TOCOL(A:A)' }, E1: { formula: 'IF(A:A>250,A:A,0)' } })
    const b = book(cells)
    const r = compute(b, { clock, limits: { spill: 1_000 } })
    expect(r.hash).toBe(compute(b, { clock }).hash)
    expect([r.sheets[0].cells.C1.spill, r.sheets[0].cells.C1.spillFill]).toEqual(['C1:C1048576', 0])
  })

  it('cells without spill bounds a spill too, spill without cells leaves arrays unbounded', () => {
    const s = openSession(book({ A1: { value: 1 } }), { clock, limits: { cells: 2_000 } })
    expect(stopped(() => s.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'SEQUENCE(1000,3)' } }])).max).toBe(2_000)
    expect(stopped(() => s.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'SUM(SEQUENCE(1000,3))' } }])).limit).toBe('cells')
    const t = openSession(book({ A1: { value: 1 } }), { clock, limits: { spill: 2_000 } })
    expect(t.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'SUM(SEQUENCE(1000,3))' } }])[0].value).toBe(4_501_500)
    expect(stopped(() => t.setCells([{ sheet: 'S', cell: 'B2', input: { formula: 'SEQUENCE(1000,3)' } }])).limit).toBe('spill')
  })

  it('spill in createInlineCalc: from the calculator or from the call', async () => {
    const calc = createInlineCalc({ limits: { spill: 1_000 } })
    const b = book({ C3: { formula: 'SEQUENCE(100,20)' } })
    const e = (await calc.compute(b, { clock }).catch((x: unknown) => x)) as CalcError
    expect(e).toBeInstanceOf(CalcError)
    expect([e.code, e.limit, e.max, e.size, e.cell]).toEqual(['limit', 'spill', 1_000, 2_000, { sheet: 'S', cell: 'C3' }])
    expect((await calc.compute(b, { clock, limits: { spill: 2_000 } })).hash).toBe(compute(b, { clock }).hash)
    expect((await calc.compute(book({ A1: { formula: 'SUM(SEQUENCE(100,20))' } }), { clock })).sheets[0].cells.A1.value).toBe(2_001_000)
    await calc.close()
  })

  it("a call's limits take the place of the calculator's field by field: the fields it does not give stay", async () => {
    const calc = createInlineCalc({ limits: { cells: 1_000 } })
    const b = book({ A1: { formula: 'SUM(SEQUENCE(2000))' } })
    const e = (await calc.compute(b, { clock, limits: { formulas: 10 } }).catch((x: unknown) => x)) as CalcError
    expect([e.code, e.limit, e.max]).toEqual(['limit', 'cells', 1_000])
    const s = await calc.openSession(book({ A1: { value: 1 } }), { clock, limits: { formulas: 10 } })
    const f = (await s.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'SUM(SEQUENCE(2000))' } }]).catch((x: unknown) => x)) as CalcError
    expect([f.code, f.limit, f.max]).toEqual(['limit', 'cells', 1_000])
    await calc.close()
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
    expect(() => compute(book({}), { limits: { spill: 1.5 } })).toThrow('opts.limits.spill: expected an integer from 1')
    expect(() => compute(book({}), { limits: { rows: 1 } as never })).toThrow('opts.limits: unknown field "rows"')
  })
})
