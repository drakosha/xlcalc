/** Reproductions from the 21.10 review of the functions: sizes past the argument limit and work outside the budget. */
import { describe, expect, it } from 'vitest'
import { CalcError, compute, createInlineCalc, type BookInput } from '../src/index.js'

type Cells = BookInput['sheets'][number]['cells']
const book = (cells: Cells): BookInput => ({ sheets: [{ name: 'S', cells }] })
const one = (formula: string, execution: 'closures' | 'eval') => compute(book({ A1: { formula } }), { execution, display: false }).sheets[0].cells.A1.value

/** Milliseconds a call takes. */
function took(f: () => unknown): number {
  const t0 = performance.now()
  f()
  return performance.now() - t0
}

async function stopped(p: Promise<unknown>): Promise<CalcError> {
  try {
    await p
  } catch (e) {
    expect(e).toBeInstanceOf(CalcError)
    return e as CalcError
  }
  throw new Error('expected CalcError')
}

describe.each(['closures', 'eval'] as const)('review 21.10 functions, %s', (execution) => {
  it('fits LINEST, TREND, LOGEST and GROWTH over more observations than one call can take as arguments', () => {
    const n = 140000
    const cells: Cells = {}
    for (let i = 1; i <= n; i++) {
      cells[`A${i}`] = { value: 2 * i + 1 }
      cells[`B${i}`] = { value: i }
      cells[`C${i}`] = { value: Math.exp(i / n) }
    }
    const [ya, xb, yc] = ['A', 'B', 'C'].map((c) => `${c}1:${c}${n}`)
    Object.assign(cells, {
      D1: { formula: `INDEX(LINEST(${ya},${xb}),1,1)` },
      D2: { formula: `INDEX(LINEST(${ya},${xb}),1,2)` },
      D3: { formula: `TREND(${ya},${xb},${n + 1})` },
      D4: { formula: `INDEX(LOGEST(${yc},${xb}),1,1)` },
      D5: { formula: `GROWTH(${yc},${xb},${n + 1})` },
      D6: { formula: `SLOPE(${ya},${xb})` },
    })
    const out = compute(book(cells), { execution, display: false }).sheets[0].cells
    expect(out.D1.value).toBeCloseTo(2, 9)
    expect(out.D2.value).toBeCloseTo(1, 6)
    expect(out.D3.value).toBeCloseTo(2 * (n + 1) + 1, 6)
    expect(out.D4.value).toBeCloseTo(Math.exp(1 / n), 9)
    expect(out.D5.value).toBeCloseTo(Math.exp((n + 1) / n), 9)
    expect(out.D6.value).toBeCloseTo(2, 9)
  })

  it('FREQUENCY counts by a search over the sorted bins: the answers of the scan, 100,000 values and bins at once', () => {
    expect(one('TEXTJOIN(",",,FREQUENCY({1,2,2,3,5,7},{3,2,2,9,0}))', execution)).toBe('1,3,0,2,0,0')
    expect(one('TEXTJOIN(",",,FREQUENCY({-1,0,4,4,8,9.5},{4,-0.5,4,9}))', execution)).toBe('3,1,0,1,1')
    // the scan made 5·10^9 comparisons here (seconds); the search 1.7·10^6
    let v: unknown
    const ms = took(() => (v = one('SUM(FREQUENCY(SEQUENCE(100000),SEQUENCE(100000)))', execution)))
    expect(v).toBe(100000)
    expect(ms).toBeLessThan(1500)
  })

  it('AMORDEGRC over millions of periods with a non-zero depreciation is fast and stops at the deadline', async () => {
    // each period used to round through decimal text: 10^7 periods took about 3 s
    let v: unknown
    const ms = took(() => (v = one('AMORDEGRC(1E15,40000,40001,0,1E7,1E-9,0)', execution)))
    expect(typeof v).toBe('number')
    expect(ms).toBeLessThan(1000)
    const calc = createInlineCalc()
    const t0 = performance.now()
    const e = await stopped(calc.compute(book({ A1: { formula: 'AMORDEGRC(1E15,40000,40001,0,9E8,1E-9,0)' } }), { execution, timeoutMs: 100 }))
    expect(e.code).toBe('soft_timeout')
    expect(performance.now() - t0).toBeLessThan(1500)
  })

  it('FILTER keeps its answer within limits.cells: an include true past the occupied rows selects the whole column', () => {
    const cells: Cells = { A1: { value: 1 }, A2: { value: 2 }, B1: { value: 'x' }, C1: { value: 'done' } }
    for (const formula of ['ROWS(FILTER(A:A,B:B=B:B))', 'FILTER(A:A,ISBLANK(B:B))', 'SUM(FILTER(A:C,C:C<>"done"))']) {
      let e: unknown
      try {
        compute(book({ ...cells, E5: { formula } }), { execution, limits: { cells: 100_000 } })
      } catch (x) {
        e = x
      }
      expect(e, formula).toBeInstanceOf(CalcError)
      const c = e as CalcError
      expect([c.code, c.limit, c.max, c.cell], formula).toEqual(['limit', 'cells', 100_000, { sheet: 'S', cell: 'E5' }])
    }
    const fits = book({ ...cells, E5: { formula: 'ROWS(FILTER(A:A,B:B=B:B))' }, E6: { formula: 'SUM(FILTER(A1:A2,A1:A2>1))' }, E7: { formula: 'ROWS(UNIQUE(A:A))' } })
    expect(compute(fits, { execution, limits: { cells: 2_000_000 } }).hash).toBe(compute(fits, { execution }).hash)
    expect(compute(fits, { execution, display: false }).sheets[0].cells.E5.value).toBe(1048576)
  })

  it('TAKE and DROP of a whole column answer a reference without an index per row', () => {
    const cells: Cells = { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 4 }, B2: { value: 8 } }
    const formulas = ['SUM(DROP(A:A,1))', 'SUM(TAKE(A:B,,1))', 'ROWS(DROP(A:A,-2))', 'SUM(TAKE(A:A,-1048575))', 'ROWS(DROP(A:B,1048576))', 'SUM(DROP(A:B,1,-1))']
    for (const [i, formula] of formulas.entries()) cells[`D${i + 1}`] = { formula }
    const out = compute(book(cells), { execution, display: false }).sheets[0].cells
    expect(formulas.map((_, i) => out[`D${i + 1}`].value)).toEqual([6, 7, 1048574, 6, { error: 'CALC' }, 6])
    // each of these built a 1,048,576-entry index array per evaluation (20 ms); 200 of them took seconds
    const many: Cells = { A1: { value: 1 } }
    for (let r = 1; r <= 200; r++) many[`C${r}`] = { formula: r % 2 ? 'SUM(DROP(A:A,1))' : 'SUM(TAKE(A:A,,1))' }
    expect(took(() => compute(book(many), { execution, display: false }))).toBeLessThan(1000)
  })
})
