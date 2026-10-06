/**
 * The shared argument layer (`src/args.ts`), rule by rule: coercion by kind on real ranges of a computed
 * book, empty input, paired arrays, the answer rules (non-finite numbers), and the function specs derived from kinds.
 */
import { describe, expect, it } from 'vitest'
import { makeCaller, pairs, specOf, type FnDef, type Grid } from '../src/args.js'
import { createState } from '../src/engine.js'
import { E, isErr, Rng, type V } from '../src/values.js'

// sheet S: A1:A8 a mix, B1:B3 dates with a time, C1 an error, D1 blank, E1:E3 numbers
const state = createState(
  {
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1 }, A2: { value: '2' }, A3: { value: 'abc' }, A4: { value: true }, A6: { value: 5 }, A7: { value: false }, A8: { value: 3.5 },
          B1: { value: 45688.75 }, B2: { value: { date: '2025-02-01' } }, B3: { value: 45700.5 },
          C1: { formula: '1/0' },
          E1: { value: 1 }, E2: { value: 2 }, E3: { value: 3 },
        },
      },
    ],
  },
  { now: 46000, seed: 1 },
)
const cx = state.rt.cx
const rng = (r1: number, c1: number, r2: number, c2: number) => new Rng(0, r1, c1, r2, c2)
const MIX = rng(0, 0, 7, 0)
const DATES = rng(0, 1, 2, 1)
const ERR = rng(0, 2, 0, 2)
const NUMS = rng(0, 4, 2, 4)

/** Calls a one-argument (or repeated) function of the given kind that returns its coerced argument as JSON text. */
const coerce = (args: string, ...a: (V | Rng)[]) => {
  const r = makeCaller({ args, rep: args.length === 1 && 'NAZVD'.includes(args) ? 1 : undefined, fn: (x: unknown) => JSON.stringify(x) }, cx)(...a)
  return isErr(r) ? r : JSON.parse(r as string)
}

describe('argument layer', () => {
  it('rule 1: numbers (`n`, `i`, `x`), text, booleans', () => {
    expect(coerce('n', true)).toBe(1)
    expect(coerce('n', '3')).toBe(3)
    expect(coerce('n', null)).toBe(0)
    expect(coerce('n', 'abc')).toBe(E.VALUE)
    expect(coerce('n', E.NA)).toBe(E.NA)
    expect(coerce('i', -2.7)).toBe(-2)
    expect(coerce('x', true)).toBe(E.VALUE)
    expect(coerce('x', '4')).toBe(4)
    expect(coerce('s', 0.1 + 0.2)).toBe('0.3')
    expect(coerce('s', false)).toBe('FALSE')
    expect(coerce('b', 'true')).toBe(true)
    expect(coerce('b', 2)).toBe(true)
    expect(coerce('b', 'abc')).toBe(E.VALUE)
    expect(coerce('v', E.NA)).toBe(E.NA)
    expect(makeCaller({ args: 'a', fn: (v: V) => (isErr(v) ? 'error passed' : 'value') }, cx)(E.NA)).toBe('error passed')
  })

  it('rule 1: lists from ranges and arguments (`N` SUM, `A` AVERAGEA, `Z` strict)', () => {
    expect(coerce('N', MIX)).toEqual([1, 5, 3.5])
    expect(coerce('N', '2', true)).toEqual([2, 1])
    expect(coerce('N', 'abc')).toBe(E.VALUE)
    expect(coerce('N', MIX, ERR)).toBe(E.DIV0)
    expect(coerce('A', MIX)).toEqual([1, 0, 0, 1, 5, 0, 3.5])
    expect(coerce('A', 'abc')).toBe(E.VALUE)
    expect(coerce('Z', NUMS)).toEqual([1, 2, 3])
    expect(coerce('Z', MIX)).toBe(E.VALUE)
    expect(coerce('Z', true)).toBe(E.VALUE)
  })

  it('rule 1: values with their origin (`V`), grids (`G`, `M`)', () => {
    expect(coerce('V', rng(0, 0, 2, 0), 'x', E.NA)).toEqual([[1, '2', 'abc'], 'x', { code: 'NA' }])
    const whole = makeCaller({ args: 'G', fn: (g: Grid) => `${g.rows}x${g.cols} ${g.h}x${g.w}` }, cx)(new Rng(0, 0, 4, 1048575, 4))
    expect(whole).toBe('1048576x1 8x1')
    expect(makeCaller({ args: 'M', fn: () => 'called' }, cx)(rng(0, 2, 2, 2))).toBe(E.DIV0)
  })

  it('rule 2: date and time text in date arguments, the date range', () => {
    expect(coerce('d', '2025-01-31')).toBe(45688)
    expect(coerce('d', '1/2/2025')).toBe(45659)
    expect(coerce('d', '2/29/1900')).toBe(60)
    expect(coerce('d', '45688')).toBe(45688)
    expect(coerce('d', 45688.75)).toBe(45688)
    expect(coerce('d', -1)).toBe(E.NUM)
    expect(coerce('d', 2958466)).toBe(E.NUM)
    expect(coerce('d', '2/30/2025')).toBe(E.VALUE)
    expect(coerce('t', '18:30')).toBeCloseTo(18.5 / 24, 12)
    expect(coerce('t', '2025-01-31 6:00 PM')).toBe(45688.75)
  })

  it('rule 3: empty input gives the declared error', () => {
    const avg: FnDef = { args: 'N', rep: 1, empty: 'DIV0', fn: (xs: number[]) => xs.length }
    expect(makeCaller(avg, cx)(rng(0, 3, 4, 3))).toBe(E.DIV0)
    expect(makeCaller(avg, cx)(rng(0, 3, 4, 3), 7)).toBe(1)
  })

  it('rule 4: number pairs drop non-numbers; different sizes are #N/A', () => {
    const g = (rows: number, cols: number, v: V[]): Grid => ({ rows, cols, h: rows, w: cols, v })
    expect(pairs(g(4, 1, [1, 'a', 3, null]), g(4, 1, [10, 20, true, 40]))).toEqual([[1], [10]])
    expect(pairs(g(2, 1, [1, 2]), g(1, 2, [3, 4]))).toEqual([[1, 2], [3, 4]])
    expect(pairs(g(2, 1, [1, 2]), g(3, 1, [1, 2, 3]))).toBe(E.NA)
  })

  it('rule 5: a non-finite answer is #NUM!, an exception #VALUE!', () => {
    expect(makeCaller({ args: 'n', fn: (x: number) => x / 0 }, cx)(1)).toBe(E.NUM)
    expect(makeCaller({ args: 'n', fn: () => Math.sqrt(-1) }, cx)(1)).toBe(E.NUM)
    expect(makeCaller({ args: '', fn: () => ([] as number[]).reduce((a, b) => a + b) }, cx)()).toBe(E.VALUE)
  })

  it('rule 6: holidays as whole serial numbers; text or a boolean in the list is #VALUE!', () => {
    expect(coerce('D', DATES)).toEqual([45688, 45689, 45700])
    expect(coerce('D', '2025-01-31')).toEqual([45688])
    expect(coerce('D', MIX)).toBe(E.VALUE)
  })

  it('repeated arguments arrive as one array after the fixed ones', () => {
    const f = makeCaller({ args: 'n? N', rep: 1, fn: (first: number | undefined, rest: number[]) => `${first}|${rest.join(',')}` }, cx)
    expect(f(1, NUMS, 4)).toBe('1|1,2,3,4')
  })

  it('specs: argument counts, range positions and parity follow the kinds', () => {
    const sumifs = specOf('SUMIFS', { args: 'R R a', rep: 2, fn: () => 0 })
    expect([sumifs.min, sumifs.max, sumifs.parity]).toEqual([3, 255, 'odd'])
    expect([0, 1, 2, 3, 4].map(sumifs.range)).toEqual([true, true, false, true, false])
    const countifs = specOf('COUNTIFS', { args: 'R a', rep: 2, fn: () => 0 })
    expect([countifs.min, countifs.max, countifs.parity]).toEqual([2, 254, 'even'])
    expect(specOf('LOG', { args: 'n n?', fn: () => 0 })).toMatchObject({ min: 1, max: 2 })
    expect(() => specOf('BAD', { args: 'q', fn: () => 0 })).toThrow(/unknown argument kind q/)
  })
})
