/**
 * The measurement's case generator: the array-size guard over every case, and the set's completeness.
 */
import { describe, expect, it } from 'vitest'
import { arity, buildCases, parseKinds } from './cases.js'
import { ARRAY_SIZE, SIGNATURES, SIZE_LIMIT } from './signatures.js'

const set = buildCases()

describe('coverage measurement cases', () => {
  it('only small number literals in array-size positions', () => {
    const bad: string[] = []
    for (const c of set.cases) {
      for (const i of ARRAY_SIZE[c.fn] ?? []) {
        if (i >= c.args.length) continue
        const a = c.args[i]
        if (typeof a !== 'number' || !Number.isFinite(a) || Math.abs(a) > SIZE_LIMIT) bad.push(c.formula)
      }
    }
    expect(bad).toEqual([])
    // in the book the size argument is written as that literal
    for (const c of set.cases.filter((x) => x.fn in ARRAY_SIZE)) expect(c.formula).not.toMatch(/MUNIT\((?!-?[\d.]+\))/)
  })

  it('array functions with whole-number arguments are sorted out: a size or an index', () => {
    // CHOOSECOLS, CHOOSEROWS and SORT take column and row numbers: the array is no larger than the input
    const index = ['CHOOSECOLS', 'CHOOSEROWS', 'SORT']
    const withInt = SIGNATURES.filter((s) => s.picks && parseKinds(s).some((k) => k.k === 'i')).map((s) => s.name)
    expect(withInt.sort()).toEqual([...Object.keys(ARRAY_SIZE), ...index].sort())
  })

  it("argument counts within Excel's, 1 to 25 cases per function", () => {
    const per = new Map<string, number>()
    for (const c of set.cases) per.set(c.fn, (per.get(c.fn) ?? 0) + 1)
    expect(per.size).toBe(402)
    for (const n of per.values()) expect(n).toBeGreaterThan(0)
    for (const n of per.values()) expect(n).toBeLessThanOrEqual(25)
    for (const s of SIGNATURES) {
      const { min, max } = arity(s)
      for (const c of set.cases.filter((x) => x.fn === s.name)) expect(c.args.length >= min && c.args.length <= max).toBe(true)
    }
  })

  it('determinism: the same seed gives the same cases', () => {
    expect(buildCases().cases).toEqual(set.cases)
  })
})
