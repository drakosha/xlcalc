/**
 * Statistical functions beyond their Excel cases (`test/excel/cases/stat.json`): the special functions
 * against mpmath at 40 digits, the ceilings that turn a long iteration into an answer or #NUM! instead of a
 * hang, and the compatibility names answering as their modern versions.
 */
import { describe, expect, it } from 'vitest'
import { compute, type CellInput } from '../src/index.js'
import { gammafn } from '../src/functions/jstat.js'
import { chisqInvRt, gammaInv } from '../src/functions/statistical/inverse.js'
import { gammaPQ, ibeta, lgamma, normInv } from '../src/functions/statistical/special.js'
import { numbersOf } from '../src/functions/statistical/common.js'
import { E } from '../src/values.js'

const near = (x: number, ref: number, tol = 1e-14) => expect(Math.abs(x - ref), `${x} vs ${ref}`).toBeLessThanOrEqual(tol * Math.max(1, Math.abs(ref)))

describe('special functions against mpmath', () => {
  it("gammafn: Cody's rational fit and reflection", () => {
    // jStat started the denominator at 0 (GAMMA(1.5) was 0.88622692794) and nudged the reflection by 3.6e-16
    for (const [x, ref] of [
      [0.25, 3.6256099082219083],
      [1.5, 0.88622692545275801],
      [-0.5, -3.5449077018110321],
      [-2.000001, -499999.53853888008],
      [1e-5, 99999.422794225559],
      [170.5, 5.5620924145599996e305],
    ])
      near(gammafn(x), ref)
    expect(gammafn(-3)).toBe(Infinity)
  })

  it('lgamma on both sides of the Stirling series', () => {
    near(lgamma(0.001), 6.9071788853838537)
    near(lgamma(9.5), 11.689333420797268)
    near(lgamma(10.5), 13.940625219403764)
    near(lgamma(1e10), 220258509288.81058)
  })

  it('incomplete gamma: the series, the continued fraction, far tails', () => {
    near(gammaPQ(2000, 2000 + Math.sqrt(2000))[0], 0.84136459489783147)
    near(gammaPQ(3, 200)[1], 2.7956093736608835e-83, 1e-96)
    expect(gammaPQ(3, 200)[1] / 2.7956093736608835e-83).toBeCloseTo(1, 12)
  })

  it('incomplete beta: the continued fraction, a far tail, and a ≫ b where the fraction cancels', () => {
    near(ibeta(0.96, 0.04, 1000, 30)[0], 0.02712433112361332)
    expect(ibeta(1e-8, 1 - 1e-8, 2.5, 3.5)[0] / 1.0864977254389269e-19).toBeCloseTo(1, 12)
    // the t distribution with 10^10 degrees of freedom at 1.959999998, two-tailed
    const t2 = 1.959999998 ** 2
    near(ibeta(1e10 / (1e10 + t2), t2 / (1e10 + t2), 5e9, 0.5)[0], 0.049995790557933518, 1e-13)
  })

  it('the normal quantile in both tails', () => {
    near(normInv(1e-20), -9.2623400897984076)
    near(normInv(0.3), -0.52440051270804082)
    near(normInv(1e-300), -37.0470962993612, 1e-13)
    near(normInv(1 - 1e-10, 1e-10), 6.361340902404056, 1e-13)
  })
})

describe('numbers of values', () => {
  it('skip text, booleans and blanks; the first error in order is the answer, at the selected positions too', () => {
    expect(numbersOf([1, 'a', true, null, 2])).toEqual([1, 2])
    expect(numbersOf([1, 'a', E.NA, 2, E.DIV0])).toBe(E.NA)
    const vs = [5, E.DIV0, 'x', 7, E.NA]
    expect(numbersOf(vs, [3, 0, 2])).toEqual([7, 5])
    expect(numbersOf(vs, [3, 4, 1])).toBe(E.NA)
    expect(numbersOf(vs, [])).toEqual([])
  })
})

describe('iteration ceilings', () => {
  it('shapes of 10^10 and more finish in milliseconds', () => {
    const t0 = performance.now()
    expect(Number.isFinite(chisqInvRt(0.05, 1e10) as number)).toBe(true)
    expect(Number.isFinite(gammaInv(0.7, 0.3, 1e10))).toBe(true)
    expect(Number.isFinite(gammaPQ(5e9, 5e9 + 1e5)[0])).toBe(true)
    expect(Number.isFinite(ibeta(0.5, 0.5, 5e9, 5e9)[0])).toBe(true)
    expect(performance.now() - t0).toBeLessThan(1000)
  })

  it('a book of extreme calls computes without hanging', () => {
    const cells: Record<string, CellInput> = {}
    const formulas = ['CHISQ.DIST(1E+9,1E+10,TRUE)', 'GAMMA.DIST(1E+9,1E+9,1,TRUE)', 'BINOM.DIST(500000000,1000000000,0.5,TRUE)', 'T.DIST(1.96,1E+10,TRUE)', 'F.DIST(2,1E+10,1E+10,TRUE)', 'BETA.INV(0.3,1E+10,2)', 'BINOM.DIST(5,1E+10,0.5,TRUE)']
    formulas.forEach((f, i) => (cells[`A${i + 1}`] = { formula: f }))
    const t0 = performance.now()
    const out = compute({ sheets: [{ name: 'S', cells }] }).sheets[0].cells
    expect(performance.now() - t0).toBeLessThan(2000)
    expect(out.A7.value).toEqual({ error: 'NUM' })
    for (let i = 1; i <= 6; i++) expect(typeof out[`A${i}`].value, formulas[i - 1]).toBe('number')
  })
})

describe('compatibility names', () => {
  it('answer as their modern versions', () => {
    const data: Record<string, CellInput> = {}
    ;[3, 7, 7, 2, -1.5, 10].forEach((v, i) => (data[`A${i + 1}`] = { value: v }))
    ;[1.5, 6, 8, 2.5, 0, 9].forEach((v, i) => (data[`B${i + 1}`] = { value: v }))
    const pairs = [
      ['STDEV(A1:A6)', 'STDEV.S(A1:A6)'],
      ['VARP(A1:A6)', 'VAR.P(A1:A6)'],
      ['MODE(A1:A6)', 'MODE.SNGL(A1:A6)'],
      ['RANK(7,A1:A6)', 'RANK.EQ(7,A1:A6)'],
      ['PERCENTILE(A1:A6,0.3)', 'PERCENTILE.INC(A1:A6,0.3)'],
      ['QUARTILE(A1:A6,3)', 'QUARTILE.INC(A1:A6,3)'],
      ['COVAR(A1:A6,B1:B6)', 'COVARIANCE.P(A1:A6,B1:B6)'],
      ['TTEST(A1:A6,B1:B6,2,1)', 'T.TEST(A1:A6,B1:B6,2,1)'],
      ['CHIDIST(3,4)', 'CHISQ.DIST.RT(3,4)'],
      ['NORMSDIST(1.2)', 'NORM.S.DIST(1.2,TRUE)'],
      ['LOGNORMDIST(4,1,0.5)', 'LOGNORM.DIST(4,1,0.5,TRUE)'],
      ['HYPGEOMDIST(1,4,8,20)', 'HYPGEOM.DIST(1,4,8,20,FALSE)'],
      ['TDIST(1.5,5,2)', 'T.DIST.2T(1.5,5)'],
      ['TINV(0.1,5)', 'T.INV.2T(0.1,5)'],
    ]
    pairs.forEach(([a, b], i) => {
      data[`C${i + 1}`] = { formula: a }
      data[`D${i + 1}`] = { formula: b }
    })
    const out = compute({ sheets: [{ name: 'S', cells: data }] }).sheets[0].cells
    pairs.forEach(([a], i) => {
      expect(typeof out[`C${i + 1}`].value, a).toBe('number')
      expect(out[`C${i + 1}`].value, a).toBe(out[`D${i + 1}`].value)
    })
  })
})
