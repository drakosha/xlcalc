/**
 * Financial and date functions beyond the Excel case set: the shortcuts the engine takes (whole weeks in
 * WORKDAY and NETWORKDAYS, closed forms in DB and CUMIPMT) against plain counting, and arguments that would
 * make a naive loop run for millions of steps.
 */
import { describe, expect, it } from 'vitest'
import { mulberry32 } from '../src/synth.js'
import { compute, validateFormula, type CellInput, type CellValue } from '../src/index.js'

/** Formulas computed on sheet S with the given cells; values in order. */
function values(formulas: string[], cells: Record<string, CellInput> = {}): CellValue[] {
  const f: Record<string, CellInput> = { ...cells }
  formulas.forEach((x, i) => (f[`Z${i + 1}`] = { formula: x }))
  const r = compute({ sheets: [{ name: 'S', cells: f }] }, { clock: { now: '2026-10-05T12:00:00Z', seed: 1 } })
  return formulas.map((_, i) => r.sheets[0].cells[`Z${i + 1}`].value)
}

/** Weekend days (0 Sunday) of a WORKDAY.INTL code 1–7 or 11–17; serial 1 is a Sunday in Excel. */
const weekendOf = (code: number) => (code < 10 ? [(code + 5) % 7, (code + 6) % 7] : [code - 11])
const isWork = (s: number, code: number, hol: Set<number>) => !weekendOf(code).includes((s + 6) % 7) && !hol.has(s)

describe('workdays against day-by-day counting', () => {
  const rnd = mulberry32(11)
  const cases = Array.from({ length: 60 }, () => {
    const start = 40000 + Math.floor(rnd() * 3000)
    const days = Math.floor(rnd() * 400) - 200
    const code = [1, 2, 7, 11, 14, 17][Math.floor(rnd() * 6)]
    const hol = Array.from({ length: 12 }, () => start + Math.floor(rnd() * 600) - 300)
    return { start, days, code, hol }
  })
  const cells: Record<string, CellInput> = {}
  cases.forEach((c, i) => c.hol.forEach((h, j) => (cells[`${String.fromCharCode(65 + j)}${i + 1}`] = { value: h })))
  const got = values(
    cases.flatMap((c, i) => [`WORKDAY.INTL(${c.start},${c.days},${c.code},A${i + 1}:L${i + 1})`, `NETWORKDAYS.INTL(${c.start},${c.start + c.days},${c.code},A${i + 1}:L${i + 1})`]),
    cells,
  )

  it('WORKDAY.INTL and NETWORKDAYS.INTL with weekends and holidays', () => {
    cases.forEach((c, i) => {
      const hol = new Set(c.hol)
      let d = c.start
      for (let left = Math.abs(c.days); left > 0; ) {
        d += Math.sign(c.days)
        if (isWork(d, c.code, hol)) left--
      }
      const [a, b] = c.days >= 0 ? [c.start, c.start + c.days] : [c.start + c.days, c.start]
      let n = 0
      for (let s = a; s <= b; s++) if (isWork(s, c.code, hol)) n++
      expect([got[2 * i], got[2 * i + 1]], JSON.stringify(c)).toEqual([d, c.days >= 0 ? n : -n])
    })
  })
})

describe('closed forms against the period-by-period computation', () => {
  it('DB: the book value declines geometrically after the first year', () => {
    const [cost, salvage, life, month] = [1_000_000, 100_000, 6, 7]
    const rate = Math.round((1 - Math.pow(salvage / cost, 1 / life)) * 1000) / 1000
    let total = (cost * rate * month) / 12
    const expected = [total]
    for (let p = 2; p <= life + 1; p++) {
      const dep = ((cost - total) * rate * (p > life ? 12 - month : 12)) / 12
      total += dep
      expected.push(dep)
    }
    const got = values(expected.map((_, i) => `DB(${cost},${salvage},${life},${i + 1},${month})`)) as number[]
    got.forEach((g, i) => expect(g).toBeCloseTo(expected[i], 6))
  })

  it('CUMIPMT and CUMPRINC equal the sums of IPMT and PPMT', () => {
    const sums = (s: number, e: number, t: number, part: 'IPMT' | 'PPMT') =>
      values(Array.from({ length: e - s + 1 }, (_, k) => `${part}(0.0075,${s + k},360,125000,0,${t})`)).reduce((a, b) => (a as number) + (b as number), 0) as number
    for (const [s, e, t] of [[1, 12, 0], [5, 30, 0], [1, 1, 1], [1, 24, 1], [7, 7, 1], [13, 360, 0]]) {
      const [ci, cp] = values([`CUMIPMT(0.0075,360,125000,${s},${e},${t})`, `CUMPRINC(0.0075,360,125000,${s},${e},${t})`]) as number[]
      expect(ci).toBeCloseTo(sums(s, e, t, 'IPMT'), 6)
      expect(cp).toBeCloseTo(sums(s, e, t, 'PPMT'), 6)
    }
  })
})

describe('large period counts and spans stay fast', () => {
  it('millions of periods or days, whole-column holidays, coupons from 1900 to 9999', () => {
    const t0 = performance.now()
    const got = values([
      'WORKDAY(1,2000000)',
      'WORKDAY(1,3000000)',
      'NETWORKDAYS(1,2958465)',
      'NETWORKDAYS(1,2958465,Y:Y)',
      'DB(1000000,100000,1E9,1E9)',
      'CUMIPMT(1E-6,1E6,1000,1,1E6,0)',
      'AMORDEGRC(10000,DATE(2008,8,19),DATE(2008,12,31),0,1E9,1E-9,0)',
      'ACCRINT(DATE(1900,3,1),DATE(1900,6,1),DATE(9999,12,1),0.1,1000,4)',
      'PRICE(DATE(1900,3,1),DATE(9999,12,1),0.05,0.06,100,4)',
      'YIELD(DATE(1900,3,1),DATE(9999,12,1),0.05,90,100,4)',
      // an issue date in a quasi-coupon period that begins before 1900 is #NUM! in Excel
      'ODDFPRICE(DATE(1900,4,1),DATE(9999,12,1),DATE(1900,3,15),DATE(1900,6,1),0.05,0.06,100,4)',
    ])
    expect(performance.now() - t0).toBeLessThan(3000)
    expect(got[1]).toEqual({ error: 'NUM' })
    expect(got[2]).toBe(got[3])
    for (const [i, v] of got.entries()) if (i !== 1) expect(typeof v, String(i)).toBe('number')
  })
})

describe('argument counts as Excel accepts them', () => {
  it('ACCRINTM needs par (Excel rejects a workbook with ACCRINTM of three arguments)', () => {
    expect(validateFormula('ACCRINTM(39539,39614,0.1)')).toMatchObject({ ok: false, code: 'syntax' })
    expect(validateFormula('ACCRINTM(39539,39614,0.1,1000)').ok).toBe(true)
  })
})
