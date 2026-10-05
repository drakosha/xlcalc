/**
 * Function behaviour the Excel measurement does not cover, mostly where this engine departs from formula.js.
 * The expected answers follow Excel's documentation, or extend what the measured cases show (the 1900
 * calendar, fraction formats); they have not been run in Excel. Also guards against whole-column ranges:
 * the referenced size, no call-stack overflow.
 */
import { describe, expect, it } from 'vitest'
import { compute, type CellInput, type CellValue } from '../src/index.js'

const err = (code: Extract<CellValue, { error: unknown }>['error']): CellValue => ({ error: code })

/** Computes formulas over sheet S (cells as given) and returns their values in order. */
function values(formulas: string[], cells: Record<string, CellInput> = {}, other: Record<string, CellInput> = {}): CellValue[] {
  const f: Record<string, CellInput> = {}
  formulas.forEach((x, i) => (f[`Z${i + 1}`] = { formula: x }))
  const r = compute({ sheets: [{ name: 'S', cells: { ...cells, ...f } }, { name: 'T', cells: other }] }, { clock: { now: '2026-10-05T12:00:00Z', seed: 1 } })
  return formulas.map((_, i) => r.sheets[0].cells[`Z${i + 1}`].value)
}

describe('functions beyond the measurement', () => {
  it('zero to the power of zero is #NUM!, to a negative power #DIV/0! (POWER and ^)', () => {
    expect(values(['POWER(0,0)', 'POWER(0,-1)', '0^0', '0^-1', '2^-1'])).toEqual([err('NUM'), err('DIV0'), err('NUM'), err('DIV0'), 0.5])
  })

  it('ROUND on the decimal representation of any magnitude', () => {
    expect(values(['ROUND(1E-7,9)', 'ROUND(2.15,400)', 'ROUND(1.005,2)'])).toEqual([1e-7, 2.15, 1.01])
  })

  it('LOG: base 1 is #DIV/0!, base 0 and a non-positive number #NUM!', () => {
    expect(values(['LOG(8,1)', 'LOG(8,0)', 'LOG(0)', 'LOG(-1,2)', 'LOG(8,2)'])).toEqual([err('DIV0'), err('NUM'), err('NUM'), err('NUM'), 3])
  })

  it('rounding on the decimal representation: TRUNC, FIXED, CEILING and FLOOR', () => {
    expect(values(['TRUNC(4.35,2)', 'TRUNC(-4.35,1)', 'FIXED(1.005,2)', 'FIXED(-1234.567,-1)', 'CEILING(0.3,0.1)', 'FLOOR(0.3,0.1)'])).toEqual([4.35, -4.3, '1.01', '-1,230', 0.3, 0.3])
  })

  it('COUNTA counts empty text, COUNTBLANK counts it as blank', () => {
    const cells = { A1: { formula: '""' }, A2: { value: 1 } }
    expect(values(['COUNTA(A1:A3)', 'COUNTBLANK(A1:A3)'], cells)).toEqual([2, 2])
  })

  it('whole columns and rows have their referenced size', () => {
    expect(values(['ROWS(A:A)', 'COLUMNS(1:1)', 'COUNTBLANK(A:A)', 'ROWS(A1:B3)'], { A1: { value: 1 }, A5: { value: 2 } })).toEqual([1048576, 16384, 1048574, 3])
  })

  it('XLOOKUP over whole columns of sheets with different numbers of rows', () => {
    const t: Record<string, CellInput> = {}
    for (let r = 1; r <= 10; r++) t[`B${r}`] = { value: r * 10 }
    expect(values(['XLOOKUP("b",A:A,T!B:B)'], { A1: { value: 'a' }, A2: { value: 'b' } }, t)).toEqual([20])
  })

  it('MAX, MIN and MEDIAN over a column of 200,000 numbers', () => {
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 200_000; r++) cells[`A${r}`] = { value: r }
    expect(values(['MAX(A:A)', 'MIN(A:A)', 'MEDIAN(A:A)'], cells)).toEqual([200000, 1, 100000.5])
  })

  it('text: TRIM removes spaces only, UNICODE reads a code point, LEFT and RIGHT reject a negative count, REPT stops at 32,767 characters', () => {
    expect(values(['TRIM("  a \tb  ")', 'UNICODE("😀")', 'LEFT("abc",-1)', 'RIGHT("abc",-1)', 'LEN(REPT("ab",16384))', 'REPT("ab",16383)=REPT("ab",16383)'])).toEqual(['a \tb', 128512, err('VALUE'), err('VALUE'), err('VALUE'), true])
  })

  it('PROB pairs each value with its own probability (unsorted values)', () => {
    const cells = { A1: { value: 3 }, A2: { value: 1 }, A3: { value: 2 }, B1: { value: 0.5 }, B2: { value: 0.2 }, B3: { value: 0.3 } }
    expect(values(['PROB(A1:A3,B1:B3,1)', 'PROB(A1:A3,B1:B3,2,3)', 'PROB(A1:A3,B1:B2,1)'], cells)).toEqual([0.2, 0.8, err('NA')])
  })

  it('F.DIST.RT truncates the degrees of freedom', () => {
    const [a, b] = values(['F.DIST.RT(15.2069,6.9,4.5)', 'F.DIST.RT(15.2069,6,4)'])
    expect(a).toBe(b)
  })

  it('financial: zero divisors are #DIV/0!, PDURATION needs positive values', () => {
    expect(values(['ISPMT(0.1,1,0,100)', 'NPV(-1,100,200)', 'PDURATION(0.1,-100,200)'])).toEqual([err('DIV0'), err('DIV0'), err('NUM')])
  })

  it('TEXT: fractions and the empty format', () => {
    expect(values(['TEXT(4.34,"# ?/?")', 'TEXT(0.75,"?/?")', 'TEXT(2.5,"# ?/4")', 'TEXT(-4.34,"# ?/?")', 'TEXT(4,"# ?/?")', 'TEXT(1.25,"# ??/??")', 'TEXT(5,"")'])).toEqual([
      '4 1/3',
      '3/4',
      '2 2/4',
      '-4 1/3',
      '4    ',
      '1  1/4 ',
      '',
    ])
  })

  it('DATE and EOMONTH in the 1900 calendar', () => {
    expect(values(['DATE(1900,2,29)', 'DATE(1900,1,0)', 'DATE(1900,1,-1)', 'EOMONTH(0,0)', 'EOMONTH(60,0)', 'EOMONTH(1,-1)', 'EOMONTH("2/29/1900",0)'])).toEqual([60, 0, err('NUM'), 31, 59, err('NUM'), 59])
  })
})
