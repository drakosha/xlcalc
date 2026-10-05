/**
 * Function behaviour the Excel measurement does not cover. The expected answers follow Excel's
 * documentation, or extend what the measured cases show (the 1900 calendar, fraction formats); they have
 * not been run in Excel.
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
  it('zero to the power of zero is #NUM!, to a negative power #DIV/0!', () => {
    expect(values(['0^0', '0^-1', '2^-1'])).toEqual([err('NUM'), err('DIV0'), 0.5])
  })

  it('ROUND on the decimal representation of any magnitude', () => {
    expect(values(['ROUND(1E-7,9)', 'ROUND(2.15,400)', 'ROUND(1.005,2)'])).toEqual([1e-7, 2.15, 1.01])
  })

  it('XLOOKUP over whole columns of sheets with different numbers of rows', () => {
    const t: Record<string, CellInput> = {}
    for (let r = 1; r <= 10; r++) t[`B${r}`] = { value: r * 10 }
    expect(values(['XLOOKUP("b",A:A,T!B:B)'], { A1: { value: 'a' }, A2: { value: 'b' } }, t)).toEqual([20])
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
