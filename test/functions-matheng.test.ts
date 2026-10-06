/**
 * Math and engineering functions beyond the Excel case set (`bench/coverage/sets/matheng.ts`): huge Bessel orders
 * answer at once, whole columns, and round trips through the number systems, Roman numerals and CONVERT.
 */
import { describe, expect, it } from 'vitest'
import { compute, type CellInput, type CellValue } from '../src/index.js'

const err = (code: Extract<CellValue, { error: unknown }>['error']): CellValue => ({ error: code })

function values(formulas: string[], cells: Record<string, CellInput> = {}): CellValue[] {
  const f: Record<string, CellInput> = {}
  formulas.forEach((x, i) => (f[`Z${i + 1}`] = { formula: x }))
  const r = compute({ sheets: [{ name: 'S', cells: { ...cells, ...f } }] }, { clock: { now: '2026-10-06T12:00:00Z', seed: 1 } })
  return formulas.map((_, i) => r.sheets[0].cells[`Z${i + 1}`].value)
}

describe('math and engineering functions beyond the Excel cases', () => {
  it('Bessel functions of huge orders answer at once: 0 where the value underflows, #NUM! where it overflows', () => {
    const t0 = performance.now()
    const v = values(['BESSELJ(1,1E10)', 'BESSELI(1,1E10)', 'BESSELK(1,1E10)', 'BESSELY(1,1E10)', 'BESSELK(1000,1E10)', 'BESSELJ(1E8,2E8)', 'BESSELI(-3,1E9)'])
    expect(performance.now() - t0).toBeLessThan(1000)
    expect(v).toEqual([0, 0, err('NUM'), err('NUM'), 0, 0, 0])
  })

  it('whole columns: only the occupied part is read, blank cells past it still count', () => {
    const cells: Record<string, CellInput> = { A1: { value: 4 }, A2: { value: 6 }, B1: { value: 2 }, B2: { value: 3 }, C1: { value: '1+i' }, C2: { value: 2 } }
    expect(values(['SUMPRODUCT(A:A,B:B)', 'SUMX2PY2(A:A,B:B)', 'IMSUM(C:C)', 'IMPRODUCT(C:C)', 'IMPRODUCT(C1:C2)', 'GCD(A:A)', 'LCM(A1:A2)', 'LCM(A:A)'], cells)).toEqual([26, 65, '3+i', '0', '2+2i', 2, 12, 0])
  })

  it('number systems: DEC2BIN, DEC2OCT and DEC2HEX and back give the number for the whole ten-digit range', () => {
    const ns = [-512, -511, -100, -1, 0, 1, 99, 255, 511]
    const formulas = ns.flatMap((n) => [`BIN2DEC(DEC2BIN(${n}))`, `OCT2DEC(DEC2OCT(${n * 1048576}))`, `HEX2DEC(DEC2HEX(${n * 1073741824}))`, `HEX2DEC(BIN2HEX(DEC2BIN(${n})))`])
    expect(values(formulas)).toEqual(ns.flatMap((n) => [n, n * 1048576, n * 1073741824, n]))
  })

  it('ROMAN in each form reads back with ARABIC; BASE reads back with DECIMAL', () => {
    const ns = Array.from({ length: 572 }, (_, i) => 1 + i * 7)
    const formulas = ns.flatMap((n) => [0, 1, 2, 3, 4].map((f) => `ARABIC(ROMAN(${n},${f}))`))
    expect(values(formulas)).toEqual(ns.flatMap((n) => [n, n, n, n, n]))
    expect(values(['DECIMAL(BASE(123456789012345,36),36)', 'DECIMAL(BASE(123456789,7,40),7)'])).toEqual([123456789012345, 123456789])
  })

  it('CONVERT there and back gives the number, prefixes included', () => {
    const units = ['lbm', 'u', 'Nmi', 'parsec', 'yr', 'mmHg', 'Torr', 'lbf', 'eV', 'HPh', 'BTU', 'PS', 'ga', 'uk_gal', 'Pica^3', 'us_acre', 'kbyte', 'Gibit', 'admkn', 'mph', 'km2', 'cm3']
    const base = ['g', 'g', 'm', 'm', 'sec', 'Pa', 'Pa', 'N', 'J', 'J', 'J', 'W', 'T', 'l', 'm3', 'm2', 'bit', 'byte', 'm/s', 'kn', 'ha', 'tsp']
    const v = values(units.map((u, i) => `CONVERT(CONVERT(3,"${u}","${base[i]}"),"${base[i]}","${u}")`)) as number[]
    for (const [i, x] of v.entries()) expect(x, units[i]).toBeCloseTo(3, 12)
    expect(values(['CONVERT(25,"C","F")', 'CONVERT(CONVERT(25,"C","Reau"),"Reau","C")', 'CONVERT(1,"m","kg")'])).toEqual([77, 25, err('NA')])
  })
})
