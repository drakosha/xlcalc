/**
 * Functions over two arrays and one cell (the sets `binom-sumx-held`, `sumx-cell-held`, `sumproduct-held` and
 * `tests-cell-held`, measured in Excel for Mac and Windows): a reference to one cell is read first, argument by
 * argument, its error the answer and a blank cell #VALUE!, while an array of one goes on to the sizes; SUMPRODUCT takes
 * each array's error before the sizes. And BINOM.DIST.RANGE past the factorials' range, by the binomial densities.
 */
import { describe, expect, it } from 'vitest'
import { compute, type BookInput, type CellValue } from '../src/index.js'

/** D: A1 blank, A7 `1/0`, A8 `NA()`, B1:B3 numbers, E1:E3 numbers with `1/0` in E2. */
const DATA: BookInput['sheets'][number] = {
  name: 'D',
  cells: { A7: { formula: '1/0' }, A8: { formula: 'NA()' }, B1: { value: 4 }, B2: { value: 8 }, B3: { value: 15 }, E1: { value: 1 }, E2: { formula: '1/0' }, E3: { value: 3 } },
}

const err = (code: string): CellValue => ({ error: code }) as CellValue

const CASES: [string, CellValue][] = [
  ['SUMX2MY2(D!A7,D!B1:B2)', err('DIV0')],
  ['SUMXMY2(D!B1:B2,D!A7)', err('DIV0')],
  ['SUMX2PY2(D!A1,D!B1:B2)', err('VALUE')],
  ['SUMX2MY2(D!A1,D!A7)', err('VALUE')],
  ['SUMX2MY2(D!A1,D!B1)', err('VALUE')],
  ['SUMX2MY2({#DIV/0!},{1,2})', err('NA')],
  ['SUMX2MY2({#N/A},D!A7)', err('DIV0')],
  ['SUMX2MY2(D!E1:E3,D!B1:B2)', err('NA')],
  ['CORREL({#DIV/0!},{1,2})', err('NA')],
  ['CORREL(D!A1,D!B1:B2)', err('VALUE')],
  ['CHISQ.TEST(D!A1,D!B1)', err('VALUE')],
  ['CHISQ.TEST({#DIV/0!},{1,2})', err('NA')],
  ['T.TEST(D!A1,D!B1:B3,3,1)', err('VALUE')],
  ['T.TEST({#DIV/0!},D!B1:B3,3,2)', err('NUM')],
  ['FORECAST(1,D!A1,D!B1:B2)', err('VALUE')],
  ['SUMPRODUCT(D!A7,D!B1:B2)', err('DIV0')],
  ['SUMPRODUCT(D!B1:B2,D!E1:E3)', err('DIV0')],
  ['SUMPRODUCT({#DIV/0!},{1,2})', err('DIV0')],
  ['SUMPRODUCT(D!A1,D!A7)', err('VALUE')],
  ['SUMPRODUCT(D!A1,D!B1)', err('VALUE')],
  ['SUMPRODUCT(D!B1:B2,D!B1:B3)', err('VALUE')],
]

const BINOM: [string, number][] = [
  ['BINOM.DIST.RANGE(171,0.5,85)', 0.06074976462616549],
  ['BINOM.DIST.RANGE(200,0.5,90,110)', 0.8626333491313694],
  ['BINOM.DIST.RANGE(1030,0.5,515)', 0.024855129936574464],
  ['BINOM.DIST.RANGE(1000000,0.5,499000,501000)', 0.9546076640906558],
  ['BINOM.DIST.RANGE(2147483646,0.5,1073741823)', 0.000017217699699243103],
]

describe.each(['closures', 'eval'] as const)('functions over two arrays and one cell (%s)', (execution) => {
  const cells = Object.fromEntries([...CASES, ...BINOM].map(([f], i) => [`A${i + 1}`, { formula: f }]))
  const got = compute({ sheets: [DATA, { name: 'S', cells }] }, { execution }).sheets[1].cells
  it.each(CASES.map(([f, v], i) => [f, v, i] as const))('%s', (_, want, i) => expect(got[`A${i + 1}`].value).toEqual(want))
  it.each(BINOM.map(([f, v], i) => [f, v, CASES.length + i] as const))('%s', (_, want, i) => {
    const v = got[`A${i + 1}`].value as number
    expect(Math.abs(v - want)).toBeLessThanOrEqual(1e-9 * Math.abs(want))
  })
  it('BINOM.DIST.RANGE at 2^31 − 1 trials is #NUM!, as BINOM.DIST', () => {
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { formula: 'BINOM.DIST.RANGE(2147483647,0.5,1)' } } }] }, { execution })
    expect(r.sheets[0].cells.A1.value).toEqual(err('NUM'))
  })
})
