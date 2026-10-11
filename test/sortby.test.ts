/** SORTBY with a by-array of one value, as Excel answers it (`bench/coverage/sets/sortby-one-held.ts`). */
import { describe, expect, it } from 'vitest'
import { compute, type BookInput } from '../src/index.js'

const clock = { now: '2026-10-11T12:00:00Z', seed: 7 }

/** A1:A4 4, 8, 15, 16; B1:B4 2, 7, 1, 9; D1:G1 1, 2, 3, 5; formulas in column Z. */
const bookOf = (formulas: string[]): BookInput => ({
  sheets: [{ name: 'S', cells: {
    ...Object.fromEntries([4, 8, 15, 16].map((value, i) => [`A${i + 1}`, { value }])),
    ...Object.fromEntries([2, 7, 1, 9].map((value, i) => [`B${i + 1}`, { value }])),
    ...Object.fromEntries([1, 2, 3, 5].map((value, i) => [`${'DEFG'[i]}1`, { value }])),
    ...Object.fromEntries(formulas.map((formula, i) => [`Z${i + 1}`, { formula }])),
  } }],
})

describe.each(['closures', 'eval'] as const)('SORTBY by one value, %s', (execution) => {
  const values = (formulas: string[]) => {
    const cells = compute(bookOf(formulas), { clock, execution, display: false }).sheets[0].cells
    return formulas.map((_, i) => cells[`Z${i + 1}`].value)
  }

  it('one value beside a by-array of the array\'s rows, before or after it, is #VALUE!', () => {
    expect(values(['ARRAYTOTEXT(SORTBY(A1:A4,B1:B4,1,5,1),1)', 'ARRAYTOTEXT(SORTBY(A1:A4,5,1,B1:B4,1),1)', 'ARRAYTOTEXT(SORTBY(A1:A4,B1:B4,1,{5},1),1)'])).toEqual(
      new Array(3).fill({ error: 'VALUE' }),
    )
  })

  it('one value fits an array of one column or one row, alone or with another, and nothing wider', () => {
    expect(values(['ARRAYTOTEXT(SORTBY(A1:A4,5,1,7,-1),1)', 'ARRAYTOTEXT(SORTBY(D1:G1,5),1)', 'ARRAYTOTEXT(SORTBY(A1:B4,5),1)', 'ARRAYTOTEXT(SORTBY(D1:G1,5,1,D1:G1,-1),1)', 'SORTBY(A1,B1,1,5,1)'])).toEqual([
      '{4;8;15;16}', '{1,2,3,5}', { error: 'VALUE' }, { error: 'VALUE' }, 4,
    ])
  })

  it('with orders by position, the call at each position answers #VALUE!', () => {
    expect(values(['ARRAYTOTEXT(SORTBY(A1:A4,B1:B4,{1;-1},5,1),1)', 'ARRAYTOTEXT(SORTBY(A1:A4,5,{1;-1}),1)'])).toEqual(['{#VALUE!;#VALUE!}', '{4;}'])
  })
})
