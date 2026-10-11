/**
 * Date text without a year in en-US (`1/31`, `Jan 31`, `31-Jan`) is a date of the clock's year, as Excel en-US
 * completes it with the current year (the `date-noyear` set): such cells are volatile, and a session after
 * `setClock` equals a full computation with that clock, through cached criteria and criteria indexes too.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, ComputeResult } from '../src/contract.js'
import { compute, createSession } from '../src/index.js'

const at = (year: number, seed = 1) => ({ now: `${year}-10-05T12:00:00Z`, seed })

const book: BookInput = {
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: '2/29' },
        A2: { value: '1/31/2025' },
        A3: { value: 46053 },
        A4: { value: 46418 },
        B1: { formula: 'VALUE("Jan 31")' },
        B2: { formula: 'B1+1' },
        B3: { formula: 'YEAR(A1)' },
        B4: { formula: 'TEXT("31-Jan","yyyy-mm-dd")' },
        B5: { formula: 'COUNTIF(A3:A4,"1/31")' },
        B6: { formula: 'COUNTIF(A3:A4,"1/31")+0' },
        B7: { formula: 'VALUE(A2)' },
        B8: { formula: 'DAY(A1)' },
      },
    },
  ],
}

const values = (r: ComputeResult) => Object.fromEntries(Object.entries(r.sheets[0].cells).map(([a, c]) => [a, c.value]))
const volatile = (r: ComputeResult) => r.volatile.map((c) => c.cell).sort()

describe('en-US date text without a year depends on the clock', () => {
  it("takes the clock's year; February 29 of a year that is not leap is February 2029", () => {
    const v = values(compute(book, { clock: at(2026) }))
    expect([v.B1, v.B3, v.B4, v.B5, v.B6, v.B8]).toEqual([46053, 2029, '2026-01-31', 1, 1, 1])
    const w = values(compute(book, { clock: at(2028) }))
    expect([w.B1, w.B3, w.B4, w.B5, w.B6, w.B8]).toEqual([46783, 2028, '2028-01-31', 0, 0, 29])
  })

  it('such cells are volatile, with the cells over them; text with a year is not', () => {
    expect(volatile(compute(book, { clock: at(2026) }))).toEqual(['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B8'])
  })

  it('a session after setClock equals a full computation with that clock', () => {
    const s = createSession(book, { clock: at(2026) })
    for (const year of [2028, 2027, 2026]) {
      s.setClock(at(year, year))
      const full = compute(book, { clock: at(year, year) })
      expect(values(s.getResult()), String(year)).toEqual(values(full))
      expect(s.getResult().hash).toBe(full.hash)
    }
  })
})
