/**
 * Date text without a year (`1.5` in ru-RU is 1 May of the clock's year) makes a cell depend on the clock as
 * TODAY does: such cells are volatile, and a session after `setClock` equals a full computation with that clock,
 * also where the year is read through cached criteria and criteria indexes. A book whose text names a year, and
 * an en-US book, depend on no clock.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, ComputeResult } from '../src/contract.js'
import { compute, openSession } from '../src/index.js'
import './locales.js'

const at = (year: number, seed = 1) => ({ now: `${year}-10-05T12:00:00Z`, seed })

const book: BookInput = {
  locale: 'ru-RU',
  sheets: [
    {
      name: 'S',
      cells: {
        // text dates without a year and with one, numbers to count
        A1: { value: '1.5' },
        A2: { value: '31.12.2025' },
        A3: { value: 46143 },
        A4: { value: 46508 },
        B1: { formula: 'VALUE("1.5")' },
        B2: { formula: 'B1+1' },
        B3: { formula: 'YEAR(A1)' },
        B4: { formula: 'TEXT("1.5","ДД.ММ.ГГГГ")' },
        B5: { formula: 'COUNTIF(A3:A4,"01.05")' },
        // the same criterion twice: the second must not take the first one's cached reading
        B6: { formula: 'COUNTIF(A3:A4,"01.05")+0' },
        // an index over text dates without a year, read by two formulas
        B7: { formula: 'COUNTIF(A1:A2,46143)' },
        B8: { formula: 'COUNTIF(A1:A2,46508)' },
        B9: { formula: 'VALUE(A2)' },
        B10: { formula: 'DATE(2025,1,1)' },
      },
    },
  ],
}

const values = (r: ComputeResult) => Object.fromEntries(Object.entries(r.sheets[0].cells).map(([a, c]) => [a, c.value]))
const volatile = (r: ComputeResult) => r.volatile.map((c) => c.cell).sort()

describe('date text without a year depends on the clock', () => {
  it("takes the clock's year", () => {
    const v = values(compute(book, { clock: at(2026) }))
    expect([v.B1, v.B3, v.B4, v.B5, v.B6, v.B7, v.B8]).toEqual([46143, 2026, '01.05.2026', 1, 1, 1, 0])
    const w = values(compute(book, { clock: at(2027) }))
    expect([w.B1, w.B3, w.B4, w.B5, w.B6, w.B7, w.B8]).toEqual([46508, 2027, '01.05.2027', 1, 1, 0, 1])
  })

  it('such cells are volatile, with the cells over them; text with a year is not', () => {
    expect(volatile(compute(book, { clock: at(2026) }))).toEqual(['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8'])
    expect(compute({ ...book, locale: 'en-US' }, { clock: at(2026) }).volatile).toEqual([])
  })

  it('a session after setClock equals a full computation with that clock', () => {
    const s = openSession(book, { clock: at(2026) })
    for (const year of [2027, 2028, 2026]) {
      s.setClock(at(year, year))
      const full = compute(book, { clock: at(year, year) })
      const r = s.result()
      expect(values(r), String(year)).toEqual(values(full))
      expect(r.hash).toBe(full.hash)
      expect(volatile(r)).toEqual(volatile(full))
    }
  })

  it('setClock reports the cells whose value the year changed', () => {
    const s = openSession(book, { clock: at(2026) })
    expect(s.setClock(at(2027)).map((c) => c.cell).sort()).toEqual(['B1', 'B2', 'B3', 'B4', 'B7', 'B8'])
  })

  it('an edit that removes the text without a year leaves the clock', () => {
    const s = openSession(book, { clock: at(2026) })
    s.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'VALUE("1.5.2026")' } }])
    expect(volatile(s.result())).not.toContain('B1')
    expect(volatile(s.result())).not.toContain('B2')
    s.setClock(at(2027))
    expect(values(s.result())).toEqual(values(compute(s.book(), { clock: at(2027) })))
  })
})
