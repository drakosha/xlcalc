/**
 * The 1904 date system beyond the Excel cases (`test/excel/workbook/date1904.json`): the clock and a session, date
 * text in ru-RU, hashes, `displayValue`, input checks.
 */
import { describe, expect, it } from 'vitest'
import { compute, displayValue, InputError, createSession, type BookInput, type CellInput } from '../src/index.js'
import './locales.js'

const clock = { now: '2026-10-07T18:00:00Z', seed: 1 }

const cells: Record<string, CellInput> = {
  A1: { formula: 'TODAY()' },
  A2: { formula: 'NOW()' },
  A3: { value: { date: '2025-01-31' } },
  A4: { formula: 'A3+1' },
  A5: { formula: 'YEAR(DATEVALUE("1/31"))' },
  A6: { formula: 'TEXT(A3,"yyyy-mm-dd")' },
}
const book = (dateSystem?: 1900 | 1904): BookInput => ({ sheets: [{ name: 'S', cells }], ...(dateSystem ? { dateSystem } : {}) })
const values = (b: BookInput) => Object.fromEntries(Object.entries(compute(b, { clock, display: false }).sheets[0].cells).map(([k, c]) => [k, c.value]))

describe('the 1904 date system', () => {
  it('counts days from 1904-01-01: the clock, typed dates and date text', () => {
    const [a, b] = [values(book()), values(book(1904))]
    expect(b).toEqual({ A1: (a.A1 as number) - 1462, A2: (a.A2 as number) - 1462, A3: 44226, A4: 44227, A5: 2026, A6: '2025-01-31' })
  })

  it('a session with a new clock and edits gives the values of a full computation', () => {
    const s = createSession(book(1904), { clock })
    s.setClock({ now: '2027-03-01T06:00:00Z', seed: 2 })
    s.setCells([{ sheet: 'S', cell: 'B1', input: { value: { date: '1904-01-02' } } }, { sheet: 'S', cell: 'B2', input: { formula: 'B1-DATE(1904,1,1)' } }])
    const full = compute(s.getBook(), { clock: s.clock })
    expect(s.getResult().sheets).toEqual(full.sheets)
    expect(s.getResult().hash).toBe(full.hash)
    expect(full.sheets[0].cells.B2.value).toBe(1)
    s.close()
  })

  it('reads ru-RU date text in the 1904 system', () => {
    const b: BookInput = { locale: 'ru-RU', dateSystem: 1904, sheets: [{ name: 'S', cells: { A1: { formula: 'DATEVALUE("31.01.2025")' }, A2: { formula: 'TEXT(-1,"ДД.ММ.ГГГГ")' }, A3: { formula: '"31.12.1903"+0' } } }] }
    expect(Object.values(compute(b, { clock, display: false }).sheets[0].cells).map((c) => c.value)).toEqual([44226, '-02.01.1904', { error: 'VALUE' }])
  })

  it('goes into both hashes only when it is 1904', () => {
    const [x, y, z] = [compute(book(), { clock }), compute(book(1900), { clock }), compute(book(1904), { clock })]
    expect(y.hash).toBe(x.hash)
    expect(y.inputHash).toBe(x.inputHash)
    expect(z.inputHash).not.toBe(x.inputHash)
    const same: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }
    expect(compute({ ...same, dateSystem: 1904 }, { clock }).hash).not.toBe(compute(same, { clock }).hash)
  })

  it('displayValue shows a value by it', () => {
    expect(displayValue(0, { format: 'm/d/yyyy', dateSystem: 1904 })).toEqual({ text: '1/1/1904' })
    expect(displayValue(-0.5, { format: 'h:mm', dateSystem: 1904 })).toEqual({ text: '-12:00' })
    expect(displayValue(-0.5, { format: 'h:mm', width: 5 }).text).toBe('#####')
  })

  it('is 1900 or 1904', () => {
    expect(() => compute({ ...book(), dateSystem: 1905 as 1904 })).toThrow(InputError)
    expect(() => displayValue(1, { dateSystem: '1904' as unknown as 1904 })).toThrow(/opts\.dateSystem: expected one of 1900, 1904/)
  })
})
