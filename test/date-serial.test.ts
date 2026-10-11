/**
 * `dateToSerial` and `serialToDate`: the texts `CellInput.date` takes and the serial numbers a book gives them, in both
 * date systems, each the other's inverse; Excel's phantom 1900-02-29 as the engine has it; no date before the system's
 * day 0 or after 9999-12-31; the input checks.
 */
import { describe, expect, it } from 'vitest'
import { compute, dateToSerial, InputError, serialToDate } from '../src/index.js'

describe('dateToSerial', () => {
  it('gives the serial a cell with { date } holds, in the book\'s date system', () => {
    const texts = ['2026-01-31', '2026-01-31T18:00', '2026-01-31 18:00', '2026-01-31T18:00:30', '2026-01-31T18:00:30.5', '1900-01-01', '1900-02-28', '1900-02-29', '1900-03-01', '1904-01-01', '9999-12-31']
    for (const dateSystem of [1900, 1904] as const) {
      const ok = texts.filter((t) => dateSystem === 1900 || t >= '1904')
      const cells = Object.fromEntries(ok.map((t, i) => [`A${i + 1}`, { value: { date: t } }]))
      const book = compute({ dateSystem, sheets: [{ name: 'S', cells }] }).sheets[0].cells
      // a result rounds to 15 significant digits
      expect(ok.map((t) => Number(dateToSerial(t, { dateSystem }).toPrecision(15)))).toEqual(ok.map((_, i) => book[`A${i + 1}`].value))
    }
    expect([dateToSerial('2026-01-31'), dateToSerial('2026-01-31T18:00'), dateToSerial('1900-02-29'), dateToSerial('2026-01-31', { dateSystem: 1904 })]).toEqual([46053, 46053.75, 60, 44591])
  })

  it('refuses text it cannot read and a date before the system\'s day 0, as InputError', () => {
    expect(() => dateToSerial('31.01.2026')).toThrow(new InputError('text', 'invalid date "31.01.2026": expected ISO such as "2025-01-31" or "2025-01-31T18:00:00"'))
    expect(() => dateToSerial('10000-01-01')).toThrow(/invalid date/)
    expect(() => dateToSerial('1899-12-30')).toThrow(new InputError('text', `"1899-12-30" is before the 1900 date system's day 0 (1900-01-00): Excel has no such date`))
    expect(() => dateToSerial('1903-12-31', { dateSystem: 1904 })).toThrow(/before the 1904 date system's day 0 \(1904-01-01\)/)
    expect(() => dateToSerial(new Date() as never)).toThrow(new InputError('text', 'expected a string, got an object'))
    expect(() => dateToSerial('2026-01-31', { dateSystem: 1901 as never })).toThrow(/opts\.dateSystem: expected one of 1900, 1904/)
  })
})

describe('serialToDate', () => {
  it('gives the date, and the time where there is one, as CellInput.date takes it', () => {
    expect([46053, 46053.75, 46053 + (18 * 3600 + 30.5) / 86400, 0, 0.5, 59, 60, 61, 2958465].map((s) => serialToDate(s))).toEqual([
      '2026-01-31', '2026-01-31T18:00', '2026-01-31T18:00:30.5', '1900-01-00', '1900-01-00T12:00', '1900-02-28', '1900-02-29', '1900-03-01', '9999-12-31',
    ])
    expect([serialToDate(0, { dateSystem: 1904 }), serialToDate(44591.25, { dateSystem: 1904 })]).toEqual(['1904-01-01', '2026-01-31T06:00'])
    expect(serialToDate(2958465.99999999)).toBe('9999-12-31T23:59:59.999')
  })

  it('is the inverse of dateToSerial, to the millisecond', () => {
    for (const dateSystem of [1900, 1904] as const)
      for (let i = 0; i < 2000; i++) {
        const serial = Math.round(((i * 7919) % 2_900_000) * 1000 + (i % 97) * 13.7) / 1000
        const text = serialToDate(serial, { dateSystem })
        expect(Math.abs(dateToSerial(text, { dateSystem }) - serial) * 86_400_000, `${serial} ${text}`).toBeLessThan(0.5)
      }
  })

  it('refuses what is no date: negative, past 9999-12-31, not a finite number', () => {
    const fails = (serial: unknown, dateSystem?: 1904) => expect(() => serialToDate(serial as number, dateSystem ? { dateSystem } : undefined)).toThrow(InputError)
    for (const s of [-1, -0.5, 2958466, Number.NaN, Infinity, '1', null]) fails(s)
    fails(2958465 - 1462 + 1, 1904)
    expect(() => serialToDate(-1)).toThrow(new InputError('serial', 'expected the serial number of a date, from 0 to 2958465 (9999-12-31) in the 1900 date system, got number -1'))
  })
})
