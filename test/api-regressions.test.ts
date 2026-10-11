/**
 * Reproductions from the review of the API by a first-time user and by the two hosts, in both execution modes: each
 * failed before its fix.
 */
import { describe, expect, it } from 'vitest'
import { compute, displayValue, InputError, createSession, parseEntry, type BookInput } from '../src/index.js'
import { writeXlsx } from '../src/xlsx.js'

const clock = { now: '2026-10-10T12:00:00Z', seed: 7 }

describe.each(['closures', 'eval'] as const)('API regressions, %s', (execution) => {
  it('a fill shows nothing where there is no width: displayValue without one, a number typed into a header cell', () => {
    expect(displayValue(1234.5, { numberFormat: 44 })).toEqual({ text: ' $1,234.50 ' })
    expect(displayValue(-5, { numberFormat: '0*x;-0*-' })).toEqual({ text: '-5' })
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 'Qty', numberFormat: 44 }, A2: { value: 1 } }, tables: [{ name: 'T', range: 'A1:A2', columns: [{ name: 'Qty' }] }] }] }
    const s = createSession(book, { clock, execution })
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 7, numberFormat: 44 } }])
    expect(s.getBook().sheets[0].tables![0].columns[0].name).toBe(' $7.00 ')
    s.close()
  })

  it("a date typed in as 1900-02-29 is Excel's phantom day, serial 60, as Excel takes it typed (entry set typed-held, more-18)", () => {
    const cells = { A1: { value: { date: '1900-02-29' } }, A2: { value: { date: '1900-02-29T12:00' } }, B1: { formula: 'DAY(A1)&"."&MONTH(A1)' }, B2: { formula: 'A2-A1' } }
    const r = compute({ sheets: [{ name: 'S', cells }] }, { clock, execution }).sheets[0].cells
    expect([r.A1.value, r.A2.value, r.B1.value, r.B2.value]).toEqual([60, 60.5, '29.2', 0.5])
    // as the same text typed into a cell, and the days around it as they were
    expect(parseEntry('1900-02-29')).toEqual({ ok: true, input: { value: 60, numberFormat: 14 } })
    const around = compute({ sheets: [{ name: 'S', cells: { A1: { value: { date: '1900-02-28' } }, A2: { value: { date: '1900-03-01' } } } }] }, { clock, execution }).sheets[0].cells
    expect([around.A1.value, around.A2.value]).toEqual([59, 61])
  })

  it('writeXlsx given a session where a result goes refuses it with what to pass, not a TypeError', async () => {
    const s = createSession({ sheets: [{ name: 'S', cells: { A1: { formula: '1+1' } } }] }, { clock, execution })
    await expect(writeXlsx(s.getBook(), { result: s as never })).rejects.toThrow(InputError)
    await expect(writeXlsx(s.getBook(), { result: s as never })).rejects.toThrow(/pass session\.getSnapshot\(\)/)
    s.close()
  })
})
