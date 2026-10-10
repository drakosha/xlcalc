/**
 * Reproductions from the review of the API by a first-time user and by the two hosts, in both execution modes: each
 * failed before its fix.
 */
import { describe, expect, it } from 'vitest'
import { displayValue, InputError, openSession, type BookInput } from '../src/index.js'
import { writeXlsx } from '../src/xlsx.js'

const clock = { now: '2026-10-10T12:00:00Z', seed: 7 }

describe.each(['closures', 'eval'] as const)('API regressions, %s', (execution) => {
  it('a fill shows nothing where there is no width: displayValue without one, a number typed into a header cell', () => {
    expect(displayValue(1234.5, { format: 44 })).toEqual({ text: ' $1,234.50 ' })
    expect(displayValue(-5, { format: '0*x;-0*-' })).toEqual({ text: '-5' })
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 'Qty', format: 44 }, A2: { value: 1 } }, tables: [{ name: 'T', range: 'A1:A2', columns: [{ name: 'Qty' }] }] }] }
    const s = openSession(book, { clock, execution })
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 7, format: 44 } }])
    expect(s.book().sheets[0].tables![0].columns[0].name).toBe(' $7.00 ')
    s.close()
  })

  it('writeXlsx given a session where a result goes refuses it with what to pass, not a TypeError', async () => {
    const s = openSession({ sheets: [{ name: 'S', cells: { A1: { formula: '1+1' } } }] }, { clock, execution })
    await expect(writeXlsx(s.book(), { result: s as never })).rejects.toThrow(InputError)
    await expect(writeXlsx(s.book(), { result: s as never })).rejects.toThrow(/pass session\.getSnapshot\(\)/)
    s.close()
  })
})
