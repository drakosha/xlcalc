/**
 * Number literals and dates typed, from the review of 21.10, as Excel answered the entry sets `review-held`, `-2` and
 * `-3`: a formula's number literal from 1E+308 on, or not 0 and below 1E-309, is refused, and one up to the smallest
 * normal number is 0 (it was Infinity, written `Infinity`, which a file round trip made a name).
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, parseEntry, validateFormula } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'

const clock = { now: '2026-10-22T12:00:00Z', seed: 7 }

describe('a number literal past the bounds of a formula (fix-u345)', () => {
  it.each(['1E400', '-1E400', '1E308', '0.1E309', 'ISNUMBER(1E400)', '1E-310', '1E-320', '0E-400', '0.0E400'])('%s is refused', (f) => {
    expect(validateFormula(f)).toMatchObject({ ok: false, code: 'syntax' })
    expect(parseEntry(`=${f}`, { now: '2026-10-22' })).toMatchObject({ ok: false, code: 'syntax' })
  })

  it('a literal up to the smallest normal number is 0, its text 0; just below 1E+308 it stays', () => {
    const typed = (t: string) => parseEntry(t, { now: '2026-10-22' })
    expect(typed('=1E-309')).toEqual({ ok: true, input: { formula: '0' } })
    expect(typed('=2.2250738585072014E-308')).toEqual({ ok: true, input: { formula: '0' } })
    expect(typed('=ISNUMBER(1E-309)')).toEqual({ ok: true, input: { formula: 'ISNUMBER(0)' } })
    expect(typed('=1E-307')).toMatchObject({ ok: true })
    expect(typed('=9.999999999999999E307')).toMatchObject({ ok: true })
    expect(typed('=0.01E309')).toMatchObject({ ok: true })
  })

  describe.each(['closures', 'eval'] as const)('execution %s', (execution) => {
    it('a book holding one computes the cell as a formula that does not read, and no Infinity reaches a file', async () => {
      const book = { sheets: [{ name: 'S', cells: { A1: { formula: '1E400' }, A2: { formula: 'ISNUMBER(1E-309)' }, A3: { formula: '1E307*100' } } }] }
      const r = compute(book, { clock, execution })
      expect(r.sheets[0].cells.A1).toMatchObject({ value: { error: 'ERROR' } })
      expect(r.diagnostics[0]).toMatchObject({ cell: 'A1', check: 'syntax' })
      expect(r.sheets[0].cells.A2).toMatchObject({ value: true, formula: 'ISNUMBER(0)' })
      expect(r.sheets[0].cells.A3).toMatchObject({ value: { error: 'NUM' } })
      const back = await readXlsx((await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A2: book.sheets[0].cells.A2 } }] } })).data)
      expect(back.book.sheets[0].cells.A2).toEqual({ formula: 'ISNUMBER(0)' })
      const s = createSession(book, { clock, execution })
      expect(s.setCellText('S', 'B1', '=1E400')).toMatchObject({ ok: false, code: 'syntax' })
    })
  })
})
