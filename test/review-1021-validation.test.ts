/**
 * Data validation, from the review of 21.10: a rule whose own ranges overlap is one rule over each cell, not a conflict,
 * and a list through a defined name sees the blanks of its cells past the source sheet's last value.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, InputError, type BookInput, type DataValidation } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

const withRules = (validations: DataValidation[]): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 5 } }, validations }] })
const whole = (range: string): DataValidation => ({ range, type: 'whole', formula1: '1', formula2: '9' })

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  describe('a rule whose own ranges overlap (fix-u308)', () => {
    it('is accepted and checks every cell of its ranges', () => {
      for (const range of ['A1:A5 A3:A7', 'A1 A1', 'A1:A5 A5']) {
        expect(() => compute(withRules([whole(range)]), { execution, clock })).not.toThrow()
        const s = createSession(withRules([whole(range)]), { execution, clock })
        expect(s.validateEntry('S', 'A1', { value: 3 })).toEqual({ ok: true })
        expect(s.validateEntry('S', 'A1', { value: 30 }).ok).toBe(false)
      }
      const s = createSession(withRules([whole('A1:A5 A3:A7')]), { execution, clock })
      expect(s.validateEntry('S', 'A7', { value: 30 }).ok).toBe(false)
    })

    it('two rules over one cell are still refused', () => {
      expect(() => createSession(withRules([whole('A1:A5'), whole('A3:A7')]), { execution, clock })).toThrow(InputError)
      expect(() => createSession(withRules([whole('A1:A5 C1'), whole('B1 A3')]), { execution, clock })).toThrow(/A1:A5 C1/)
    })
  })

  describe('a list through a defined name with blank cells past the source sheet last value (fix-u307)', () => {
    const named = (ref: string, more: BookInput['sheets'][number]['cells'] = {}): BookInput => ({
      names: [{ name: 'Items', ref }],
      sheets: [
        { name: 'S', cells: {}, validations: [{ range: 'A1', type: 'list', formula1: 'Items' }] },
        { name: 'L', cells: { A1: { value: 'a' }, A2: { value: 'b' }, A3: { value: 'c' }, ...more } },
      ],
    })
    const entry = (book: BookInput) => createSession(book, { execution, clock }).validateEntry('S', 'A1', { value: 'Foo' }).ok

    it('accepts anything, whatever other columns of the sheet hold', () => {
      for (const ref of ["'L'!$A$1:$A$100", "'L'!$A:$A"]) {
        expect(entry(named(ref)), ref).toBe(true)
        expect(entry(named(ref, { B50: { value: 1 } })), ref).toBe(true)
        expect(entry(named(ref, { B3: { value: 1 } })), ref).toBe(true)
      }
    })

    it('a name without a blank, or blanks not ignored, still refuses', () => {
      expect(entry(named("'L'!$A$1:$A$3", { B50: { value: 1 } }))).toBe(false)
      const strict = named("'L'!$A$1:$A$100")
      strict.sheets[0].validations![0].allowBlank = false
      expect(entry(strict)).toBe(false)
      const s = createSession(named("'L'!$A$1:$A$100"), { execution, clock })
      expect(s.validateEntry('S', 'A1', { value: 'b' })).toEqual({ ok: true })
      // the dropdown leaves out the blanks past the last value
      expect(s.getValidation('S', 'A1')!.list).toHaveLength(3)
    })
  })
})

describe('a file with a rule whose own ranges overlap (fix-u308)', () => {
  it('reads back', async () => {
    const book = withRules([whole('A1:A5 A3:A7')])
    const r = await readXlsx((await writeXlsx({ book })).data)
    expect(r.book.sheets[0].validations).toMatchObject([{ range: 'A1:A5 A3:A7', type: 'whole' }])
  })
})
