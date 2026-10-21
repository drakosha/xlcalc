/**
 * Data validation, from the review of 21.10: a rule whose own ranges overlap is one rule over each cell, not a conflict.
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
})

describe('a file with a rule whose own ranges overlap (fix-u308)', () => {
  it('reads back', async () => {
    const book = withRules([whole('A1:A5 A3:A7')])
    const r = await readXlsx((await writeXlsx({ book })).data)
    expect(r.book.sheets[0].validations).toMatchObject([{ range: 'A1:A5 A3:A7', type: 'whole' }])
  })
})
