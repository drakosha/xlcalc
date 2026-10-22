/**
 * Data validation, from the review of 21.10: a rule whose own ranges overlap is one rule over each cell, not a conflict,
 * a list through a defined name sees the blanks of its cells past the source sheet's last value, and a rule's formulas
 * are written for the corner of the rectangle around its ranges (the valid set `anchor-held`).
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

describe.each(['closures', 'eval'] as const)('a rule over ranges that do not start at their corner, execution %s (fix-u310)', (execution) => {
  // A1:A6 hold 1 to 6: `$A3>2` written for H1, the corner, checks A5 at H3 and A4 at J2, as Excel reads it from a file
  const cells = Object.fromEntries([1, 2, 3, 4, 5, 6].map((r) => [`A${r}`, { value: r }]))
  const book = (range: string, formula1: string): BookInput => ({ sheets: [{ name: 'S', cells, validations: [{ range, type: 'custom', formula1 }] }] })

  it.each(['H3:H6 J1:J3', 'J1:J3 H3:H6'])('%s: the verdicts follow the corner H1', async (range) => {
    const s = createSession(book(range, '$A3>2'), { execution, clock })
    for (const [cell, ok] of [['J1', true], ['J2', true], ['H3', true]] as const) expect(s.validateEntry('S', cell, { value: 0 }).ok, cell).toBe(ok)
    const t = createSession(book(range, '$A1>2'), { execution, clock })
    for (const [cell, ok] of [['J1', false], ['J2', false], ['J3', true], ['H3', true], ['H4', true]] as const) expect(t.validateEntry('S', cell, { value: 0 }).ok, cell).toBe(ok)
    // the file keeps the formula for the corner, and reads it back so
    const back = await readXlsx((await writeXlsx({ book: book(range, '$A1>2') })).data)
    expect(back.book.sheets[0].validations).toEqual([{ range, type: 'custom', formula1: '$A1>2' }])
  })

  it('a row deleted above the corner moves it, the formula stays written for it', () => {
    const s = createSession(book('H3:H6 J2:J3', '$A2>2'), { execution, clock })
    expect(s.editStructure({ op: 'deleteRows', sheet: 'S', row: 1 }).ok).toBe(true)
    expect(s.getBook().sheets[0].validations).toEqual([{ range: 'H2:H5 J1:J2', type: 'custom', formula1: '$A1>2' }])
  })
})

describe('a file with a rule whose own ranges overlap (fix-u308)', () => {
  it('reads back', async () => {
    const book = withRules([whole('A1:A5 A3:A7')])
    const r = await readXlsx((await writeXlsx({ book })).data)
    expect(r.book.sheets[0].validations).toMatchObject([{ range: 'A1:A5 A3:A7', type: 'whole' }])
  })
})
