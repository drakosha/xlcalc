/**
 * Data validation in the model: checked as input, taking no part in computation or the values hash, in `inputHash` only
 * when a book has rules, moved by structure operations (the rules measured in `test/excel/valid/`), changed through a
 * session without recomputing, a cell's rule and list given by `Session.getValidation` and values checked by
 * `Session.validateEntry`, in both execution modes.
 */
import { describe, expect, it } from 'vitest'
import { compute, createInlineCalc, InputError, createSession, editStructure, type BookInput, type CellInput, type DataValidation, type Execution, type SheetInput, type StructureOp } from '../src/index.js'
import { pool } from './fixtures/pool.js'
import './locales.js'
import { sameValuesApart } from './formula-values.js'

const clock = { now: '2026-10-10T12:00:00Z', seed: 1 }

/** S: values, a sum, rules of several types; L: a list's items; the name Fruits over them. */
const book = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, B1: { formula: 'SUM(A1:A3)' }, H1: { value: 10 }, C2: { value: 'Yes' }, D2: { value: 4 } },
      validations: [
        { range: 'C2:C9', type: 'list', formula1: '"Yes,No"', prompt: 'Pick one' },
        { range: 'D2:D9', type: 'whole', operator: 'lessThan', formula1: '$H$1' },
        { range: 'E2:E9 G2', type: 'custom', formula1: 'E2>A2' },
        { range: 'F2:F9', type: 'list', formula1: 'Fruits' },
      ],
    },
    { name: 'L', cells: { A1: { value: 'apple' }, A2: { value: 1.5, numberFormat: '0.00' }, A3: { value: 45658, numberFormat: 'dd.mm.yyyy' } } },
  ],
  names: [{ name: 'Fruits', ref: "'L'!$A$1:$A$3" }],
})

const bare = (b: BookInput): BookInput => ({ ...b, sheets: b.sheets.map(({ validations: _, ...s }): SheetInput => s) })

const rules = (b: BookInput, sheet = 'S') => (b.sheets.find((s) => s.name === sheet)!.validations ?? []).map((v) => `${v.range} ${v.formula1}`)

const after = (op: StructureOp, b = book()) => {
  const r = editStructure(b, op)
  if (!r.ok) throw new Error(r.message)
  return r.book
}

describe('data validation as input', () => {
  const fails = (validations: unknown, message: string) => expect(() => compute({ sheets: [{ name: 'S', cells: {}, validations } as SheetInput] })).toThrow(message)
  it('wrong shapes throw InputError with the path', () => {
    expect(() => compute({ sheets: [{ name: 'S', cells: {}, validations: {} } as never] })).toThrow(InputError)
    fails({}, 'book.sheets[0].validations: expected an array of rules')
    fails([{ range: 'B2', type: 'list' }], 'validations[0].formula1: a rule of type list needs formula1')
    fails([{ range: 'B0', type: 'any' }], 'validations[0].range: expected ranges of the sheet')
    fails([{ range: 'B2', type: 'whole', formula1: '1' }], 'validations[0].formula2: between and notBetween need formula2')
    fails([{ range: 'B2', type: 'whole', operator: 'equal', formula1: '1', formula2: '2' }], 'only between and notBetween take formula2')
    fails([{ range: 'B2', type: 'list', operator: 'equal', formula1: '"a"' }], 'a rule of type list has no operator')
    fails([{ range: 'B2', type: 'any', formula1: '1' }], 'a rule of type any has no formula')
    fails([{ range: 'B2', type: 'bad' }], 'validations[0].type')
    fails([{ range: 'B2', type: 'any', errorStyle: 'fatal' }], 'validations[0].errorStyle')
    fails([{ range: 'B2', type: 'any', showError: 'yes' }], 'validations[0].showError')
    fails([{ range: 'B2', type: 'any', hint: 'x' }], 'unknown field "hint"')
    fails([{ range: 'B2:B9', type: 'any' }, { range: 'A1 B5', type: 'any' }], 'validations[1].range: A1 B5 meets B2:B9: a cell has one rule at most')
    expect(() => compute({ sheets: [{ name: 'S', cells: {}, validations: [{ range: 'B2:B9 D:D', type: 'any' }, { range: 'F4', type: 'any' }] }] })).not.toThrow()
    expect(() => compute({ sheets: [{ name: 'S', cells: {}, validations: [{ range: '4:5', type: 'any' }] }] })).not.toThrow()
  })

  it('values and the values hash are those of the book without rules; inputHash takes them in only when there are any', () => {
    const r = sameValuesApart(book(), bare(book()), clock)
    // rules in any order hash alike
    const swapped = book()
    swapped.sheets[0].validations!.reverse()
    expect(compute(swapped, { clock }).inputHash).toBe(r.inputHash)
    expect(compute(bare(book()), { clock }).inputHash).toBe(compute({ ...bare(book()), sheets: bare(book()).sheets.map((s) => ({ ...s })) }, { clock }).inputHash)
  })
})

describe('data validation under structure operations', () => {
  it('ranges grow with lines inserted inside and right after them, move with lines inserted before, and formulas follow', () => {
    expect(rules(after({ op: 'insertRows', sheet: 'S', row: 1 }))).toEqual(['C3:C10 "Yes,No"', 'D3:D10 $H$2', 'E3:E10 G3 E3>A3', 'F3:F10 Fruits'])
    expect(rules(after({ op: 'insertRows', sheet: 'S', row: 10, count: 2 }))).toEqual(['C2:C11 "Yes,No"', 'D2:D11 $H$1', 'E2:E11 G2 E2>A2', 'F2:F11 Fruits'])
    expect(rules(after({ op: 'insertRows', sheet: 'S', row: 3 }))).toEqual(['C2:C10 "Yes,No"', 'D2:D10 $H$1', 'E2:E10 G2:G3 E2>A2', 'F2:F10 Fruits'])
    expect(rules(after({ op: 'insertColumns', sheet: 'S', column: 'F' }))).toEqual(['C2:C9 "Yes,No"', 'D2:D9 $I$1', 'E2:F9 H2 E2>A2', 'G2:G9 Fruits'])
  })

  it('deleted lines leave ranges, a rule losing its first cell has its formulas written for the next', () => {
    expect(rules(after({ op: 'deleteRows', sheet: 'S', row: 2 }))).toEqual(['C2:C8 "Yes,No"', 'D2:D8 $H$1', 'E2:E8 E2>A2', 'F2:F8 Fruits'])
    expect(rules(after({ op: 'deleteColumns', sheet: 'S', column: 'E' }))).toEqual(['C2:C9 "Yes,No"', 'D2:D9 $G$1', 'F2 F2>C2', 'E2:E9 Fruits'])
    expect(rules(after({ op: 'deleteRows', sheet: 'S', row: 2, count: 8 }))).toEqual([])
    expect(after({ op: 'deleteRows', sheet: 'S', row: 2, count: 8 }).sheets[0].validations).toBeUndefined()
  })

  it("a rule of one cell on the sheet's last row or column keeps ending there when lines before it are deleted, as a longer one does (the set edge-held)", () => {
    const edge: BookInput = { sheets: [{ name: 'S', cells: {}, validations: [{ range: 'C1048576', type: 'any', prompt: 'r' }, { range: 'XFD3', type: 'any', prompt: 'c' }] }] }
    expect(rules(after({ op: 'deleteRows', sheet: 'S', row: 5, count: 2 }, edge))).toEqual(['C1048574:C1048576 undefined', 'XFD3 undefined'])
    expect(rules(after({ op: 'deleteColumns', sheet: 'S', column: 'B' }, edge))).toEqual(['B1048576 undefined', 'XFC3:XFD3 undefined'])
  })

  it("a move takes rules along and keeps the target's where the moved cells have none; a copy pastes them, a plain copy clears them", () => {
    expect(rules(after({ op: 'moveRange', from: { sheet: 'S', range: 'E2:E3' }, to: { sheet: 'S', cell: 'J5' } }))).toEqual(['C2:C9 "Yes,No"', 'D2:D9 $H$1', 'E4:E9 G2 E4>A4', 'F2:F9 Fruits', 'J5:J6 J5>A2'])
    expect(rules(after({ op: 'moveRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', cell: 'C4' } }))).toEqual(['C2:C9 "Yes,No"', 'D2:D9 $H$1', 'E2:E9 G2 E2>A2', 'F2:F9 Fruits'])
    expect(rules(after({ op: 'copyRange', from: { sheet: 'S', range: 'E2' }, to: { sheet: 'S', range: 'J2:J4' } }))).toEqual(['C2:C9 "Yes,No"', 'D2:D9 $H$1', 'E2:E9 G2 J2:J4 E2>A2', 'F2:F9 Fruits'])
    expect(rules(after({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'C3' } }))).toEqual(['C2 C4:C9 "Yes,No"', 'D2:D9 $H$1', 'E2:E9 G2 E2>A2', 'F2:F9 Fruits'])
    // to another sheet a rule's formulas move as a copy moves them
    expect(rules(after({ op: 'moveRange', from: { sheet: 'S', range: 'E2:E3' }, to: { sheet: 'L', cell: 'C5' } }), 'L')).toEqual(['C5:C6 C5>#REF!'])
  })

  it("renamed and deleted sheets in the rules' formulas", () => {
    expect(rules(after({ op: 'renameSheet', sheet: 'L', name: 'Lists' })).at(-1)).toBe('F2:F9 Fruits')
    const b = book()
    b.sheets[0].validations![3] = { range: 'F2:F9', type: 'list', formula1: "'L'!$A$1:$A$3" }
    expect(rules(after({ op: 'renameSheet', sheet: 'L', name: 'Lists' }, b)).at(-1)).toBe("F2:F9 'Lists'!$A$1:$A$3")
    expect(rules(after({ op: 'deleteSheet', sheet: 'L' }, b)).at(-1)).toBe('F2:F9 #REF!$A$1:$A$3')
  })
})

describe.each(['closures', 'eval'] as Execution[])('data validation in a session (%s)', (execution) => {
  it("validation gives a cell's rule, its formulas as at the cell, a list's items with their shown text", () => {
    const s = createSession(book(), { clock, execution })
    expect(s.getValidation('S', 'A1')).toBeNull()
    expect(s.getValidation('S', 'C5')).toEqual({ rule: book().sheets[0].validations![0], formula1: '"Yes,No"', list: [{ value: 'Yes', text: 'Yes', input: { value: 'Yes' } }, { value: 'No', text: 'No', input: { value: 'No' } }] })
    expect(s.getValidation('S', 'G2')).toEqual({ rule: book().sheets[0].validations![2], formula1: 'G2>C2' })
    expect(s.getValidation('S', 'E5')!.formula1).toBe('E5>A5')
    expect(s.getValidation('S', 'F3')!.list).toEqual([{ value: 'apple', text: 'apple', input: { value: 'apple' } }, { value: 1.5, text: '1.50', input: { value: 1.5 } }, { value: 45658, text: '01.01.2025', input: { value: 45658 } }])
    // the list follows the book
    s.setCells([{ sheet: 'L', cell: 'A2', input: { value: 'pear' } }])
    expect(s.getValidation('S', 'F3')!.list).toEqual([{ value: 'apple', text: 'apple', input: { value: 'apple' } }, { value: 'pear', text: 'pear', input: { value: 'pear' } }, { value: 45658, text: '01.01.2025', input: { value: 45658 } }])
    s.setNames([{ name: 'Fruits', ref: '#REF!' }])
    expect(s.getValidation('S', 'F3')!.list).toEqual({ error: 'REF' })
  })

  it("a picked item's input is what Excel writes: a cell's value as it is, a literal item read as typed in en-US, whatever the book's locale", () => {
    const s = createSession({
      locale: 'ru-RU',
      sheets: [{ name: 'S', cells: { E1: { value: '001' }, E2: { value: 0.5, numberFormat: '0%' }, E3: { value: 'ИСТИНА' } }, validations: [{ range: 'B1:B9', type: 'list', formula1: '$E$1:$E$3' }, { range: 'C1:C9', type: 'list', formula1: '"001,1/2,x"' }] }],
    }, { clock, execution })
    expect((s.getValidation('S', 'B2')!.list as { input: unknown }[]).map((i) => i.input)).toEqual([{ value: '001' }, { value: 0.5 }, { value: 'ИСТИНА' }])
    expect((s.getValidation('S', 'C2')!.list as { input: unknown }[]).map((i) => i.input)).toEqual([{ value: 1 }, { value: 46024, numberFormat: 16 }, { value: 'x' }])
    s.setCells([{ sheet: 'S', cell: 'B2', input: { numberFormat: '0%' } }])
    s.setCells([{ sheet: 'S', cell: 'B2', input: (s.getValidation('S', 'B2')!.list as { input: CellInput }[])[0].input }])
    expect(s.getBook().sheets[0].cells.B2).toEqual({ value: '001', numberFormat: '0%' })
  })

  it('validateEntry checks a value as Excel does, with the value put in and taken back', () => {
    const s = createSession(book(), { clock, execution })
    const hash = s.getHash()
    const ok = (cell: string, input: Parameters<typeof s.validateEntry>[2]) => s.validateEntry('S', cell, input).ok
    expect(ok('A9', { value: 'anything' })).toBe(true)
    expect([ok('C3', { value: 'No' }), ok('C3', { value: 'no' }), ok('C3', { value: ' No ' }), ok('C3', null)]).toEqual([true, false, true, true])
    expect([ok('D3', { value: 9 }), ok('D3', { value: 10 }), ok('D3', { value: 9.5 }), ok('D3', { value: '5' }), ok('D3', { formula: 'A1*3' })]).toEqual([true, false, false, false, true])
    expect([ok('E2', { value: 3 }), ok('E2', { value: 1 }), ok('E3', { formula: 'A3' })]).toEqual([true, false, false])
    expect([ok('F2', { value: 'APPLE' }), ok('F2', { value: 1.5 }), ok('F2', { value: 'kiwi' })]).toEqual([true, true, false])
    const r = s.validateEntry('S', 'D2', { value: 11 })
    expect(r).toEqual({ ok: false, code: 'validation', message: "This value doesn't match the data validation restrictions defined for this cell.", rule: book().sheets[0].validations![1], severity: 'stop' })
    // nothing changed
    expect(s.getHash()).toBe(hash)
    expect(s.getBook().sheets[0].cells.D2).toEqual({ value: 4 })
    expect(() => s.validateEntry('S', 'D2', { value: [] } as never)).toThrow(InputError)
    expect(() => s.validateEntry('X', 'D2', null)).toThrow(InputError)
  })

  it('blanks: ignored by default, checked as a blank where they are not; a formula referring to a blank cell lets anything in', () => {
    const v = (rule: Omit<DataValidation, 'range'>) => createSession({ sheets: [{ name: 'S', cells: { Z1: { value: null } }, validations: [{ range: 'A1', ...rule }] }] }, { clock, execution })
    expect(v({ type: 'whole', operator: 'greaterThan', formula1: '0' }).validateEntry('S', 'A1', null).ok).toBe(true)
    expect(v({ type: 'whole', operator: 'greaterThan', formula1: '0', allowBlank: false }).validateEntry('S', 'A1', null).ok).toBe(false)
    expect(v({ type: 'whole', formula1: '-1', formula2: '1', allowBlank: false }).validateEntry('S', 'A1', null).ok).toBe(true)
    expect(v({ type: 'textLength', operator: 'lessThan', formula1: '1', allowBlank: false }).validateEntry('S', 'A1', null).ok).toBe(true)
    expect(v({ type: 'custom', formula1: 'A1>$Z$1+5' }).validateEntry('S', 'A1', { value: 1 }).ok).toBe(true)
    expect(v({ type: 'custom', formula1: 'A1>$Z$1+5', allowBlank: false }).validateEntry('S', 'A1', { value: 1 }).ok).toBe(false)
    expect(v({ type: 'decimal', operator: 'equal', formula1: '0.3' }).validateEntry('S', 'A1', { formula: '0.1+0.2' }).ok).toBe(true)
    expect(v({ type: 'whole', formula1: '1', formula2: '10' }).validateEntry('S', 'A1', { value: true }).ok).toBe(false)
  })

  it('rules change through setLayout and recompute nothing; structure operations move them in the session', () => {
    const s = createSession(book(), { clock, execution })
    expect(s.setLayout([{ sheet: 'S', validations: [{ range: 'A1:A3', type: 'whole', operator: 'greaterThan', formula1: '1' }] }]).changes).toEqual([])
    expect(s.getBook().sheets[0].validations).toEqual([{ range: 'A1:A3', type: 'whole', operator: 'greaterThan', formula1: '1' }])
    expect(s.validateEntry('S', 'A1', { value: 1 }).ok).toBe(false)
    expect(() => s.setLayout([{ sheet: 'S', validations: [{ range: 'A1', type: 'list' } as never] }])).toThrow(InputError)
    s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 })
    expect(s.getBook().sheets[0].validations![0].range).toBe('A2:A4')
    s.setLayout([{ sheet: 'S', validations: [] }])
    expect(s.getBook().sheets[0].validations).toBeUndefined()
    expect(compute(s.getBook(), { clock, execution }).hash).toBe(s.getHash())
  })

  it('an inline session and a session in a thread give the same', async () => {
    const threads = await pool({ threads: 1 })
    const t = await threads.createSession(book(), { clock, execution })
    expect((await t.getValidation('S', 'C2'))!.list).toEqual([{ value: 'Yes', text: 'Yes', input: { value: 'Yes' } }, { value: 'No', text: 'No', input: { value: 'No' } }])
    expect(await t.validateEntry('S', 'D2', { value: 9 })).toEqual({ ok: true })
    await threads.close()
    const calc = createInlineCalc()
    const s = await calc.createSession(book(), { clock, execution })
    expect((await s.getValidation('S', 'F3'))!.list).toEqual([{ value: 'apple', text: 'apple', input: { value: 'apple' } }, { value: 1.5, text: '1.50', input: { value: 1.5 } }, { value: 45658, text: '01.01.2025', input: { value: 45658 } }])
    expect(await s.validateEntry('S', 'D2', { value: 11 })).toEqual({ ok: false, code: 'validation', message: "This value doesn't match the data validation restrictions defined for this cell.", rule: book().sheets[0].validations![1], severity: 'stop' })
    await calc.close()
  })
})
