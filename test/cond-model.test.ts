/**
 * Conditional formats in the model: checked as input, in the input hash, kept by a session's book, replaced through
 * setLayout, outside the values hash; Session.getCells gives each cell its look, in both execution modes.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, ConditionalFormat } from '../src/contract.js'
import { compute, InputError, createSession } from '../src/index.js'
import { inputHash } from '../src/canon.js'

const CLOCK = { now: '2026-10-10T12:00:00Z', seed: 1 }
const RED = { fill: { color: '#FFC7CE' } }

function book(rules: ConditionalFormat[]): BookInput {
  const cells: BookInput['sheets'][number]['cells'] = {}
  for (let r = 1; r <= 5; r++) cells[`A${r}`] = { value: r }
  return { sheets: [{ name: 'S', cells, conditionalFormats: rules }] }
}

const GT2: ConditionalFormat = { type: 'cellIs', ranges: ['A1:A5'], operator: 'greaterThan', formulas: ['2'], style: RED }

const bad = (rules: unknown[], message: RegExp) => expect(() => compute(book(rules as ConditionalFormat[]))).toThrow(message)

describe('conditional formats as input', () => {
  it('a book with every kind of rule is taken', () => {
    const rules: ConditionalFormat[] = [
      GT2,
      { type: 'expression', ranges: ['B:B', '3:4'], formula: 'ISODD(ROW())', stopIfTrue: true, style: { font: { bold: false, underline: 'double' }, border: { top: { style: 'thin' } }, numberFormat: 14 } },
      { type: 'containsText', ranges: ['C1'], text: 'x' },
      { type: 'containsBlanks', ranges: ['C1'], formula: 'LEN(C1)=0' },
      { type: 'timePeriod', ranges: ['C1'], period: 'nextWeek' },
      { type: 'top10', ranges: ['A1:A5'], rank: 100, percent: true, bottom: true },
      { type: 'aboveAverage', ranges: ['A1:A5'], below: true, stdDev: 3 },
      { type: 'duplicateValues', ranges: ['A1:A5'] },
      { type: 'colorScale', ranges: ['A1:A5'], points: [{ value: { type: 'num', value: '1' }, color: '#000000' }, { value: { type: 'formula', value: '$A$5' }, color: '#FFFFFF' }] },
      { type: 'dataBar', ranges: ['A1:A5'], min: { type: 'autoMin' }, max: { type: 'percentile', value: '90' }, color: '#638EC6', minLength: 0, maxLength: 100, axis: 'middle', direction: 'rightToLeft', excel2007: false },
      { type: 'iconSet', ranges: ['A1:A5'], icons: '4Rating', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '25' }, gte: false }, { value: { type: 'num', value: '3' }, icon: null }, { value: { type: 'num', value: '4' }, icon: { set: '5Boxes', index: 4 } }], reverse: true, showValue: false },
    ]
    expect(compute(book(rules), { clock: CLOCK }).sheets[0].cells.A1.value).toBe(1)
  })

  it('wrong rules throw InputError with the path', () => {
    bad([{ type: 'cellIs', ranges: ['A1'], operator: 'greaterThan', formulas: ['1', '2'] }], /conditionalFormats\[0\]\.formulas: greaterThan takes 1 formula, got 2/)
    bad([{ type: 'cellIs', ranges: [], operator: 'equal', formulas: ['1'] }], /ranges: a conditional format applies to one range at least/)
    bad([{ type: 'cellIs', ranges: ['a1'], operator: 'equal', formulas: ['1'] }], /ranges\[0\]: expected a range/)
    bad([{ type: 'colorScale', ranges: ['A1'], points: [{ value: { type: 'min' }, color: '#000000' }] }], /a colour scale has two or three points, got 1/)
    bad([{ type: 'iconSet', ranges: ['A1'], icons: '3Arrows', points: [{ value: { type: 'min' } }] }], /the set 3Arrows has 3 icons, so 3 thresholds; got 1/)
    bad([{ type: 'dataBar', ranges: ['A1'], min: { type: 'num' }, max: { type: 'max' }, color: '#000000' }], /min\.value: expected a string/)
    bad([{ type: 'dataBar', ranges: ['A1'], min: { type: 'min', value: '1' }, max: { type: 'max' }, color: '#000000' }], /a threshold of type min has no value/)
    bad([{ type: 'dataBar', ranges: ['A1'], min: { type: 'min' }, max: { type: 'max' }, color: '#000000', minLength: 60, maxLength: 40 }], /the shortest bar \(60\) is longer than the longest \(40\)/)
    bad([{ type: 'aboveAverage', ranges: ['A1'], equal: true, stdDev: 1 }], /a rule over standard deviations has no equal/)
    bad([{ type: 'top10', ranges: ['A1'], rank: 101, percent: true }], /rank: expected an integer from 1 to 100/)
    bad([{ type: 'expression', ranges: ['A1'], formula: 'A1', style: { alignment: {} } }], /unknown field "alignment"/)
    bad([{ type: 'expression', ranges: ['A1'], formula: 'A1', operator: 'equal' }], /unknown field "operator"/)
    bad([{ type: 'nope', ranges: ['A1'] }], /type: expected one of/)
  })

  it('the input hash counts the rules, the values hash does not', () => {
    const [a, b] = [compute(book([]), { clock: CLOCK }), compute(book([GT2]), { clock: CLOCK })]
    expect(a.hash).toBe(b.hash)
    expect(a.inputHash).not.toBe(b.inputHash)
    expect(inputHash(book([]))).toBe(inputHash({ sheets: [{ name: 'S', cells: book([]).sheets[0].cells }] }))
  })
})

describe.each(['closures', 'eval'] as const)('conditional formats in a session (%s)', (execution) => {
  it('Session.getCells gives the look, an empty cell with one too; result() gives none', () => {
    const s = createSession(book([GT2, { type: 'expression', ranges: ['B1:B2'], formula: '$A1=1', style: { font: { italic: true } } }]), { clock: CLOCK, execution })
    const w = s.getCells('S', 'A1:B5')
    expect(w.A1).toEqual({ value: 1, text: '1' })
    expect(w.A3).toEqual({ value: 3, text: '3', appearance: { style: RED } })
    expect(w.B1).toEqual({ value: null, text: '', appearance: { style: { font: { italic: true } } } })
    expect(w.B2).toBeUndefined()
    expect(s.getResult().sheets[0].cells.A3).toEqual({ value: 3, text: '3' })
  })

  it('setLayout replaces the rules; the book keeps them; an empty list removes them', () => {
    const s = createSession(book([GT2]), { clock: CLOCK, execution })
    expect(s.setLayout([{ sheet: 'S', conditionalFormats: [{ ...GT2, formulas: ['4'] }] }]).changes).toEqual([])
    expect(s.getCells('S', 'A3').A3.appearance).toBeUndefined()
    expect(s.getCells('S', 'A5').A5.appearance).toEqual({ style: RED })
    expect(s.getBook().sheets[0].conditionalFormats).toEqual([{ ...GT2, formulas: ['4'] }])
    s.setLayout([{ sheet: 'S', conditionalFormats: [] }])
    expect(s.getBook().sheets[0].conditionalFormats).toBeUndefined()
    expect(() => s.setLayout([{ sheet: 'S', conditionalFormats: [{ type: 'top10', ranges: ['A1'], rank: 0 }] } as never])).toThrow(InputError)
  })

  it('an edit changes the looks the next read gives: statistics over the edited cells are taken again', () => {
    const scale: ConditionalFormat = { type: 'colorScale', ranges: ['A1:A5'], points: [{ value: { type: 'min' }, color: '#000000' }, { value: { type: 'max' }, color: '#FF0000' }] }
    const s = createSession(book([scale, { type: 'expression', ranges: ['B1:B5'], formula: '$A1>3', style: RED }]), { clock: CLOCK, execution })
    expect(s.getCells('S', 'A5').A5.appearance).toEqual({ style: { fill: { color: '#FF0000' } } })
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 9 } }])
    const w = s.getCells('S', 'A1:B5')
    expect(w.A1.appearance).toEqual({ style: { fill: { color: '#FF0000' } } })
    expect(w.A5.appearance).toEqual({ style: { fill: { color: '#6D0000' } } })
    expect(w.B1.appearance).toEqual({ style: RED })
  })

  it('a rule over another sheet and a name computes there', () => {
    const b: BookInput = { ...book([{ type: 'expression', ranges: ['A1:A5'], formula: "$A1>'T'!$A$1+lim", style: RED }]), names: [{ name: 'lim', ref: '1' }] }
    b.sheets.push({ name: 'T', cells: { A1: { value: 2 } } })
    const s = createSession(b, { clock: CLOCK, execution })
    expect(Object.entries(s.getCells('S', 'A1:A5')).filter(([, v]) => v.appearance).map(([a]) => a)).toEqual(['A4', 'A5'])
    s.setCells([{ sheet: 'T', cell: 'A1', input: { value: 0 } }])
    expect(Object.entries(s.getCells('S', 'A1:A5')).filter(([, v]) => v.appearance).map(([a]) => a)).toEqual(['A2', 'A3', 'A4', 'A5'])
  })
})
