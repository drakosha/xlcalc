/**
 * Conditional formats, from the review of 21.10: a copy or a fill that moves a rule's first cell makes its formulas for
 * the new first cell, as clearing formats over the same cells does, so the cells the paste leaves keep their looks.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, ConditionalFormat, StructureOp } from '../src/contract.js'
import { createSession, editStructure } from '../src/index.js'

const RED = { fill: { color: '#FF0000' } }
const B = [0, 10, 10, 0, 10, 0, 0, 10, 0, 10, 0, 0]

function book(rule: ConditionalFormat): BookInput {
  const cells: BookInput['sheets'][number]['cells'] = { D1: { value: 'x' }, D2: { value: 'y' } }
  B.forEach((v, i) => (cells[`B${i + 1}`] = { value: v }))
  return { sheets: [{ name: 'S', cells, conditionalFormats: [rule] }, { name: 'T', cells: { A1: { value: 1 } } }] }
}
const expr = (range: string, formula: string): ConditionalFormat => ({ type: 'expression', ranges: [range], formula, style: RED })
const cellIs: ConditionalFormat = { type: 'cellIs', ranges: ['A1:A6'], operator: 'lessThan', formulas: ['$B1'], style: RED }

/** The rule's ranges and formulas after the operation. */
function after(rule: ConditionalFormat, op: StructureOp): string[] {
  const r = editStructure(book(rule), op)
  if (!r.ok) throw new Error(r.code)
  return (r.book.sheets[0].conditionalFormats ?? []).map((x) => `${x.ranges.join(' ')} ${x.type === 'cellIs' ? x.formulas.join(';') : 'formula' in x ? x.formula : ''}`)
}

const copyOver: StructureOp = { op: 'copyRange', from: { sheet: 'S', range: 'D1:D2' }, to: { sheet: 'S', range: 'A1:A2' } }
const fromOther: StructureOp = { op: 'copyRange', from: { sheet: 'T', range: 'A1:A2' }, to: { sheet: 'S', range: 'A1:A2' } }
const fillUp: StructureOp = { op: 'copyRange', from: { sheet: 'S', range: 'A5' }, to: { sheet: 'S', range: 'A3:A4' } }

describe('a copy or a fill that moves a rule first cell', () => {
  it('rewrites the rule formulas for the new first cell', () => {
    expect(after(expr('A1:A6', '$B1>3'), copyOver)).toEqual(['A3:A6 $B3>3'])
    expect(after(expr('A1:A6', '$B1>3'), fromOther)).toEqual(['A3:A6 $B3>3'])
    expect(after(cellIs, copyOver)).toEqual(['A3:A6 $B3'])
    expect(after(expr('A5:A10', '$B5>3'), fillUp)).toEqual(['A3:A10 $B3>3'])
  })

  it('leaves a rule whose first cell stays as it was', () => {
    expect(after(expr('A5:A10', '$B5>3'), { op: 'copyRange', from: { sheet: 'S', range: 'A10' }, to: { sheet: 'S', range: 'A11:A12' } })).toEqual(['A5:A12 $B5>3'])
    expect(after(expr('A1:A6', '$B1>3'), { op: 'copyRange', from: { sheet: 'S', range: 'D1' }, to: { sheet: 'S', range: 'A3' } })).toEqual(['A1:A2 A4:A6 $B1>3'])
  })
})

describe.each(['closures', 'eval'] as const)('a session keeps the looks of the cells a paste leaves (%s)', (execution) => {
  /** The cells of `range` red (R) or not (-), after the operation. */
  function looks(rule: ConditionalFormat, op: StructureOp, rows: number[]): string {
    const s = createSession(book(rule), { execution })
    expect(s.editStructure(op).ok).toBe(true)
    const got = s.getCells('S', `A${rows[0]}:A${rows.at(-1)}`)
    return rows.map((r) => (got[`A${r}`]?.appearance?.style?.fill ? 'R' : '-')).join('')
  }

  it('after a paste over the rule first cells', () => {
    expect(looks(expr('A1:A6', '$B1>3'), copyOver, [3, 4, 5, 6])).toBe('R-R-')
    expect(looks(cellIs, copyOver, [3, 4, 5, 6])).toBe('R-R-')
  })

  it('after a fill up', () => {
    expect(looks(expr('A5:A10', '$B5>3'), fillUp, [3, 4, 5, 6, 7, 8, 9, 10])).toBe('R-R--R-R')
  })
})

describe('a rule formula made for a new first cell wraps round the sheet edge, as the rule computes', () => {
  /** B1:B8 = h a a b b c c d, the rule over A3:A8 comparing each row with the one above. */
  const above = (): BookInput => {
    const cells: BookInput['sheets'][number]['cells'] = {}
    ;['h', 'a', 'a', 'b', 'b', 'c', 'c', 'd'].forEach((v, i) => (cells[`B${i + 1}`] = { value: v }))
    return { sheets: [{ name: 'S', cells, conditionalFormats: [expr('A3:A8', '$B3<>$B2')] }] }
  }
  const fill: StructureOp = { op: 'copyRange', from: { sheet: 'S', range: 'A3' }, to: { sheet: 'S', range: 'A1:A2' } }
  const text = (op: StructureOp) => {
    const r = editStructure(above(), op)
    if (!r.ok) throw new Error(r.code)
    return r.book.sheets[0].conditionalFormats!.map((x) => `${x.ranges.join(' ')} ${'formula' in x ? x.formula : ''}`)
  }

  it('a fill up over the sheet first row, and a copy of the rule to the first rows elsewhere', () => {
    expect(text(fill)).toEqual(['A1:A8 $B1<>$B1048576'])
    expect(text({ op: 'copyRange', from: { sheet: 'S', range: 'A3:A4' }, to: { sheet: 'S', range: 'D1' } })).toEqual(['D1:D2 $B1<>$B1048576', 'A3:A8 $B3<>$B2'])
  })

  it.each(['closures', 'eval'] as const)('the cells the rule covers show as the rule computes them (%s)', (execution) => {
    const s = createSession(above(), { execution })
    expect(s.editStructure(fill).ok).toBe(true)
    const got = s.getCells('S', 'A1:A8')
    expect([1, 2, 3, 4, 5, 6, 7, 8].map((r) => (got[`A${r}`]?.appearance?.style?.fill ? 'R' : '-')).join('')).toBe('RR-R-R-R')
  })
})
