/**
 * Conditional formats under structure operations, by the rules measured in Excel (the set `cond/struct-held`), in
 * `restructure` and in a session (both execution modes).
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, ConditionalFormat, StructureOp } from '../src/contract.js'
import { openSession, restructure } from '../src/index.js'

const RED = { fill: { color: '#FFC7CE' } }
const SCALE: ConditionalFormat = { type: 'colorScale', ranges: ['C1:C10'], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] }

function book(): BookInput {
  const cells: BookInput['sheets'][number]['cells'] = {}
  for (let r = 1; r <= 10; r++) cells[`A${r}`] = { value: r }
  return {
    sheets: [
      {
        name: 'S',
        cells,
        conditionalFormats: [
          { type: 'cellIs', ranges: ['B2:B10'], operator: 'greaterThan', formulas: ['$Z$2'], style: RED },
          { type: 'expression', ranges: ['D2:D10', 'F2:F4'], formula: '$A2>3', style: { font: { bold: true } } },
          SCALE,
          { type: 'expression', ranges: ['H2:H5'], formula: "H2>'T'!$A$1", style: { font: { italic: true } } },
        ],
      },
      { name: 'T', cells: { A1: { value: 3 } }, conditionalFormats: [{ type: 'expression', ranges: ['A1:A3'], formula: '$A1>S!$A$3', style: RED }] },
    ],
  }
}

/** Each sheet's rules as `ranges formula`, in order. */
function rules(op: StructureOp, b = book()): Record<string, string[]> {
  const r = restructure(b, op)
  if (!r.ok) throw new Error(r.code)
  return Object.fromEntries(r.book.sheets.map((s) => [s.name, (s.conditionalFormats ?? []).map((x) => `${x.ranges.join(' ')} ${'formula' in x && x.formula ? x.formula : x.type === 'cellIs' ? x.formulas.join(';') : x.type}`)]))
}

describe('conditional formats under structure operations', () => {
  it('inserted rows join a range that holds the row above them, and move a range below', () => {
    expect(rules({ op: 'insertRows', sheet: 'S', row: 5, count: 2 }).S).toEqual(['B2:B12 $Z$2', 'D2:D12 F2:F6 $A2>3', 'C1:C12 colorScale', "H2:H7 H2>'T'!$A$1"])
    expect(rules({ op: 'insertRows', sheet: 'S', row: 1 }).S).toEqual(['B3:B11 $Z$3', 'D3:D11 F3:F5 $A3>3', 'C2:C11 colorScale', "H3:H6 H3>'T'!$A$1"])
    expect(rules({ op: 'insertColumns', sheet: 'S', column: 'D' }).S[2]).toBe('C1:D10 colorScale')
  })

  it('deleted lines shrink ranges; a colour scale that lost some goes after the others; a rule with nothing left goes; the first cell moves on', () => {
    expect(rules({ op: 'deleteRows', sheet: 'S', row: 5 }).S).toEqual(['B2:B9 $Z$2', 'D2:D9 F2:F4 $A2>3', "H2:H4 H2>'T'!$A$1", 'C1:C9 colorScale'])
    expect(rules({ op: 'deleteRows', sheet: 'S', row: 2 }).S).toEqual(['B2:B9 #REF!', 'D2:D9 F2:F3 $A2>3', "H2:H4 H2>'T'!$A$1", 'C1:C9 colorScale'])
    expect(rules({ op: 'deleteColumns', sheet: 'S', column: 'B' }).S[0]).toBe('C2:C10 E2:E4 $A2>3')
    expect(rules({ op: 'deleteColumns', sheet: 'S', column: 'A' }).S[1]).toBe('C2:C10 E2:E4 #REF!>3')
    expect(rules({ op: 'deleteRows', sheet: 'S', row: 1, count: 10 }).S).toEqual([])
  })

  it('a move takes its cells rules along, the cells it lands on keep theirs; a rule its move rewrote goes last; to another sheet, first', () => {
    expect(rules({ op: 'moveRange', from: { sheet: 'S', range: 'B8:B10' }, to: { sheet: 'S', cell: 'D8' } }).S.slice(0, 2)).toEqual(['B2:B7 D8:D10 $Z$2', 'D2:D10 F2:F4 $A2>3'])
    expect(rules({ op: 'moveRange', from: { sheet: 'S', range: 'A1:A10' }, to: { sheet: 'S', cell: 'K1' } }).S.at(-1)).toBe('D2:D10 F2:F4 $K2>3')
    const t = rules({ op: 'moveRange', from: { sheet: 'S', range: 'B2:B10' }, to: { sheet: 'T', cell: 'C1' } })
    expect(t.T).toEqual(["C1:C9 'S'!$Z$2", 'A1:A3 $A1>S!$A$3'])
  })

  it('a pasted rule is a new one, first; a fill that makes one rectangle with its rule grows it; cells pasted on lose theirs', () => {
    expect(rules({ op: 'copyRange', from: { sheet: 'S', range: 'B2:B3' }, to: { sheet: 'S', range: 'J5' } }).S.slice(0, 2)).toEqual(['J5:J6 $Z$2', 'B2:B10 $Z$2'])
    expect(rules({ op: 'copyRange', from: { sheet: 'S', range: 'B10' }, to: { sheet: 'S', range: 'B11:B12' } }).S[0]).toBe('B2:B12 $Z$2')
    expect(rules({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'B6' } }).S[0]).toBe('B2:B5 B7:B10 $Z$2')
    expect(rules({ op: 'copyRange', from: { sheet: 'S', range: 'B2:B3' }, to: { sheet: 'T', range: 'D1' } }).T).toEqual(['D1:D2 $Z$2', 'A1:A3 $A1>S!$A$3'])
  })

  it('sheets renamed and deleted: the references follow, a deleted sheet is #REF!', () => {
    expect(rules({ op: 'renameSheet', sheet: 'T', name: 'U V' }).S.at(-1)).toBe("H2:H5 H2>'U V'!$A$1")
    expect(rules({ op: 'deleteSheet', sheet: 'T' }).S.at(-1)).toBe('H2:H5 H2>#REF!')
    expect(rules({ op: 'renameSheet', sheet: 'S', name: 'Q' }).T).toEqual(["A1:A3 $A1>'Q'!$A$3"])
  })
})

describe.each(['closures', 'eval'] as const)('a session restructures its rules (%s)', (execution) => {
  it('the rules move with the operation and the looks follow', () => {
    const s = openSession(book(), { clock: { now: '2026-10-10T12:00:00Z', seed: 1 }, execution })
    expect(s.cells('S', 'C1').C1).toBeUndefined()
    s.setCells([{ sheet: 'S', cell: 'C1', input: { value: 1 } }, { sheet: 'S', cell: 'C2', input: { value: 2 } }])
    expect(s.cells('S', 'C1:C2').C2.look).toEqual({ style: { fill: { color: '#63BE7B' } } })
    const r = s.restructure({ op: 'insertRows', sheet: 'S', row: 1 })
    expect(r.ok).toBe(true)
    expect(s.book().sheets[0].conditionalFormats![2].ranges).toEqual(['C2:C11'])
    expect(s.cells('S', 'C3').C3.look).toEqual({ style: { fill: { color: '#63BE7B' } } })
  })
})
