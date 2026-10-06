/**
 * Lookup and database functions beyond the Excel case set: the formula's own cell (ROW, COLUMN), sessions,
 * large whole columns, and the documented differences (INDEX gives a value, LOOKUP past its result range).
 */
import { describe, expect, it } from 'vitest'
import { compute, openSession, type CellInput, type CellValue } from '../src/index.js'

const err = (code: Extract<CellValue, { error: unknown }>['error']): CellValue => ({ error: code })
const clock = { now: '2026-10-06T12:00:00Z', seed: 1 }

function sheet(cells: Record<string, CellInput>): Record<string, CellValue> {
  const r = compute({ sheets: [{ name: 'S', cells }] }, { clock })
  return Object.fromEntries(Object.entries(r.sheets[0].cells).map(([a, c]) => [a, c.value]))
}

describe('lookup and database functions beyond the Excel cases', () => {
  it('ROW and COLUMN without an argument give the formula cell; of a range its first row or column', () => {
    const v = sheet({ C7: { formula: 'ROW()' }, D7: { formula: 'COLUMN()' }, E7: { formula: 'ROW(A2:A4)' }, F7: { formula: 'SUM(ROW(A2:A4))' } })
    // Excel passes {2;3;4} to SUM: 9 there (README, known differences)
    expect([v.C7, v.D7, v.E7, v.F7]).toEqual([7, 4, 2, 2])
  })

  it('INDEX gives a value, not a reference: a whole column of a range is #VALUE! even in SUM (Excel sums it)', () => {
    const v = sheet({ A1: { value: 1 }, B1: { value: 2 }, A2: { value: 3 }, B2: { value: 4 }, D1: { formula: 'SUM(INDEX(A1:B2,0,2))' }, D2: { formula: 'INDEX(A1:B2,2,2)' } })
    expect([v.D1, v.D2]).toEqual([err('VALUE'), 4])
  })

  it('LOOKUP past the end of a shorter result range is #REF! (Excel reads on)', () => {
    const cells: Record<string, CellInput> = { E1: { formula: 'LOOKUP(3,A1:A3,B1:B2)' }, E2: { formula: 'LOOKUP(2,A1:A3,B1:B2)' } }
    for (let r = 1; r <= 3; r++) cells[`A${r}`] = { value: r }
    cells.B1 = { value: 'x' }
    cells.B2 = { value: 'y' }
    const v = sheet(cells)
    expect([v.E1, v.E2]).toEqual([err('REF'), 'y'])
  })

  it('approximate lookups over whole columns of 200,000 sorted numbers, exact ones too', () => {
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 200_000; r++) {
      cells[`A${r}`] = { value: r * 2 }
      cells[`B${r}`] = { value: r }
    }
    const keys = [1, 2, 3, 199_999, 400_000, 400_001]
    keys.forEach((k, i) => {
      cells[`D${i + 1}`] = { formula: `VLOOKUP(${k},A:B,2)` }
      cells[`E${i + 1}`] = { formula: `MATCH(${k},A:A,0)` }
    })
    const v = sheet(cells)
    expect(keys.map((_, i) => v[`D${i + 1}`])).toEqual([err('NA'), 1, 1, 99_999, 200_000, 200_000])
    expect(keys.map((_, i) => v[`E${i + 1}`])).toEqual([err('NA'), 1, err('NA'), err('NA'), 200_000, err('NA')])
  })

  it('a session recomputes lookups and database functions after edits in their ranges', () => {
    const cells: Record<string, CellInput> = {
      A1: { value: 'Tree' },
      B1: { value: 'Profit' },
      A2: { value: 'Apple' },
      B2: { value: 10 },
      A3: { value: 'Pear' },
      B3: { value: 20 },
      D1: { value: 'Tree' },
      D2: { value: 'Pear' },
      F1: { formula: 'VLOOKUP("pear",A2:B3,2,FALSE)' },
      F2: { formula: 'DSUM(A1:B3,"Profit",D1:D2)' },
      F3: { formula: 'INDEX(B:B,MATCH("Apple",A:A,0))' },
    }
    const s = openSession({ sheets: [{ name: 'S', cells }] }, { clock })
    const changed = s.setCells([
      { sheet: 'S', cell: 'B3', input: { value: 25 } },
      { sheet: 'S', cell: 'B2', input: { value: 11 } },
    ])
    expect(Object.fromEntries(changed.filter((c) => c.cell.startsWith('F')).map((c) => [c.cell, c.value]))).toEqual({ F1: 25, F2: 25, F3: 11 })
    expect(s.result().hash).toBe(compute({ sheets: [{ name: 'S', cells: { ...cells, B3: { value: 25 }, B2: { value: 11 } } }] }, { clock }).hash)
    s.close()
  })

  it('database criteria rows past the occupied part are blank and select every record', () => {
    const v = sheet({ A1: { value: 'Tree' }, A2: { value: 'Pear' }, A3: { value: 'Apple' }, C1: { value: 'Tree' }, C2: { value: 'Pear' }, E1: { formula: 'DCOUNTA(A1:A3,"Tree",C1:C2)' }, E2: { formula: 'DCOUNTA(A1:A3,"Tree",C:C)' } })
    expect([v.E1, v.E2]).toEqual([1, 2])
  })
})
