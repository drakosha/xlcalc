/**
 * SUBTOTAL and AGGREGATE beyond the Excel case set (`bench/coverage/sets/subtotal.ts`): a session follows the
 * cells that become or stop being nested (a formula with SUBTOTAL or AGGREGATE) the same as a full computation.
 */
import { describe, expect, it } from 'vitest'
import { compute, openSession, type BookInput, type CellReplace, type CellInput } from '../src/index.js'
import { mulberry32 } from '../src/synth.js'

const clock = { now: '2026-10-06T12:00:00Z', seed: 1 }

function book(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 12; r++) cells[`A${r}`] = { value: r * 3 }
  cells.A4 = { formula: 'SUBTOTAL(9,A1:A3)' }
  cells.A8 = { formula: 'AGGREGATE(9,4,A5:A7)*2' }
  cells.B1 = { formula: 'SUBTOTAL(9,A1:A12)' }
  cells.B2 = { formula: 'SUBTOTAL(103,A:A)' }
  cells.B3 = { formula: 'AGGREGATE(14,0,A1:A12,2)' }
  cells.B4 = { formula: 'AGGREGATE(9,4,A1:A12)' }
  cells.B5 = { formula: 'SUM(A1:A12)-B1' }
  return { sheets: [{ name: 'S', cells }] }
}

function apply(b: BookInput, e: CellReplace): BookInput {
  const cells = { ...b.sheets[0].cells }
  if (e.input === null) delete cells[e.cell]
  else cells[e.cell] = e.input
  return { sheets: [{ name: 'S', cells }] }
}

describe('SUBTOTAL and AGGREGATE beyond the Excel cases', () => {
  it('a session skips the cells that become nested and counts those that stop being nested, as a full computation does', () => {
    const rnd = mulberry32(11)
    const inputs: (CellInput | null)[] = [
      { formula: 'SUBTOTAL(9,A1:A2)' },
      { formula: 'IF(TRUE,AGGREGATE(4,6,A1:A2),0)' },
      { formula: 'A1*2' },
      { value: 7 },
      { formula: 'SUBTOTAL(1,A1)+1' },
      null,
      { formula: 'SUBTOTAL(9,A1' }, // invalid
    ]
    let b = book()
    const session = openSession(b, { clock })
    for (let i = 0; i < 150; i++) {
      const edit: CellReplace = { sheet: 'S', cell: `A${5 + Math.floor(rnd() * 8)}`, input: inputs[Math.floor(rnd() * inputs.length)] }
      session.setCells([edit])
      b = apply(b, edit)
      const full = compute(b, { clock })
      expect(session.result().sheets[0].cells, `edit ${i}: ${JSON.stringify(edit)}`).toEqual(full.sheets[0].cells)
      expect(session.result().hash).toBe(full.hash)
    }
    session.close()
  })

  it('the cell edited into a SUBTOTAL drops out of an enclosing SUBTOTAL at once', () => {
    const session = openSession(book(), { clock })
    const sum = () => session.result().sheets[0].cells.B1.value
    // A4 and A8 are nested: 1..12 times 3 without them
    expect(sum()).toBe(3 * (78 - 4 - 8))
    session.setCells([{ sheet: 'S', cell: 'A4', input: { value: 100 } }])
    expect(sum()).toBe(3 * (78 - 4 - 8) + 100)
    session.setCells([{ sheet: 'S', cell: 'A1', input: { formula: 'SUBTOTAL(9,A2:A3)' } }])
    expect(sum()).toBe(3 * (78 - 1 - 4 - 8) + 100)
    session.close()
  })
})
