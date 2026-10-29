/**
 * SUBTOTAL and AGGREGATE beyond the Excel case set (`bench/coverage/sets/subtotal.ts`): a session follows the
 * cells that become or stop being nested (a formula with SUBTOTAL or AGGREGATE) the same as a full computation; one
 * range read cell by cell, skipping what they skip, answers as the range made whole.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput, type CellReplace, type CellInput } from '../src/index.js'
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
    const session = createSession(b, { clock })
    for (let i = 0; i < 150; i++) {
      const edit: CellReplace = { sheet: 'S', cell: `A${5 + Math.floor(rnd() * 8)}`, input: inputs[Math.floor(rnd() * inputs.length)] }
      session.setCells([edit])
      b = apply(b, edit)
      const full = compute(b, { clock })
      expect(session.getResult().sheets[0].cells, `edit ${i}: ${JSON.stringify(edit)}`).toEqual(full.sheets[0].cells)
      expect(session.getResult().hash).toBe(full.hash)
    }
    session.close()
  })

  it('the cell edited into a SUBTOTAL drops out of an enclosing SUBTOTAL at once', () => {
    const session = createSession(book(), { clock })
    const sum = () => session.getResult().sheets[0].cells.B1.value
    // A4 and A8 are nested: 1..12 times 3 without them
    expect(sum()).toBe(3 * (78 - 4 - 8))
    session.setCells([{ sheet: 'S', cell: 'A4', input: { value: 100 } }])
    expect(sum()).toBe(3 * (78 - 4 - 8) + 100)
    session.setCells([{ sheet: 'S', cell: 'A1', input: { formula: 'SUBTOTAL(9,A2:A3)' } }])
    expect(sum()).toBe(3 * (78 - 1 - 4 - 8) + 100)
    session.close()
  })

  it('one range read cell by cell, skipping hidden rows, nested cells and errors, answers as the range made whole', () => {
    const rnd = mulberry32(5)
    const pick = <T>(xs: T[]): T => xs[Math.floor(rnd() * xs.length)]
    for (let n = 0; n < 60; n++) {
      const cells: Record<string, CellInput> = {}
      const rows: Record<string, { hidden: boolean }> = {}
      const height = 5 + Math.floor(rnd() * 40)
      for (let r = 1; r <= height; r++) {
        if (rnd() < 0.2) rows[r] = { hidden: true }
        for (const col of 'ABC') {
          const x = rnd()
          const input = x < 0.5 ? { value: Math.round((rnd() - 0.3) * 1000) / 10 } : x < 0.6 ? { value: pick<string | boolean>(['t', true]) } : x < 0.65 ? { formula: pick(['1/0', 'NA()']) } : x < 0.7 ? { formula: pick(['SUBTOTAL(9,A1:A2)', 'AGGREGATE(9,6,B1:B2)']) } : x < 0.76 ? { value: pick([1e308, -1e308, 0.1, 0.2, 1e-310]) } : undefined
          if (input) cells[`${col}${r}`] = input
        }
      }
      // the same call with a blank cell as a second reference: the range made whole (`counted`)
      const calls: [string, string][] = []
      for (let i = 0; i < 40; i++) {
        const r1 = 1 + Math.floor(rnd() * height)
        const r2 = r1 + Math.floor(rnd() * (height - r1 + 3))
        const range = pick([`A${r1}:A${r2}`, `A${r1}:C${r2}`, 'B:B', `A${r1}:C${r1}`, 'A:C'])
        const code = pick([1, 2, 3, 4, 5, 9])
        const call = rnd() < 0.5 ? `SUBTOTAL(${pick([code, code + 100])},${range}` : `AGGREGATE(${code},${Math.floor(rnd() * 8)},${range}`
        calls.push([`${call})`, `${call},Z999)`])
      }
      calls.forEach(([one, whole], i) => Object.assign(cells, { [`E${i + 1}`]: { formula: one }, [`F${i + 1}`]: { formula: whole } }))
      const out = compute({ sheets: [{ name: 'S', cells, rows, filtered: rnd() < 0.5 }] }, { clock }).sheets[0].cells
      calls.forEach(([one], i) => expect(out[`E${i + 1}`].value, `book ${n}: ${one}`).toEqual(out[`F${i + 1}`].value))
    }
  })
})
