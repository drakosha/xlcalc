/**
 * Hidden rows and filtered sheets beyond the Excel cases (`test/excel/workbook/hidden.json`): a session that hides
 * and shows rows or changes the filter mode recomputes SUBTOTAL and AGGREGATE as a full computation of its book
 * does, and the new fields are checked and hashed like the rest of the input.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, openSession, restructure, type BookInput, type CellInput, type LayoutEdit } from '../src/index.js'
import { mulberry32 } from '../src/synth.js'

const clock = { now: '2026-10-07T12:00:00Z', seed: 1 }

/** S: A1:A12 powers of two, SUBTOTAL and AGGREGATE over them in C, one of them read by D1. */
function book(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 12; r++) cells[`A${r}`] = { value: 2 ** (r - 1) }
  Object.assign(cells, {
    C1: { formula: 'SUBTOTAL(9,A1:A12)' },
    C2: { formula: 'SUBTOTAL(109,A1:A12)' },
    C3: { formula: 'AGGREGATE(9,5,A:A)' },
    C4: { formula: 'AGGREGATE(9,4,A1:A12)' },
    C5: { formula: 'SUBTOTAL(103,A1:B12)' },
    C6: { formula: 'SUM(A1:A12)' },
    D1: { formula: 'C2*2' },
  })
  return { sheets: [{ name: 'S', cells }] }
}

const values = (b: BookInput) => compute(b, { clock, display: false }).sheets[0].cells

describe('hidden rows and filtered sheets', () => {
  it('SUBTOTAL and AGGREGATE skip hidden rows by their code, SUBTOTAL 1–11 only on a filtered sheet', () => {
    const b = book()
    const hidden: BookInput = { sheets: [{ ...b.sheets[0], rows: { '2:3': { hidden: true }, '12': { hidden: true } }, columns: { A: { hidden: true } } }] }
    const v = values(hidden)
    expect([v.C1.value, v.C2.value, v.C3.value, v.C4.value, v.C5.value, v.C6.value]).toEqual([4095, 4095 - 2 - 4 - 2048, 4095 - 2 - 4 - 2048, 4095, 9, 4095])
    const f = values({ sheets: [{ ...hidden.sheets[0], filtered: true }] })
    expect([f.C1.value, f.C2.value, f.C4.value]).toEqual([4095 - 2 - 4 - 2048, 4095 - 2 - 4 - 2048, 4095])
  })

  it('a session that hides and shows rows and changes the filter mode gives the values of a full computation', () => {
    const rnd = mulberry32(5)
    const session = openSession(book(), { clock })
    for (let i = 0; i < 120; i++) {
      const lo = 1 + Math.floor(rnd() * 12)
      const hi = Math.min(12, lo + Math.floor(rnd() * 3))
      const edit: LayoutEdit = rnd() < 0.2 ? { sheet: 'S', filtered: rnd() < 0.5 } : { sheet: 'S', rows: lo === hi ? `${lo}` : `${lo}:${hi}`, set: { hidden: rnd() < 0.6 ? true : null } }
      const changes = session.setLayout([edit])
      const full = compute(session.book(), { clock })
      expect(session.result().sheets[0].cells, `edit ${i}: ${JSON.stringify(edit)}`).toEqual(full.sheets[0].cells)
      expect(session.result().hash).toBe(full.hash)
      for (const c of changes) expect(c.value).toEqual(full.sheets[0].cells[c.cell].value)
    }
    session.close()
  })

  it('setLayout returns the formulas whose count changed and their dependents', () => {
    const session = openSession(book(), { clock })
    const changes = session.setLayout([{ sheet: 'S', rows: '1', set: { hidden: true } }])
    expect(changes.map((c) => `${c.cell}=${c.value}`).sort()).toEqual(['C2=4094', 'C3=4094', 'C5=11', 'D1=8188'])
    expect(session.setLayout([{ sheet: 'S', filtered: true }]).map((c) => `${c.cell}=${c.value}`)).toEqual(['C1=4094'])
    expect(session.book().sheets[0].filtered).toBe(true)
    session.setLayout([{ sheet: 'S', filtered: false }])
    expect(session.book().sheets[0].filtered).toBeUndefined()
    session.close()
  })

  it('AGGREGATE 14–19 take an array, 1–13 do not', () => {
    const cells: Record<string, CellInput> = { A1: { value: 3 }, A2: { value: 0 }, A3: { value: 5 }, B1: { formula: 'AGGREGATE(14,6,10/A1:A3,1)' }, B2: { formula: 'AGGREGATE(15,0,A1:A3*2,2)' }, B3: { formula: 'AGGREGATE(9,6,A1:A3*2)' } }
    const v = values({ sheets: [{ name: 'S', cells }] })
    expect([v.B1.value, v.B2.value, v.B3.value]).toEqual([3.33333333333333, 6, { error: 'VALUE' }])
  })

  it('the fields are checked and go into the input hash only when present', () => {
    const b = book()
    expect(() => compute({ sheets: [{ ...b.sheets[0], filtered: 1 as unknown as boolean }] })).toThrow(InputError)
    expect(() => compute({ sheets: [{ ...b.sheets[0], rows: { '2': { hidden: 'yes' as unknown as boolean } } }] })).toThrow(/rows\["2"\]\.hidden: expected a boolean/)
    const session = openSession(b, { clock })
    expect(() => session.setLayout([{ sheet: 'S', filtered: 'no' as unknown as boolean }])).toThrow(/edits\[0\]\.filtered: expected a boolean/)
    session.close()
    const h0 = compute(b, { clock }).inputHash
    expect(compute({ sheets: [{ ...b.sheets[0], filtered: true }] }, { clock }).inputHash).not.toBe(h0)
    expect(compute({ sheets: [{ ...b.sheets[0], rows: { '2': { hidden: true } } }] }, { clock }).inputHash).not.toBe(h0)
  })
})

describe('structure operations on a filtered sheet', () => {
  /** S filtered: A1:A6 the values 1…6, B1:B6 `=A*10`, rows 2 and 4 hidden. */
  const filtered = (): BookInput => {
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 6; r++) Object.assign(cells, { [`A${r}`]: { value: r }, [`B${r}`]: { formula: `A${r}*10` } })
    return { sheets: [{ name: 'S', cells, rows: { '2': { hidden: true }, '4': { hidden: true } }, filtered: true }, { name: 'T', cells: {} }] }
  }

  it('a copy over hidden rows pastes the visible rows as values, in a session as in restructure', () => {
    const op = { op: 'copyRange', from: { sheet: 'S', range: 'A1:B6' }, to: { sheet: 'T', range: 'A1' } } as const
    const r = restructure(filtered(), op)
    expect(r.ok && r.book.sheets[1].cells).toEqual({ A1: { value: 1 }, B1: { value: 10 }, A2: { value: 3 }, B2: { value: 30 }, A3: { value: 5 }, B3: { value: 50 }, A4: { value: 6 }, B4: { value: 60 } })
    const session = openSession(filtered(), { clock })
    session.setCells([{ sheet: 'S', cell: 'A3', input: { value: 7 } }])
    const s = session.restructure(op)
    expect(s.ok && s.book.sheets[1].cells.B2).toEqual({ value: 70 })
    session.close()
  })

  it('deleting rows deletes the visible ones, and mapCell follows them', () => {
    const r = restructure(filtered(), { op: 'deleteRows', sheet: 'S', row: 1, count: 3 })
    if (!r.ok) throw new Error(r.message)
    expect(Object.keys(r.book.sheets[0].cells).filter((a) => a.startsWith('A')).map((a) => `${a}=${r.book.sheets[0].cells[a].value}`)).toEqual(['A1=2', 'A2=4', 'A3=5', 'A4=6'])
    expect(r.book.sheets[0].rows).toEqual({ '1': { hidden: true }, '2': { hidden: true } })
    expect([r.mapCell({ sheet: 'S', cell: 'A3' }), r.mapCell({ sheet: 'S', cell: 'A4' })]).toEqual([null, { sheet: 'S', cell: 'A2' }])
  })

  it('a copy that left rows out is refused over several runs of visible rows, as Excel refuses it', () => {
    const r = restructure(filtered(), { op: 'copyRange', from: { sheet: 'S', range: 'A1:B6' }, to: { sheet: 'S', range: 'D1:E6' } })
    expect(r.ok || r.code).toBe('filtered')
  })
})
