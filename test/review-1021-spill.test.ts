/** Reproductions from the review of 21.10 on spills and legacy arrays: books that settle, areas a session keeps whole. */
import { describe, expect, it } from 'vitest'
import { compute, createInlineCalc, createSession, type BookInput, type CellInput } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const book = (cells: Record<string, CellInput>, extra: Partial<BookInput> = {}): BookInput => ({ sheets: [{ name: 'S', cells }], ...extra })

/** The book computed, opened as a session and entered cell by cell in both orders: each session's hash is a full computation's. */
async function sameEverywhere(b: BookInput, execution: 'closures' | 'eval') {
  const full = compute(b, { clock, execution })
  expect(createSession(b, { clock, execution }).getHash()).toBe(full.hash)
  const calc = createInlineCalc({ sliceMs: 1 })
  const inline = await calc.createSession(b, { clock, execution })
  expect(await inline.getHash()).toBe(full.hash)
  await inline.close()
  const cells = Object.entries(b.sheets[0].cells)
  for (const order of [cells, [...cells].reverse()]) {
    const s = createSession({ ...b, sheets: [{ ...b.sheets[0], cells: {} }] }, { clock, execution })
    for (const [cell, input] of order) s.setCells([{ sheet: 'S', cell, input }])
    expect(s.getHash()).toBe(compute(s.getBook(), { clock, execution }).hash)
  }
  return full.sheets[0].cells
}

describe.each(['closures', 'eval'] as const)('execution %s', (execution) => {
  it('an array that reads a blocked array and spills over its range settles (x-fuzz-modes-4)', async () => {
    const c = await sameEverywhere(book({ A1: { formula: 'SEQUENCE(2)' }, A2: { formula: 'A1+{1;2}' } }), execution)
    expect([c.A1.value, c.A2.value, c.A3.value]).toEqual([{ error: 'SPILL' }, { error: 'SPILL' }, { error: 'SPILL' }])
    const d = await sameEverywhere(book({ B12: { formula: 'SEQUENCE(3,3)' }, D13: { formula: 'SUM(B12)+{1;2;3}' } }), execution)
    expect([d.B12.value, d.D13.value, d.D15.value]).toEqual([{ error: 'SPILL' }, { error: 'SPILL' }, { error: 'SPILL' }])
    const e = await sameEverywhere(book({ A1: { formula: 'SEQUENCE(2)' }, A2: { formula: 'IFERROR(A1,0)+{1;2}' } }), execution)
    expect([e.A1.value, e.A2.value, e.A3.value]).toEqual([{ error: 'SPILL' }, 1, 2])
  })

  it('a reader of an array in a cycle through its spill settles under iterative calculation (x-fuzz-modes-4)', async () => {
    const cells = { G4: { formula: 'SUM(E7:F10)' }, E8: { formula: 'IF(D6:E9>365,D6:F8,7)' } }
    // each iteration starts from 0: in one, G4 comes before E8 spills, in many it reads the spill
    expect((await sameEverywhere(book(cells, { iteration: { maxIterations: 1 } }), execution)).G4.value).toBe(0)
    expect((await sameEverywhere(book(cells, { iteration: {} }), execution)).G4.value).toBe(42)
  })
})

describe.each(['closures', 'eval'] as const)('legacy arrays, execution %s', (execution) => {
  it('two array areas that cross with neither top inside the other are refused (x-fuzz-session-6)', () => {
    const crossed = book({ J18: { formula: '{1;2}', array: 'J18:J19' }, I19: { formula: '{3,4}', array: 'I19:J19' } })
    expect(() => compute(crossed, { clock, execution })).toThrow(/I19:J19\) overlaps the area of the array formula at J18/)
    const s = createSession(book({ J18: { formula: '{1;2}', array: 'J18:J19' } }), { clock, execution })
    const before = s.getBook()
    expect(() => s.setCells([{ sheet: 'S', cell: 'I19', input: { formula: '{3,4}', array: 'I19:J19' } }])).toThrow(/overlaps the area of the array formula at S!J18/)
    // two new areas of one edit
    expect(() => s.setCells([{ sheet: 'S', cell: 'A2', input: { formula: '1', array: 'A2:B2' } }, { sheet: 'S', cell: 'B1', input: { formula: '2', array: 'B1:B3' } }])).toThrow(/overlaps/)
    expect(s.getBook()).toEqual(before)
    // an array entered again over a new area that crosses its old one replaces it
    s.setCells([{ sheet: 'S', cell: 'J18', input: { formula: '{5,6}', array: 'J18:K18' } }, { sheet: 'S', cell: 'J19', input: null }])
    expect(s.getHash()).toBe(compute(s.getBook(), { clock, execution }).hash)
  })
})

describe.each(['closures', 'eval'] as const)('a legacy array whose formula does not prepare, execution %s', (execution) => {
  const edit = (cell: string, input: CellInput | null) => [{ sheet: 'S', cell, input }]
  it.each(['TRANSPOSE(Nope!B2:D3)', 'SUM(A1', 'FOO(A1)'])('%s keeps its area: an edit inside it is refused (x-fuzz-session-2)', async (formula) => {
    const b = book({ A1: { value: 1 }, N1: { formula, array: 'N1:O3' } })
    const s = createSession(b, { clock, execution })
    expect(() => s.setCells(edit('O2', { value: 5 }))).toThrow(/O2 is part of the array formula at S!N1 \(N1:O3\)/)
    expect(() => s.replaceCells(edit('O2', { value: 5 }))).toThrow(/part of the array formula/)
    expect(() => s.setCells(edit('M2', { formula: '1', array: 'M2:N2' }))).toThrow(/overlaps the area of the array formula at S!N1/)
    expect(compute(s.getBook(), { clock, execution }).hash).toBe(s.getHash())
    const calc = createInlineCalc({ sliceMs: 1 })
    const inline = await calc.createSession(b, { clock, execution })
    await expect(inline.setCells(edit('O2', { value: 5 }))).rejects.toThrow(/part of the array formula/)
    await inline.close()
    // entered by an edit, then cleared whole: the area goes with it
    const t = createSession(book({}), { clock, execution })
    t.setCells(edit('N1', { formula, array: 'N1:O3' }))
    expect(() => t.setCells(edit('N3', { value: 'x' }))).toThrow(/part of the array formula/)
    t.setCells(['N1', 'N2', 'N3', 'O1', 'O2', 'O3'].map((cell) => ({ sheet: 'S', cell, input: null })))
    t.setCells(edit('O2', { value: 5 }))
    expect(compute(t.getBook(), { clock, execution }).hash).toBe(t.getHash())
  })
})

describe.each(['closures', 'eval'] as const)('a legacy array over a table, execution %s', (execution) => {
  const b: BookInput = { sheets: [{ name: 'S', cells: { F13: { value: 'Name' }, G13: { value: 'Qty' } }, tables: [{ name: 'Sales', range: 'F13:G15', columns: [{ name: 'Name' }, { name: 'Qty' }] }] }] }
  const edit = (cell: string, array: string) => [{ sheet: 'S', cell, input: { formula: 'H16:I16*2', array } }]
  it('of several cells is refused by the session as by the book check (x-fuzz-session-5)', async () => {
    const s = createSession(b, { clock, execution })
    for (const area of ['F14:G14', 'E14:F14']) {
      expect(() => s.setCells(edit(area.slice(0, 3), area))).toThrow(/overlaps the table Sales \(Excel does not allow array formulas of several cells in tables\)/)
      expect(() => s.replaceCells(edit(area.slice(0, 3), area))).toThrow(/overlaps the table Sales/)
      expect(() => s.validateEntry('S', area.slice(0, 3), edit(area.slice(0, 3), area)[0].input)).toThrow(/overlaps the table Sales/)
    }
    const calc = createInlineCalc({ sliceMs: 1 })
    const inline = await calc.createSession(b, { clock, execution })
    await expect(inline.setCells(edit('F14', 'F14:G14'))).rejects.toThrow(/overlaps the table Sales/)
    await inline.close()
    // one cell is allowed, as the book check allows it
    s.setCells(edit('F14', 'F14'))
    expect(compute(s.getBook(), { clock, execution }).hash).toBe(s.getHash())
  })
})
