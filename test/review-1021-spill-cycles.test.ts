/**
 * Arrays in a cycle through their own spill beside another array over the same cells (x-fuzz-modes-6, repro-i1, fuzz
 * seeds 1754, 3315, 3914, 3600, 10059, 1507): which array spills follows the formulas, as Excel decides for a book opened
 * from its file, whatever the order of the cells and sheets, and a session entering the cells in any order reaches it.
 */
import { describe, expect, it } from 'vitest'
import { compute, createInlineCalc, createSession, type BookInput, type CellInput, type ComputeResult, type SheetInput } from '../src/index.js'
import { applyEdits, createState } from '../src/engine.js'
import { run } from '../src/exec/budget.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
type Cells = Record<string, CellInput>

/** Every order of `xs`. */
const orders = <T>(xs: T[]): T[][] => (xs.length <= 1 ? [xs] : xs.flatMap((x, i) => orders([...xs.slice(0, i), ...xs.slice(i + 1)]).map((o) => [x, ...o])))

/** The book's sheets, each with its cells in the order `cells` lists them. */
const sheetsOf = (sheets: Record<string, Cells>, cells: [string, string, CellInput][]): SheetInput[] =>
  Object.keys(sheets).map((name) => ({ name, cells: Object.fromEntries(cells.filter((x) => x[0] === name).map((x) => [x[1], x[2]])) }))

/** The values of a result by `sheet!cell`, but those of the cells `book` holds a value in (`held`, a cycle's). */
function heldOut(r: ComputeResult, book: BookInput): Record<string, unknown> {
  const held = new Set(book.sheets.flatMap((x) => Object.entries(x.cells).flatMap(([k, c]) => (c && 'held' in c && c.held !== undefined ? [`${x.name}!${k}`] : []))))
  return Object.fromEntries(r.sheets.flatMap((x) => Object.entries(x.cells).filter(([k]) => !held.has(`${x.name}!${k}`)).map(([k, c]) => [`${x.name}!${k}`, c.value])))
}

/**
 * The book in every order of its cells: one hash computed, opened as a session (on the same thread and inline), and
 * entered cell by cell into empty sheets, each session holding what a full computation of its book holds, the book's
 * own answer (but for the values a cycle holds without iterative calculation). The values of the computed book are returned, by sheet and cell.
 */
async function oneAnswer(sheets: Record<string, Cells>, execution: 'closures' | 'eval', extra: Partial<BookInput> = {}) {
  const cells = Object.entries(sheets).flatMap(([name, xs]) => Object.entries(xs).map(([cell, input]) => [name, cell, input] as [string, string, CellInput]))
  const b: BookInput = { sheets: sheetsOf(sheets, cells), ...extra }
  const full = compute(b, { clock, execution })
  for (const order of orders(cells)) {
    expect(compute({ ...b, sheets: sheetsOf(sheets, order) }, { clock, execution }).hash).toBe(full.hash)
    const s = createSession({ ...b, sheets: sheetsOf(sheets, []) }, { clock, execution })
    for (const [sheet, cell, input] of order) {
      s.setCells([{ sheet, cell, input }])
      expect(s.getHash()).toBe(compute(s.getBook(), { clock, execution }).hash)
    }
    // without iterative calculation a cycle holds what its cells showed as it formed, which the order changes: the rest is the book's
    expect(heldOut(s.getResult(), s.getBook())).toEqual(heldOut(full, s.getBook()))
    if (b.iteration) expect(s.getHash()).toBe(full.hash)
    // a session edited back and forth comes to the same answer (a legacy array is cleared whole, so not that one)
    const [sheet, cell, input] = order.find((x) => !x[2].array)!
    s.setCells([{ sheet, cell, input: null }])
    expect(s.getHash()).toBe(compute(s.getBook(), { clock, execution }).hash)
    s.setCells([{ sheet, cell, input }])
    expect(heldOut(s.getResult(), s.getBook())).toEqual(heldOut(full, s.getBook()))
    s.close()
  }
  expect(createSession(b, { clock, execution }).getHash()).toBe(full.hash)
  const calc = createInlineCalc({ sliceMs: 1 })
  const inline = await calc.createSession({ ...b, sheets: sheetsOf(sheets, []) }, { clock, execution })
  for (const [sheet, cell, input] of [...cells].reverse()) await inline.setCells([{ sheet, cell, input }])
  expect(await inline.getHash()).toBe(compute(await inline.getBook(), { clock, execution }).hash)
  expect(heldOut(await inline.getResult(), await inline.getBook())).toEqual(heldOut(full, await inline.getBook()))
  await inline.close()
  return Object.fromEntries(full.sheets.map((x) => [x.name, Object.fromEntries(Object.entries(x.cells).map(([k, c]) => [k, c.value]))]))
}

const iteration = { iteration: {} }

describe.each(['closures', 'eval'] as const)('execution %s', (execution) => {
  it('an array in a cycle through its spill gives way to another array over its range, as Excel from the file (probe15, probe17)', async () => {
    const sort = { S: { B3: { formula: 'SORT(B1:C2)' }, C2: { formula: 'C4+1' }, A4: { formula: 'SEQUENCE(1,4)' } } }
    for (const extra of [{}, iteration]) {
      const c = (await oneAnswer(sort, execution, extra)).S
      expect([c.B3, c.C2, c.A4, c.D4]).toEqual([{ error: 'SPILL' }, 4, 1, 4])
    }
    const munit = { S: { C14: { formula: 'MUNIT(3)' }, D11: { formula: 'SWITCH(E11,30,1,H1:I4,2)' } } }
    for (const extra of [{}, iteration]) {
      const c = (await oneAnswer(munit, execution, extra)).S
      expect([c.D11, c.C14, c.E16]).toEqual([{ error: 'SPILL' }, 1, 1])
    }
  })

  it('arrays of two sheets in a cycle through each other\'s spill: one answer in either order of the sheets (probe16)', async () => {
    const q = { F5: { formula: '#DIV/0!' }, D4: { formula: 'LET(x,0.1,Sheet2!E8:F10)' } }
    const s2 = { E5: { formula: "'Q data'!E5:F8*2" }, D7: { formula: 'TRANSPOSE(D4:D6)' } }
    for (const extra of [{}, iteration]) {
      const a = await oneAnswer({ Sheet2: s2, 'Q data': q }, execution, extra)
      const b = await oneAnswer({ 'Q data': q, Sheet2: s2 }, execution, extra)
      expect([a.Sheet2.E5, a.Sheet2.D7, a.Sheet2.F7, a['Q data'].D4, a['Q data'].E6]).toEqual([{ error: 'SPILL' }, 0, 0, 0, 0])
      expect(b).toEqual(a)
    }
  })

  it('an array waiting for a spill that a cycle reads settles under iterative calculation (repro-i1)', async () => {
    const wait = { S: { B1: { formula: 'SEQUENCE(2)' }, A2: { formula: '{1,2}+B2' }, C1: { formula: 'B1+C2' }, C2: { formula: 'C1*1' } } }
    for (const extra of [{}, iteration, { iteration: { maxIterations: 1 } }]) {
      const c = (await oneAnswer(wait, execution, extra)).S
      expect([c.B1, c.B2, c.A2]).toEqual([1, 2, { error: 'SPILL' }])
    }
  })

  it('a session iterates afresh the cycles through spills it computes again from no spills (seeds 1754, 3914)', async () => {
    const it20 = { iteration: { maxIterations: 20, maxChange: 0 } }
    const unique = (await oneAnswer({ S: { A4: { formula: 'IF(A5:A7>1,D2:D3,365)' }, B2: { formula: 'A4#+1' } } }, execution, it20)).S
    expect([unique.A4, unique.A5, unique.A6, unique.B2, unique.B3]).toEqual([365, 0, 365, 366, 1])
    const sum = (await oneAnswer({ S: { C3: { formula: 'SUM(D1:D3)', array: 'C3:D3' }, D2: { value: 2 } } }, execution, iteration)).S
    expect([sum.C3, sum.D3]).toEqual([200, 200])
    // the cell typed over the array that reads itself, then retyped: the session iterates as a new one would
    const s = createSession({ sheets: [{ name: 'S', cells: { A4: { formula: 'UNIQUE(A4#)' }, B2: { formula: 'A4#+1' } } }], ...it20 }, { clock, execution })
    s.setCells([{ sheet: 'S', cell: 'A4', input: { formula: 'IF(A5:A7>1,D2:D3,365)' } }])
    expect(s.getHash()).toBe(compute(s.getBook(), { clock, execution }).hash)
    expect(s.getResult().sheets[0].cells.B3.value).toBe(1)
  })

  it('whether an array reads a cell its spill may cover follows its references, not what spilled (cyc seed 1507)', async () => {
    // C2 reads D3:E4, which D1 spills over and C2's own array may: C2 gives way, in every order and with or without iteration
    for (const extra of [{}, iteration]) {
      const c = (await oneAnswer({ S: { D1: { formula: 'SEQUENCE(3,3)' }, C2: { formula: 'UNIQUE(D3:E4)' } } }, execution, extra)).S
      expect([c.D1, c.F3, c.C2]).toEqual([1, 9, { error: 'SPILL' }])
    }
  })

  it('an array that gave way stays waiting while an iteration places the winner again (cyc seed 10059)', async () => {
    const cells = { B5: { formula: '{1,2;3,4}+0*N(C5)' }, B2: { formula: 'TRANSPOSE(A3:B5)' }, A6: { formula: 'SEQUENCE(3,3)+0*N(B2)' } }
    const c = (await oneAnswer({ S: cells }, execution, { iteration: { maxIterations: 100, maxChange: 0.001 } })).S
    expect([c.B5, c.A6, c.C8]).toEqual([{ error: 'SPILL' }, 1, 9])
  })

  it('an array that reads another\'s spill in a cycle under iterative calculation settles (seed 3315)', async () => {
    const cells = { B6: { formula: 'IF(E6>0,A3,10)' }, E7: { formula: 'UNIQUE(C6:D7)' }, A6: { formula: 'D2:E2*2' }, D5: { formula: 'TRANSPOSE(E7#)' } }
    const b: BookInput = { sheets: [{ name: 'S', cells }], iteration: { maxIterations: 3, maxChange: 0.001 } }
    const full = compute(b, { clock, execution })
    for (const order of orders(Object.entries(cells))) {
      expect(compute({ ...b, sheets: [{ name: 'S', cells: Object.fromEntries(order) }] }, { clock, execution }).hash).toBe(full.hash)
      const s = createSession({ ...b, sheets: [{ name: 'S', cells: {} }] }, { clock, execution })
      for (const [cell, input] of order) s.setCells([{ sheet: 'S', cell, input }])
      expect(s.getHash()).toBe(full.hash)
    }
  })
})

it('which of two arrays spills is found once per graph, not by a walk of what depends on them for each pair', () => {
  // 1,000 arrays read by a chain of 5,000 formulas, then an array typed into each one's range: the reads of the graph's
  // edges follow the chain a few times (some 40,000 reads), not once per pair (a walk for each pair read 10 million)
  const cells: Record<string, CellInput> = { E1: { formula: 'SUM(A1:D3000)' } }
  for (let i = 0; i < 1000; i++) cells[`A${3 * i + 1}`] = { formula: 'SEQUENCE(2,2)' }
  for (let k = 2; k <= 5000; k++) cells[`E${k}`] = { formula: `E${k - 1}+1` }
  const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 46300.5, seed: clock.seed })
  let reads = 0
  let succ = state.succ
  // each list of successors the graph is given counts its reads
  Object.defineProperty(state, 'succ', { get: () => succ, set: (v: number[][]) => (succ = new Proxy(v, { get: (t, k) => (typeof k === 'string' && /^\d+$/.test(k) && reads++, Reflect.get(t, k)) })) })
  run(applyEdits(state, Array.from({ length: 1000 }, (_, i) => ({ sheet: 'S', cell: `B${3 * i + 2}`, input: { formula: 'SEQUENCE(1,2)' } }))))
  expect(reads).toBeLessThan(500_000)
})
