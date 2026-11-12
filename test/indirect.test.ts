/**
 * INDIRECT and OFFSET in the graph (`src/graph/dynamic.ts`): a formula computed again when a cell its found reference
 * reads changes, when what it finds changes, and in the right order when it reads formulas computed after it in the
 * order; reads that close a cycle and an edit that breaks it; a name that brings INDIRECT in later. Every kind of
 * session in both execution modes, each step's hash that of a full computation of the session's book. A book without
 * either computes as before: no state of reads at all.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, type AsyncSession, type BookInput, type CellInput, type Execution } from '../src/index.js'
import { createState } from '../src/engine.js'
import { synthBook } from '../src/synth.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-29T12:00:00Z', seed: 1 }
afterAll(closeCalcs)

const book = (cells: Record<string, CellInput>, more: Partial<BookInput> = {}): BookInput => ({ sheets: [{ name: 'S', cells }, { name: 'T', cells: { A1: { value: 'tee' }, B1: { formula: 'S!A1*100' } } }], ...more })

/** A chain read backwards by INDIRECT: each formula reads the one below it, which the order may compute later. */
const CHAIN: Record<string, CellInput> = {
  A1: { value: 1 },
  B1: { value: 'C3' },
  C1: { formula: 'INDIRECT("C2")+1' },
  C2: { formula: 'INDIRECT("C3")+1' },
  C3: { formula: 'INDIRECT("C4")+1' },
  C4: { formula: 'A1*10' },
  D1: { formula: 'INDIRECT(B1)' },
  D2: { formula: 'SUM(OFFSET(C1,A1,0,2,1))' },
  D3: { formula: 'SUM(C1:OFFSET(C1,A1,0))' },
  D4: { formula: 'INDIRECT("T!B1")' },
}

describe.each(['closures', 'eval'] as Execution[])('INDIRECT and OFFSET in a session, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)
    const at = async (s: AsyncSession, refs: string) => {
      expect(await s.getHash()).toBe(compute(await s.getBook(), { clock, execution }).hash)
      return Object.fromEntries(Object.entries(await s.getCells('S', refs)).map(([k, c]) => [k, c.value]))
    }
    const set = (s: AsyncSession, cell: string, input: CellInput | null) => s.setCells([{ sheet: 'S', cell, input }])

    it('reads formulas computed later, follows its target and the cells there; undo and redo', async () => {
      const s = await open(book(CHAIN), { clock, execution })
      try {
        expect(await at(s, 'C1:D4')).toEqual({ C1: 13, C2: 12, C3: 11, C4: 10, D1: 11, D2: 23, D3: 25, D4: 100 })
        // a cell the found references read
        await set(s, 'A1', { value: 2 })
        expect(await at(s, 'C1:D4')).toEqual({ C1: 23, C2: 22, C3: 21, C4: 20, D1: 21, D2: 41, D3: 66, D4: 200 })
        // the text names another cell, then a formula's cell after it in the order
        await set(s, 'B1', { value: 'C4' })
        expect(await at(s, 'D1')).toEqual({ D1: 20 })
        await set(s, 'C4', { formula: 'A1*100' })
        expect(await at(s, 'C1:D4')).toMatchObject({ C1: 203, C2: 202, C3: 201, C4: 200, D1: 200 })
        await s.undo()
        expect(await at(s, 'C1:D4')).toMatchObject({ C1: 23, C2: 22, C3: 21, C4: 20, D1: 20 })
        await s.undo()
        expect(await at(s, 'D1')).toEqual({ D1: 21 })
        await s.redo()
        expect(await at(s, 'D1')).toEqual({ D1: 20 })
      } finally {
        await s.close()
      }
    })

    it('a text that names the formula itself or a cell over it is a cycle; another text breaks it', async () => {
      const s = await open(book({ A1: { value: 5 }, B1: { value: 'A1' }, C1: { formula: 'INDIRECT(B1)*2' }, C2: { formula: 'C1+1' } }), { clock, execution })
      try {
        expect(await at(s, 'C1:C2')).toEqual({ C1: 10, C2: 11 })
        await set(s, 'B1', { value: 'C2' })
        expect(await s.getCellIssue('S', 'C1')).toMatchObject({ error: 'CYCLE' })
        expect(await at(s, 'C1:C2')).toEqual({ C1: 10, C2: 11 })
        await set(s, 'B1', { value: 'C1' })
        expect(await s.getCellIssue('S', 'C1')).toMatchObject({ error: 'CYCLE' })
        await set(s, 'B1', { value: 'A1' })
        expect(await s.getCellIssue('S', 'C1')).toBeNull()
        expect(await at(s, 'C1:C2')).toEqual({ C1: 10, C2: 11 })
        await set(s, 'A1', { value: 7 })
        expect(await at(s, 'C1:C2')).toEqual({ C1: 14, C2: 15 })
      } finally {
        await s.close()
      }
    })

    it('a name set later brings INDIRECT into a formula, which then follows its target', async () => {
      const s = await open(book({ A1: { value: 3 }, A2: { value: 4 }, B1: { value: 'A1' }, C1: { formula: 'pick*2' } }, { names: [{ name: 'pick', ref: "'S'!$A$2" }] }), { clock, execution })
      try {
        expect(await at(s, 'C1')).toEqual({ C1: 8 })
        await s.setNames([{ name: 'pick', ref: "INDIRECT('S'!$B$1)" }])
        expect(await at(s, 'C1')).toEqual({ C1: 6 })
        await set(s, 'A1', { value: 10 })
        expect(await at(s, 'C1')).toEqual({ C1: 20 })
        await set(s, 'B1', { value: 'A2' })
        expect(await at(s, 'C1')).toEqual({ C1: 8 })
      } finally {
        await s.close()
      }
    })

    it("SUMIF's sum range resized to a found criteria range, or found itself, follows the cells it is resized over", async () => {
      const cells: Record<string, CellInput> = {
        B1: { value: 30 }, B2: { value: 10 }, B3: { value: 20 }, C1: { value: 300 }, C2: { value: 100 }, C3: { value: 200 },
        E1: { formula: 'SUMIF(OFFSET(B1,1,0,2),">15",C1)' },
        E2: { formula: 'SUMIF(INDIRECT("B2:B3"),">15",C1)' },
        F1: { formula: 'SUMIF(B2:B3,">15",OFFSET(C1,0,0))' },
        F2: { formula: 'AVERAGEIF(OFFSET(B1,1,0,2),">15",C1)' },
        G1: { formula: 'SUMIF(OFFSET(B1,{0;1},0,2),">15",C1)' },
      }
      const s = await open(book(cells), { clock, execution })
      try {
        expect(await at(s, 'E1:G2')).toEqual({ E1: 100, E2: 100, F1: 100, F2: 100, G1: 300, G2: 100 })
        // C2 lies past the sum range's own cell C1, inside it once resized
        await set(s, 'C2', { value: 1000 })
        expect(await at(s, 'E1:G2')).toEqual({ E1: 1000, E2: 1000, F1: 1000, F2: 1000, G1: 300, G2: 1000 })
      } finally {
        await s.close()
      }
    })

    it('a structure operation keeps the text and moves the references of OFFSET', async () => {
      const s = await open(book({ A1: { value: 1 }, A2: { value: 2 }, B1: { formula: 'INDIRECT("A2")' }, B2: { formula: 'OFFSET(A1,1,0)' } }), { clock, execution })
      try {
        await s.editStructure({ op: 'insertRows', sheet: 'S', row: 1, count: 1 })
        const cells = (await s.getBook()).sheets[0].cells
        expect([cells.B2.formula, cells.B3.formula]).toEqual(['INDIRECT("A2")', 'OFFSET(A2,1,0)'])
        expect(await at(s, 'B2:B3')).toEqual({ B2: 1, B3: 2 })
      } finally {
        await s.close()
      }
    })
  })
})

describe('the reads converge', () => {
  it('reads that change with every round they are computed in end as a cycle, not in a loop', () => {
    // A1 reads B1 while B1 is 0; B1 reads A1 while A1 reads nothing: each computed after the other turns the other round
    const cells = { A1: { formula: 'IF(N(INDIRECT("B1"))=0,INDIRECT("B1"),1)' }, B1: { formula: 'IF(N(INDIRECT("A1"))=1,0,INDIRECT("A1"))' } }
    const r = compute(book(cells), { clock })
    expect(r.sheets[0].cells.A1.value).toBe(0)
    expect(r.diagnostics.filter((d) => d.kind === 'cell' && d.error === 'CYCLE').map((d) => d.kind === 'cell' && `${d.sheet}!${d.cell}`)).toEqual(['S!A1', 'S!B1', 'T!B1'])
  })

  it('a chain of 200 INDIRECTs read backwards settles', () => {
    const cells: Record<string, CellInput> = { A201: { value: 1 } }
    for (let i = 1; i <= 200; i++) cells[`A${i}`] = { formula: `INDIRECT("A${i + 1}")+1` }
    expect(compute(book(cells), { clock }).sheets[0].cells.A1.value).toBe(201)
  })
})

describe.each(['closures', 'eval'] as Execution[])('a reference used for its place alone, %s', (execution) => {
  it('is no circular reference over the formula\'s own cell; one read is', () => {
    const cells: Record<string, CellInput> = {
      A1: { value: 4 },
      A2: { formula: 'SUMPRODUCT(ROW(INDIRECT("1:"&A1)))' },
      A3: { formula: 'ROWS(A1:OFFSET(A1,5,0))' },
      A4: { formula: 'ROW(INDIRECT("A4"))+COLUMNS(INDEX(INDIRECT("A1:C9"),2,0))' },
      A5: { formula: 'COLUMN(OFFSET(INDIRECT("A5"),0,2))+CELL("row",INDIRECT("A5"))' },
      A6: { formula: 'ROWS(INDIRECT("A1:A9"))+SUM(INDIRECT("A1"))' },
      A7: { formula: 'SUM(INDIRECT("A7"))' },
    }
    const r = compute(book(cells), { clock, execution })
    expect(Object.fromEntries(['A2', 'A3', 'A4', 'A5', 'A6'].map((a) => [a, r.sheets[0].cells[a].value]))).toEqual({ A2: 10, A3: 6, A4: 7, A5: 8, A6: 13 })
    expect(r.diagnostics.flatMap((d) => (d.kind === 'cell' && d.sheet === 'S' ? [d.cell] : []))).toEqual(['A7'])
  })
})

describe('a book without INDIRECT and OFFSET', () => {
  it('keeps no reads: no state of them and none on any formula', () => {
    const state = createState(synthBook({ seed: 1, rows: 340 }).book, { now: 46000, seed: 1 })
    expect(state.dyn).toBeUndefined()
    expect(state.formulas.filter((f) => f && ('reads' in f || 'dynamic' in f)).length).toBe(0)
    const dynamic = createState(book(CHAIN), { now: 46000, seed: 1 })
    expect(dynamic.dyn).toBeDefined()
    expect(dynamic.formulas.filter((f) => f?.reads).map((f) => `${f!.r},${f!.c}`)).toEqual(['0,2', '1,2', '2,2', '0,3', '1,3', '2,3', '3,3'])
  })
})
