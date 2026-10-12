/**
 * Why a cell holds an error, without the book's result: `Session.getCellIssue` gives each cell's issue as
 * `getResult().diagnostics` does and the cell the error began in, through ranges, spills, sheets and cycles; `explain` marks
 * each error with that cell and leads to it, through a range by the cell the error came from and past its depth, as far
 * as `errorDepth` and never round a cycle. In a synchronous session, in the same thread and in a thread, with closures
 * and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { InputError, createSession, type BookInput, type Execution } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import './locales.js'
import { cellIssues } from './diagnostics.js'

afterAll(closeCalcs)

/** Errors of every kind: their own, from a cell, through a range and a whole column, in spills, a cycle, a check, another sheet, a chain. */
const errors = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { formula: '1/0' }, A2: { formula: 'A1*2' }, A3: { formula: 'SUM(A1:A2)' }, B1: { formula: 'SUM(A:A)' },
        C1: { formula: 'SEQUENCE(3)/{1;0;1}' }, D1: { formula: 'A1:A3' },
        E1: { formula: 'E2+1' }, E2: { formula: 'E1+1' }, E3: { formula: 'E1*2' },
        F1: { formula: 'SUM(' }, G1: { formula: 'SEQUENCE(2)' }, G2: { value: 'x' }, H1: { formula: 'NA()' }, I1: { formula: 'Other!A1+1' },
        ...Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`J${i + 1}`, { formula: i ? `J${i}+1` : '1/0' }])),
        K1: { value: 5 },
      },
    },
    { name: 'Other', cells: { A1: { formula: '1/0' } } },
  ],
})

const ORIGINS: Record<string, string> = {
  'S!A1': 'S!A1', 'S!A2': 'S!A1', 'S!A3': 'S!A1', 'S!B1': 'S!A1', 'S!C2': 'S!C1', 'S!D1': 'S!A1', 'S!D2': 'S!A1', 'S!D3': 'S!A1',
  // each cell of a cycle is where its error begins; a cell after the cycle, the cell of the cycle it reaches first
  'S!E1': 'S!E1', 'S!E2': 'S!E2', 'S!E3': 'S!E1',
  'S!F1': 'S!F1', 'S!G1': 'S!G1', 'S!H1': 'S!H1', 'S!I1': 'Other!A1', 'Other!A1': 'Other!A1',
  ...Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`S!J${i + 1}`, 'S!J1'])),
}

describe.each(['closures', 'eval'] as const)('the reason and origin of an error, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it("each cell's issue as the result's, with the cell its error began in; null without an error", async () => {
      const s = await (await opener(kind))(errors(), { clock: CLOCK, execution })
      const issues = cellIssues(await s.getResult())
      expect(issues.map((i) => `${i.sheet}!${i.cell}`).sort()).toEqual(Object.keys(ORIGINS).sort())
      for (const { sheet, cell, ...rest } of issues) {
        const got = await s.getCellIssue(sheet, cell)
        expect(got, `${sheet}!${cell}`).toEqual({ ...rest, origin: { sheet: ORIGINS[`${sheet}!${cell}`].split('!')[0], cell: ORIGINS[`${sheet}!${cell}`].split('!')[1] } })
      }
      expect(await s.getCellIssue('S', 'F1')).toEqual({ error: 'ERROR', reason: 'unexpected end of formula', check: 'syntax', at: 4, origin: { sheet: 'S', cell: 'F1' } })
      for (const cell of ['K1', 'G2', 'C1', 'Z9']) expect(await s.getCellIssue('S', cell), cell).toBeNull()
      // an edit: the error goes, and comes from elsewhere
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 1 } }, { sheet: 'S', cell: 'A2', input: { formula: 'H1' } }])
      expect(await s.getCellIssue('S', 'A1')).toBeNull()
      expect(await s.getCellIssue('S', 'B1')).toEqual({ error: 'NA', reason: "value not found, from 'S'!A2", origin: { sheet: 'S', cell: 'H1' } })
      await s.close()
    })

    it("explain leads to the origin through a range, past its depth and as far as errorDepth, and stops at a cycle", async () => {
      const s = await (await opener(kind))(errors(), { clock: CLOCK, execution })
      const a3 = await s.explain('S', 'A3')
      expect(a3.text).toBe(["S!A3 = #DIV/0!  ← SUM(A1:A2)  (error from S!A1)", "  S!A1:A2 (2 cells)", "    S!A1 = #DIV/0!  ← 1/0  (the error begins here)"].join('\n'))
      expect(a3.root).toEqual({
        ref: "'S'!A3", value: { error: 'DIV0' }, formula: 'SUM(A1:A2)', origin: "'S'!A1",
        children: [{ ref: "'S'!A1:A2", cells: 2, children: [{ ref: "'S'!A1", value: { error: 'DIV0' }, formula: '1/0', origin: "'S'!A1", children: [] }] }],
      })
      expect((await s.explain('S', 'D2')).text).toBe([
        "S!D2 = #DIV/0!  ← spilled from S!D1  (error from S!A1)",
        "  S!D1 = #DIV/0!  ← A1:A3  (error from S!A1)",
        "    S!A1:A3 (3 cells)",
        "      S!A1 = #DIV/0!  ← 1/0  (the error begins here)",
      ].join('\n'))
      // the cells of a range are all there when it is expanded
      expect((await s.explain('S', 'A3', { expand: true })).root.children[0].children.map((n) => n.ref)).toEqual(["'S'!A1", "'S'!A2"])
      const far = await s.explain('S', 'J10', { depth: 2 })
      expect([far.nodes, far.truncated, far.text.split('\n').at(-1)]).toEqual([10, false, "                  S!J1 = #DIV/0!  ← 1/0  (the error begins here)"])
      const short = await s.explain('S', 'J10', { depth: 2, errorDepth: 3 })
      expect([short.nodes, short.truncated, short.text.split('\n').at(-1)]).toEqual([6, true, "          S!J5 = #DIV/0!  ← J4+1  (error from S!J1)"])
      expect((await s.explain('S', 'J10', { depth: 1, maxNodes: 4 })).nodes).toBe(4)
      const cycle = await s.explain('S', 'E3', { depth: 0, errorDepth: 1000 })
      expect([cycle.nodes, cycle.text]).toEqual([3, ["S!E3 = 0  ← E1*2  (held: a circular reference through S!E1)", "  S!E1 = 0  ← E2+1  (held: a circular reference)", "    S!E2 = 0  ← E1+1  (held: a circular reference)"].join('\n')])
      expect((await s.explain('S', 'I1')).text).toBe("S!I1 = #DIV/0!  ← Other!A1+1  (error from Other!A1)\n  Other!A1 = #DIV/0!  ← 1/0  (the error begins here)")
      expect((await s.explain('S', 'C2')).text).toBe("S!C2 = #DIV/0!  ← spilled from S!C1  (error from S!C1)\n  S!C1 = 1  ← SEQUENCE(3)/{1;0;1}")
      // an empty cell, and errors as cells show them
      expect((await s.explain('S', 'Z9')).text).toBe("S!Z9 = empty")
      expect((await s.explain('S', 'G1')).text).toBe("S!G1 = #SPILL!  ← SEQUENCE(2)  (the error begins here)")
      await s.close()
    })
  })
})

describe('the reason and origin of an error', () => {
  it("in ru-RU, explain writes errors as the book's cells show them", () => {
    const s = createSession({ ...errors(), locale: 'ru-RU' }, { clock: CLOCK })
    expect(s.explain('S', 'A2').text).toBe("S!A2 = #ДЕЛ/0!  ← A1*2  (ошибка из S!A1)\n  S!A1 = #ДЕЛ/0!  ← 1/0  (ошибка возникает здесь)")
    expect(s.getViewport('S', 'A2').cells.A2.text).toBe('#ДЕЛ/0!')
  })

  it('the sheet and the cell are checked as input', () => {
    const s = createSession(errors(), { clock: CLOCK })
    expect(() => s.getCellIssue('Nope', 'A1')).toThrow(InputError)
    expect(() => s.getCellIssue('S', 'a1')).toThrow(InputError)
    expect(s.getCellIssue('other', 'A1')?.origin).toEqual({ sheet: 'Other', cell: 'A1' })
  })
})
