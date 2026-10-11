/**
 * A result's `diagnostics`: the cells with an error (`kind: 'cell'`, sheet by sheet and row by row) and then the names
 * with a problem (`kind: 'name'`, in the book's order), each with the fields its kind had in `issues` and `nameIssues`,
 * which a result no longer has; the hashes as before. A cell's conditional formats in `CellView.appearance`. In
 * `compute` and the three kinds of session, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, type BookInput, type ConditionalAppearance } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

/** Errors of three kinds, one coming from another cell, a name with an unknown sheet and one Excel would not take. */
const errors = (): BookInput => ({
  sheets: [{ name: 'S', cells: { A1: { value: 0 }, A2: { formula: '1/A1' }, A3: { formula: 'A2+1' }, B1: { formula: 'SUM(B1' }, C1: { formula: 'gone*2' } } }],
  names: [{ name: 'gone', ref: "'Other'!A1" }, { name: 'Q1', ref: '1' }],
})

/** The same with a value a conditional format shows bold. */
function book(): BookInput {
  const b = errors()
  Object.assign(b.sheets[0], { conditionalFormats: [{ type: 'cellIs', ranges: ['D1:D2'], operator: 'greaterThan', formulas: ['10'], style: { font: { bold: true } } }] })
  b.sheets[0].cells.D1 = { value: 12 }
  return b
}

const DIAGNOSTICS = [
  { kind: 'cell', sheet: 'S', cell: 'B1', error: 'ERROR', reason: 'expected ), found end of formula', check: 'syntax', at: 6 },
  { kind: 'cell', sheet: 'S', cell: 'C1', error: 'REF', reason: 'invalid reference' },
  { kind: 'cell', sheet: 'S', cell: 'A2', error: 'DIV0', reason: 'division by zero' },
  { kind: 'cell', sheet: 'S', cell: 'A3', error: 'DIV0', reason: "division by zero, from 'S'!A2" },
  { kind: 'name', name: 'gone', reason: 'unknown sheet Other', dropped: false },
  { kind: 'name', name: 'Q1', reason: 'a name must not look like a cell address or an R1C1 reference', dropped: true },
]

describe.each(['closures', 'eval'] as const)('diagnostics and appearance, %s', (execution) => {
  it('compute: one list of the cells, then the names, and the hashes as before', () => {
    const r = compute(errors(), { clock: { now: '2026-10-04T12:00:00Z', seed: 7 }, execution })
    expect(r.diagnostics).toEqual(DIAGNOSTICS)
    expect(r).not.toHaveProperty('issues')
    expect(r).not.toHaveProperty('nameIssues')
    // the hashes this book gave when its result had issues and nameIssues
    expect([r.inputHash, r.hash]).toEqual(['3f51848459bc673e6eec004cd6a9b3afb5273d618a120b47d5de7289bd66f4e5', '90b400512a753eeebcee7d71188fc6479ea027369b64bdaf593446d4988a9e96'])
  })

  it.each(KINDS)('%s: the result has diagnostics, a cell its appearance', async (kind) => {
    const s = await (await opener(kind))(book(), { clock: CLOCK, execution })
    expect((await s.getResult()).diagnostics).toEqual(DIAGNOSTICS)
    const appearance: ConditionalAppearance = { style: { font: { bold: true } } }
    expect((await s.getCells('S', 'D1')).D1).toEqual({ value: 12, text: '12', appearance })
    expect((await s.getViewport('S', 'D1')).cells.D1.appearance).toEqual(appearance)
    await s.close()
  })
})
