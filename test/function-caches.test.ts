/**
 * What functions keep with a long range's values between calls (holidays, number lists, sorted copies): after an
 * edit inside the range or a clock change that moves date text without a year, a session answers as a fresh
 * computation of the same book.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput, type CellEdit, type Clock } from '../src/index.js'

type Cells = BookInput['sheets'][number]['cells']
const clock: Clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** The session's values of `range` after each edit (or clock), against a fresh computation of the edited book. */
function follows(cells: Cells, range: string, steps: (CellEdit | Clock)[], execution: 'closures' | 'eval') {
  const s = createSession({ sheets: [{ name: 'S', cells: { ...cells } }] }, { clock, execution })
  let now = clock
  const book = { ...cells }
  for (const step of steps) {
    if ('now' in step) s.setClock((now = step))
    else {
      s.setCells([step])
      book[step.cell] = step.input as CellInput
    }
    const fresh = compute({ sheets: [{ name: 'S', cells: book }] }, { clock: now, execution }).sheets[0].cells
    const got = s.getCells('S', range)
    for (const [a, v] of Object.entries(got)) expect([a, v.value], JSON.stringify(step)).toEqual([a, fresh[a]?.value ?? null])
  }
  s.close()
}

const column = (n: number, value: (r: number) => CellInput, formula: (r: number) => string): Cells => {
  const c: Cells = {}
  for (let r = 1; r <= n; r++) {
    c[`H${r}`] = value(r)
    c[`B${r}`] = { formula: formula(r) }
  }
  return c
}
type CellInput = Cells[string]

describe.each(['closures', 'eval'] as const)('function caches, %s', (execution) => {
  it('WORKDAY and NETWORKDAYS over a long holidays range follow edits in it and the clock year of date text', () => {
    const cells = column(300, (r) => ({ value: 45000 + r * 3 }), (r) => (r % 2 ? `WORKDAY(${45000 + r},${r % 40},$H$1:$H$300)` : `NETWORKDAYS(${45000 + r},${45060 + r},$H$1:$H$300)`))
    // March 20 in the clock's year: 2023's is a Monday among the dates counted
    cells.H7 = { value: '3/20' }
    follows(cells, 'B1:B300', [
      { sheet: 'S', cell: 'H2', input: { value: 45004 } },
      { sheet: 'S', cell: 'H3', input: { value: true } },
      { sheet: 'S', cell: 'H3', input: { value: 45010 } },
      { now: '2023-02-01T00:00:00Z', seed: 7 },
      { sheet: 'S', cell: 'H9', input: { value: 45006 } },
      { now: '2026-02-01T00:00:00Z', seed: 7 },
      { sheet: 'S', cell: 'H7', input: { value: 45031 } },
    ], execution)
  })

  it('RANK.EQ and RANK.AVG over a long range follow edits in it: numbers, text skipped, the first error', () => {
    const cells = column(300, (r) => ({ value: (r * 37) % 101 }), (r) => `RANK.${r % 2 ? 'EQ' : 'AVG'}(H${r},$H$1:$H$300${r % 3 ? '' : ',1'})`)
    follows(cells, 'B1:B300', [
      { sheet: 'S', cell: 'H5', input: { value: 1000 } },
      { sheet: 'S', cell: 'H6', input: { value: 'x' } },
      { sheet: 'S', cell: 'H9', input: { formula: '1/0' } },
      { sheet: 'S', cell: 'H9', input: { value: -0 } },
      { sheet: 'S', cell: 'H200', input: { value: 0 } },
    ], execution)
  })
})
