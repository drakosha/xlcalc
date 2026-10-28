/**
 * What functions keep with a long range's values between calls (holidays, number lists, sorted copies) and the
 * range summaries SUBTOTAL and AGGREGATE take: after an edit inside the range, a clock change that moves date text
 * without a year or a row hidden or shown, a session answers as a fresh computation of the same book.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput, type CellEdit, type Clock } from '../src/index.js'

type Cells = BookInput['sheets'][number]['cells']
const clock: Clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** A row (1-based) hidden or shown. */
interface Row { row: string; hidden: boolean }
type Sheet = Omit<BookInput['sheets'][number], 'name' | 'cells'>

/** The session's values of `range` after each edit (clock, row), against a fresh computation of the edited book. */
function follows(cells: Cells, range: string, steps: (CellEdit | Clock | Row)[], execution: 'closures' | 'eval', sheet: Sheet = {}) {
  const s = createSession({ sheets: [{ name: 'S', cells: { ...cells }, ...sheet }] }, { clock, execution })
  let now = clock
  const book = { ...cells }
  const rows = { ...sheet.rows }
  for (const step of steps) {
    if ('now' in step) s.setClock((now = step))
    else if ('row' in step) {
      s.setLayout([{ sheet: 'S', rows: step.row, set: { hidden: step.hidden } }])
      rows[step.row] = { hidden: step.hidden }
    } else {
      s.setCells([step])
      book[step.cell] = step.input as CellInput
    }
    const fresh = compute({ sheets: [{ name: 'S', cells: book, ...sheet, rows }] }, { clock: now, execution }).sheets[0].cells
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

  it.each([false, true])('SUBTOTAL and AGGREGATE over running ranges follow hidden rows and SUBTOTAL inside them, filtered %s', (filtered) => {
    const codes = ['9', '109', '1', '103', '4', '105']
    const cells = column(300, (r) => ({ value: (r * 37) % 101 }), (r) => (r % 3 ? `SUBTOTAL(${codes[r % 6]},$H$1:H${r})` : `AGGREGATE(${['9', '2', '5'][r % 3]},${r % 2 ? 1 : 4},$H$1:H${r})`))
    follows(cells, 'B1:B300', [
      { row: '10', hidden: true },
      { sheet: 'S', cell: 'H20', input: { formula: 'SUBTOTAL(9,H1:H3)' } },
      { sheet: 'S', cell: 'H30', input: { formula: '1/0' } },
      { row: '10', hidden: false },
      { sheet: 'S', cell: 'H20', input: { value: 5 } },
      { sheet: 'S', cell: 'H30', input: { value: 'x' } },
      { row: '250', hidden: true },
    ], execution, { filtered })
  })

  it('SUBTOTAL and AGGREGATE over a long range skip SUBTOTAL inside it and the hidden rows their code skips', () => {
    const cells: Cells = {}
    for (let r = 1; r <= 400; r++) cells[`H${r}`] = { value: r }
    cells.H5 = { formula: 'SUBTOTAL(9,J1:J2)' }
    cells.J1 = { value: 1000 }
    const sums = ['SUBTOTAL(9,H1:H400)', 'SUBTOTAL(109,H1:H400)', 'AGGREGATE(9,4,H1:H400)', 'AGGREGATE(9,5,H1:H400)', 'SUBTOTAL(2,H1:H400)', 'SUM(H1:H400)']
    sums.forEach((f, i) => (cells[`B${i + 1}`] = { formula: f }))
    const all = (400 * 401) / 2 - 5
    const got = compute({ sheets: [{ name: 'S', cells, rows: { 7: { hidden: true } } }] }, { clock, execution }).sheets[0].cells
    expect(sums.map((_, i) => got[`B${i + 1}`].value)).toEqual([all, all - 7, all + 1000, all + 1000 - 7, 399, all + 1000])
  })
})
