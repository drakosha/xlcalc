/**
 * Iterative calculation (`<calcPr iterate="1" iterateCount iterateDelta>`): how many times Excel computes a cycle and
 * when it stops (a counter, a converging cycle, a diverging one, two cycles in one book), the order it computes the
 * cells of a cycle in (by position, across sheets, with precedents and dependents), the values it starts from (the
 * file's, the last ones after a recalculation or an edit), a cycle with an error, a cycle through a spill range, a
 * financial model's interest on the average balance.
 */
import type { CellInput, SheetInput } from '../../../src/contract.js'
import type { FileParts } from '../case-book.js'
import type { Step, WorkbookCase } from './cases.js'

type Spec = number | string

const cell = (v: Spec): CellInput => (typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1) } : { value: v })
const sheet = (name: string, cells: Record<string, Spec>): SheetInput => ({ name, cells: Object.fromEntries(Object.entries(cells).map(([a, v]) => [a, cell(v)])) })

/** Read after opening, after a recalculation and after a full calculation. */
const AGAIN: Step[] = ['read', 'calc', 'read', 'full', 'read']

/** Ordinary formulas, as Excel 365 saves the formulas of a cycle. */
const PLAIN = { plain: {} }

/** The value 0 stored for every formula of the sheets, as a file Excel saved stores values: without one Excel starts a cycle from no number (#VALUE!). */
const zeros = (sheets: SheetInput[]) => Object.fromEntries(sheets.flatMap((s) => Object.entries(s.cells).flatMap(([a, c]) => (c.formula === undefined ? [] : [[`${s.name}!${a}`, 0]]))))

function one(id: string, about: string, cells: Record<string, Spec>, range: string, iterate: FileParts['iterate'] = {}, extra: Partial<WorkbookCase> = {}): WorkbookCase {
  const sheets = [sheet('S', cells)]
  return { id, about, book: { sheets }, read: [{ sheet: 'S', range }], file: { iterate, cached: zeros(sheets), ...PLAIN }, steps: AGAIN, ...extra }
}

/** A cycle across the sheets S and T, one iteration. */
const two = (id: string, about: string, sheets: SheetInput[]): WorkbookCase => ({
  id,
  about,
  book: { sheets },
  read: [{ sheet: 'S', range: 'A1' }, { sheet: 'T', range: 'A1' }],
  file: { iterate: { count: 1 }, cached: zeros(sheets), ...PLAIN },
  steps: AGAIN,
})

export function buildWorkbookCases(): WorkbookCase[] {
  return [
    one('counter', 'A1 = A1+1 with dependents right of it and below it', { A1: '=A1+1', B1: '=A1', A3: '=A1*2', B2: '=A1' }, 'A1:B3'),
    one('counter-dynamic', 'the counter written as dynamic array formulas', { A1: '=A1+1', B1: '=A1' }, 'A1:B1', {}, { file: { iterate: {}, cached: { 'S!A1': 0, 'S!B1': 0 } } }),
    one('counter-blank', 'the counter whose file stores no value', { A1: '=A1+1', B1: '=A1' }, 'A1:B1', {}, { file: { iterate: {}, ...PLAIN } }),
    one('counter-typed', 'the counter typed into an empty cell after opening', { B1: '=A1' }, 'A1:B1', {}, { steps: ['read', { formula: { sheet: 'S', cell: 'A1', formula: 'A1+1' } }, 'read', 'calc', 'read'] }),
    one('counter-10', 'the counter with 10 iterations', { A1: '=A1+1', B1: '=A1' }, 'A1:B1', { count: 10 }),
    one('counter-cached', 'the counter whose file stores 50', { A1: '=A1+1', B1: '=A1' }, 'A1:B1', {}, { file: { iterate: {}, cached: { 'S!A1': 50, 'S!B1': 0 }, ...PLAIN } }),
    one('counter-edit', 'the counter, then a value typed into an unrelated cell, then into a cell it reads', { A1: '=A1+C1', C1: 1, B1: '=A1' }, 'A1:C1', {}, {
      steps: ['read', { set: { sheet: 'S', cell: 'D1', value: 5 } }, 'read', { set: { sheet: 'S', cell: 'C1', value: 2 } }, 'read'],
    }),
    one('converge', 'A1 = A1/2+1 and B1 = B1/3+1, two cycles that converge', { A1: '=A1/2+1', B1: '=B1/3+1' }, 'A1:B1'),
    one('converge-delta', 'the same with a maximum change of 0.1', { A1: '=A1/2+1', B1: '=B1/3+1' }, 'A1:B1', { delta: 0.1 }),
    one('converge-counter', 'a converging cycle beside a counter in one book', { A1: '=A1/2+1', B1: '=B1+1' }, 'A1:B1'),
    one('converge-edit', 'A1 = A1/3+C1, then C1 changes from 1 to 2', { A1: '=A1/3+C1', C1: 1 }, 'A1:C1', {}, { steps: ['read', { set: { sheet: 'S', cell: 'C1', value: 2 } }, 'read', 'calc', 'read', 'full', 'read'] }),
    one('diverge', 'A1 = A1*2+1', { A1: '=A1*2+1' }, 'A1'),
    one('order-row', 'A1 = B1+1 and B1 = A1*10, one iteration', { A1: '=B1+1', B1: '=A1*10' }, 'A1:B1', { count: 1 }),
    one('order-row-back', 'B1 = A1+1 and A1 = B1*10, one iteration', { B1: '=A1+1', A1: '=B1*10' }, 'A1:B1', { count: 1 }),
    one('order-column', 'A1 = A2+1 and A2 = A1*10, one iteration', { A1: '=A2+1', A2: '=A1*10' }, 'A1:A2', { count: 1 }),
    one('order-three', 'A1 = C1+1, B1 = A1*10, C1 = B1+100, one iteration and two', { A1: '=C1+1', B1: '=A1*10', C1: '=B1+100' }, 'A1:C1', { count: 1 }),
    one('order-three-2', 'the same with two iterations', { A1: '=C1+1', B1: '=A1*10', C1: '=B1+100' }, 'A1:C1', { count: 2 }),
    one('order-diagonal', 'B1 = A2+1 and A2 = B1*10, one iteration', { B1: '=A2+1', A2: '=B1*10' }, 'A1:B2', { count: 1 }),
    one('order-chain', 'a precedent chain into the cycle and a dependent out of it, one iteration and three', { A1: 5, B1: '=A1*2', C1: '=B1+D1', D1: '=C1/10', E1: '=C1', F1: '=E1+1' }, 'A1:F1', { count: 1 }),
    one('order-chain-3', 'the same with three iterations', { A1: 5, B1: '=A1*2', C1: '=B1+D1', D1: '=C1/10', E1: '=C1', F1: '=E1+1' }, 'A1:F1', { count: 3 }),
    one('order-before', 'a dependent above the cycle and one before it in the row', { A1: '=C3', B3: '=C3', C3: '=C3+1' }, 'A1:C3', { count: 5 }),
    two('order-sheets', 'S!A1 = T!A1+1 and T!A1 = S!A1*10 on two sheets, one iteration', [sheet('S', { A1: '=T!A1+1' }), sheet('T', { A1: '=S!A1*10' })]),
    two('order-sheets-back', 'the same with the sheets the other way round', [sheet('T', { A1: '=S!A1*10' }), sheet('S', { A1: '=T!A1+1' })]),
    one('error', 'A1 = IF(A1>5,NA(),A1+1), B1 = A1+B1, C1 = IFERROR(C1+1/0,C1+2), D1 = 1/0+D1', { A1: '=IF(A1>5,NA(),A1+1)', B1: '=A1+B1', C1: '=IFERROR(C1+1/0,C1+2)', D1: '=1/0+D1' }, 'A1:D1'),
    one('spill', 'A1 = SEQUENCE(3)+B1 spills over A1:A3, B1 = SUM(A2:A3)', { A1: '=SEQUENCE(3)+B1', B1: '=SUM(A2:A3)' }, 'A1:B3', { count: 3 }, { file: { iterate: { count: 3 }, cached: { 'S!A1': 0, 'S!B1': 0 }, plain: { arrays: ['S!A1'] } } }),
    one('interest', 'interest on the average balance: B3 = B2*(B1+B4)/2, B4 = B1+B3-100', { B1: 1000, B2: 0.1, B3: '=B2*(B1+B4)/2', B4: '=B1+B3-100', B5: '=B4-B1' }, 'B1:B5'),
    one('interest-fine', 'the same with a maximum change of 0.000001', { B1: 1000, B2: 0.1, B3: '=B2*(B1+B4)/2', B4: '=B1+B3-100', B5: '=B4-B1' }, 'B1:B5', { delta: 0.000001 }),
  ]
}
