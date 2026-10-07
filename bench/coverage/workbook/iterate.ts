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

/** The harness's mirror sheets named to sort last (`FileParts.mirrors`). */
const LAST = { mirrors: ['zMirrorT', 'zMirrorN', 'zMirrorX'] as const }

/** Ordinary formulas, as Excel 365 saves the formulas of a cycle. */
const PLAIN = { plain: {}, ...LAST }

/** The value 0 stored for every formula of the sheets, as a file Excel saved stores values: without one Excel starts a cycle from no number (#VALUE!). */
const zeros = (sheets: SheetInput[]) => Object.fromEntries(sheets.flatMap((s) => Object.entries(s.cells).flatMap(([a, c]) => (c.formula === undefined ? [] : [[`${s.name}!${a}`, 0]]))))

function one(id: string, about: string, cells: Record<string, Spec>, range: string, iterate: FileParts['iterate'] = {}, extra: Partial<WorkbookCase> = {}): WorkbookCase {
  const sheets = [sheet('S', cells)]
  return { id, about, book: { sheets }, read: [{ sheet: 'S', range }], file: { iterate, cached: zeros(sheets), ...PLAIN }, steps: AGAIN, ...extra }
}

/** A cycle across sheets (S and T unless `read` names them), one iteration; A1 of each sheet read. */
const two = (id: string, about: string, sheets: SheetInput[], read = ['S', 'T']): WorkbookCase => ({
  id,
  about,
  book: { sheets },
  read: read.map((s) => ({ sheet: s, range: 'A1' })),
  file: { iterate: { count: 1 }, cached: zeros(sheets), ...PLAIN },
  steps: AGAIN,
})

export function buildWorkbookCases(): WorkbookCase[] {
  return [
    one('counter', 'A1 = A1+1 with dependents right of it and below it', { A1: '=A1+1', B1: '=A1', A3: '=A1*2', B2: '=A1' }, 'A1:B3'),
    one('counter-dynamic', 'the counter written as dynamic array formulas', { A1: '=A1+1', B1: '=A1' }, 'A1:B1', {}, { file: { iterate: {}, cached: { 'S!A1': 0, 'S!B1': 0 }, ...LAST } }),
    one('counter-blank', 'the counter whose file stores no value', { A1: '=A1+1', B1: '=A1' }, 'A1:B1', {}, { file: { iterate: {}, ...PLAIN } }),
    one('counter-typed', 'the counter typed into an empty cell after opening', { B1: '=A1' }, 'A1:B1', {}, { steps: ['read', { formula: { sheet: 'S', cell: 'A1', formula: 'A1+1' } }, 'read', 'calc', 'read'] }),
    one('text-typed', 'A1 = A1&"x" typed into an empty cell, three iterations', { B1: '=LEN(A1)' }, 'A1:B1', { count: 3 }, { steps: ['read', { formula: { sheet: 'S', cell: 'A1', formula: 'A1&"x"' } }, 'read'] }),
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
    two('sheets-named', 'the cycle across sheets named B and A, in that order', [sheet('B', { A1: '=A!A1*10' }), sheet('A', { A1: '=B!A1+1' })], ['A', 'B']),
    two('sheets-named-back', 'the same with the sheets in the order A, B', [sheet('A', { A1: '=B!A1+1' }), sheet('B', { A1: '=A!A1*10' })], ['A', 'B']),
    {
      ...two('sheets-three', 'a cycle across X and Z with a dependent on Y between them, the sheets in the order Z, Y, X', [sheet('Z', { A1: '=X!A1*10' }), sheet('Y', { A1: '=X!A1+Z!A1' }), sheet('X', { A1: '=Z!A1+1' })], ['X', 'Y', 'Z']),
      file: { iterate: { count: 2 }, cached: { 'Z!A1': 0, 'Y!A1': 0, 'X!A1': 0 }, ...PLAIN },
    },
    {
      ...two('sheets-three-read', 'the same read Z, Y, X', [sheet('Z', { A1: '=X!A1*10' }), sheet('Y', { A1: '=X!A1+Z!A1' }), sheet('X', { A1: '=Z!A1+1' })], ['Z', 'Y', 'X']),
      file: { iterate: { count: 2 }, cached: { 'Z!A1': 0, 'Y!A1': 0, 'X!A1': 0 }, ...PLAIN },
    },
    two('sheets-named-read', 'the cycle across B and A, in that order, read B, A', [sheet('B', { A1: '=A!A1*10' }), sheet('A', { A1: '=B!A1+1' })], ['B', 'A']),
    {
      ...one('order-row-read', 'A1 = B1+1 and B1 = A1*10, one iteration, read B1 then A1', { A1: '=B1+1', B1: '=A1*10' }, 'B1', { count: 1 }),
      read: [{ sheet: 'S', range: 'B1' }, { sheet: 'S', range: 'A1' }],
    },
    {
      ...one('order-diagonal-read', 'B1 = A2+1 and A2 = B1*10, one iteration, read A2 then B1', { B1: '=A2+1', A2: '=B1*10' }, 'A2', { count: 1 }),
      read: [{ sheet: 'S', range: 'A2' }, { sheet: 'S', range: 'B1' }],
    },
    {
      ...two('dependent-sheet-before', 'a dependent on a sheet before the cycle\'s sheet, five iterations', [sheet('D', { A1: '=S!A1' }), sheet('S', { A1: '=S!A1+1' })], ['D', 'S']),
      file: { iterate: { count: 5 }, cached: { 'D!A1': 0, 'S!A1': 0 }, ...PLAIN },
    },
    one('precedent-after', 'a precedent of the cycle placed after it: C1 = H1+D1, D1 = C1/10, H1 = A1*2, one iteration', { A1: 5, C1: '=H1+D1', D1: '=C1/10', H1: '=A1*2' }, 'A1:H1', { count: 1 }),
    one('dependent-chain', 'a dependent before the cycle and a dependent of it after: A1 = C1, C1 = C1+1, D1 = A1, five iterations', { A1: '=C1', C1: '=C1+1', D1: '=A1' }, 'A1:D1', { count: 5 }),
    one('delta-dependent', 'a converging cycle and a dependent that magnifies its changes: B1 = A1*1000', { A1: '=A1/2+1', B1: '=A1*1000' }, 'A1:B1'),
    one('delta-cycle', 'the same magnified inside the cycle: A1 = B1/2000+1, B1 = A1*1000', { A1: '=B1/2000+1', B1: '=A1*1000' }, 'A1:B1'),
    two('sheets-case', 'the cycle across sheets named Zeta and alpha, in that order', [sheet('Zeta', { A1: '=alpha!A1*10' }), sheet('alpha', { A1: '=Zeta!A1+1' })], ['alpha', 'Zeta']),
    two('sheets-digits', 'the cycle across sheets named S2 and S10, in that order', [sheet('S2', { A1: '=S10!A1*10' }), sheet('S10', { A1: '=S2!A1+1' })], ['S10', 'S2']),
    {
      ...two('sheets-five', 'a counter across X and Y, five iterations', [sheet('X', { A1: '=Y!A1+1' }), sheet('Y', { A1: '=X!A1' })], ['X', 'Y']),
      file: { iterate: { count: 5 }, cached: { 'X!A1': 0, 'Y!A1': 0 }, ...PLAIN },
    },
    ...[1, 2, 3, 4, 10, 100].map((n) => ({
      ...two(`xs-counter-${n}`, `a counter across X and Y, ${n} iterations`, [sheet('X', { A1: '=Y!A1+1' }), sheet('Y', { A1: '=X!A1' })], ['X', 'Y']),
      file: { iterate: { count: n }, cached: { 'X!A1': 0, 'Y!A1': 0 }, ...PLAIN },
    })),
    {
      ...two('xs-times', 'a counter across X and Y with Y = X*1, one iteration', [sheet('X', { A1: '=Y!A1+1' }), sheet('Y', { A1: '=X!A1*1' })], ['X', 'Y']),
      file: { iterate: { count: 1 }, cached: { 'X!A1': 0, 'Y!A1': 0 }, ...PLAIN },
    },
    {
      ...two('xs-st', 'a counter across S and T, one iteration', [sheet('S', { A1: '=T!A1+1' }), sheet('T', { A1: '=S!A1' })]),
      file: { iterate: { count: 1 }, cached: { 'S!A1': 0, 'T!A1': 0 }, ...PLAIN },
    },
    {
      ...two('xs-st-times', 'S!A1 = T!A1+1, T!A1 = S!A1*10, three iterations', [sheet('S', { A1: '=T!A1+1' }), sheet('T', { A1: '=S!A1*10' })]),
      file: { iterate: { count: 3 }, cached: { 'S!A1': 0, 'T!A1': 0 }, ...PLAIN },
    },
    {
      ...two('xs-ab-5', 'a counter across A and B, five iterations', [sheet('A', { A1: '=B!A1+1' }), sheet('B', { A1: '=A!A1' })], ['A', 'B']),
      file: { iterate: { count: 5 }, cached: { 'A!A1': 0, 'B!A1': 0 }, ...PLAIN },
    },
    {
      ...two('one-x-5', 'X!A1 = X!A1+1 alone on sheet X, five iterations', [sheet('X', { A1: '=A1+1' })], ['X']),
      file: { iterate: { count: 5 }, cached: { 'X!A1': 0 }, ...PLAIN },
    },
    {
      ...two('one-v-5', 'V!A1 = V!A1+1 alone on sheet V (before the harness\'s sheets WN, WT, WX), five iterations', [sheet('V', { A1: '=A1+1' })], ['V']),
      file: { iterate: { count: 5 }, cached: { 'V!A1': 0 }, ...PLAIN },
    },
    {
      ...two('xs-three', 'a counter across X, Y and Z, six iterations', [sheet('X', { A1: '=Z!A1+1' }), sheet('Y', { A1: '=X!A1' }), sheet('Z', { A1: '=Y!A1' })], ['X', 'Y', 'Z']),
      file: { iterate: { count: 6 }, cached: { 'X!A1': 0, 'Y!A1': 0, 'Z!A1': 0 }, ...PLAIN },
    },
    {
      ...two('xs-converge', 'a converging cycle across X and Y: X = Y/2+1, Y = X', [sheet('X', { A1: '=Y!A1/2+1' }), sheet('Y', { A1: '=X!A1' })], ['X', 'Y']),
      file: { iterate: {}, cached: { 'X!A1': 0, 'Y!A1': 0 }, ...PLAIN },
    },
    {
      ...two('xs-through', 'a counter on S through a cell of T: S!A1 = S!B1+1, S!B1 = T!A1, T!A1 = S!A1, five iterations', [sheet('S', { A1: '=B1+1', B1: '=T!A1' }), sheet('T', { A1: '=S!A1' })], ['S', 'T']),
      read: [{ sheet: 'S', range: 'A1:B1' }, { sheet: 'T', range: 'A1' }],
      file: { iterate: { count: 5 }, cached: { 'S!A1': 0, 'S!B1': 0, 'T!A1': 0 }, ...PLAIN },
    },
    {
      ...two('xs-one-sheet-two', 'two cycles, one on X (X!A1 = X!A1+1) and one across X and Y, five iterations', [sheet('X', { A1: '=A1+1', B1: '=Y!A1+1' }), sheet('Y', { A1: '=X!B1' })], ['X', 'Y']),
      read: [{ sheet: 'X', range: 'A1:B1' }, { sheet: 'Y', range: 'A1' }],
      file: { iterate: { count: 5 }, cached: { 'X!A1': 0, 'X!B1': 0, 'Y!A1': 0 }, ...PLAIN },
    },
    one('text', 'A1 = A1&"x" alone, B1 its length', { A1: '=A1&"x"', B1: '=LEN(A1)' }, 'A1:B1'),
    one('boolean', 'A1 = NOT(A1) alone', { A1: '=NOT(A1)' }, 'A1'),
    one('error-flip', 'A1 = IF(ISERROR(A1),1,IF(A1=1,NA(),1)) alone, B1 = ISNA(A1)', { A1: '=IF(ISERROR(A1),1,IF(A1=1,NA(),1))', B1: '=ISNA(A1)' }, 'A1:B1'),
    one('delta-equal', 'A1 = A1/2+1 with a maximum change of 0.5, which its second change equals', { A1: '=A1/2+1' }, 'A1', { delta: 0.5 }),
    one('type-late', 'A1 goes 1, 2, 3, then the text "x", then "y": does a number turned text count as a change', { A1: '=IF(ISNUMBER(A1),IF(A1<3,A1+1,"x"),"y")' }, 'A1'),
    one('error-late', 'A1 goes 1, 2, 3, then #N/A, then #DIV/0!', { A1: '=IF(ISNUMBER(A1),IF(A1<3,A1+1,NA()),1/0)' }, 'A1'),
    one('bool-late', 'A1 goes 1, 2, 3, then TRUE, then FALSE', { A1: '=IF(ISNUMBER(A1),IF(A1<3,A1+1,TRUE),FALSE)' }, 'A1'),
    one('text-alone', 'A1 = A1&"x" alone, five iterations', { A1: '=A1&"x"' }, 'A1', { count: 5 }),
    one('text-number', 'A1 = A1&"x" and B1 = B1+1, five iterations', { A1: '=A1&"x"', B1: '=B1+1' }, 'A1:B1', { count: 5 }),
    one('text-steady', 'a converging cycle beside a cycle holding the same text from its first iteration', { A1: '=A1/2+1', B1: '=IF(B1=0,"x","x")' }, 'A1:B1'),
    one('error-steady', 'a converging cycle beside a cycle holding the same error from its first iteration', { A1: '=A1/2+1', B1: '=IF(ISERROR(B1),NA(),NA())' }, 'A1:B1'),
    one('error', 'A1 = IF(A1>5,NA(),A1+1), B1 = A1+B1, C1 = IFERROR(C1+1/0,C1+2), D1 = 1/0+D1', { A1: '=IF(A1>5,NA(),A1+1)', B1: '=A1+B1', C1: '=IFERROR(C1+1/0,C1+2)', D1: '=1/0+D1' }, 'A1:D1'),
    one('spill', 'A1 = SEQUENCE(3)+B1 spills over A1:A3, B1 = SUM(A2:A3)', { A1: '=SEQUENCE(3)+B1', B1: '=SUM(A2:A3)' }, 'A1:B3', { count: 3 }, { file: { iterate: { count: 3 }, cached: { 'S!A1': 0, 'S!B1': 0 }, plain: { arrays: ['S!A1'] }, ...LAST } }),
    one('interest', 'interest on the average balance: B3 = B2*(B1+B4)/2, B4 = B1+B3-100', { B1: 1000, B2: 0.1, B3: '=B2*(B1+B4)/2', B4: '=B1+B3-100', B5: '=B4-B1' }, 'B1:B5'),
    one('interest-fine', 'the same with a maximum change of 0.000001', { B1: 1000, B2: 0.1, B3: '=B2*(B1+B4)/2', B4: '=B1+B3-100', B5: '=B4-B1' }, 'B1:B5', { delta: 0.000001 }),
  ]
}
