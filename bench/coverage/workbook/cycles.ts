/**
 * Circular references without iterative calculation, as Excel for Windows takes them (`cycle`: the Mac's step refuses
 * these books, as Excel for Mac shows a warning window on opening). A sheet component's check found that Excel leaves
 * a cycle's cells as they are: a cycle typed into a book (A1 = 2, B1 = A1, then A1 = B1/2+1, then D1 = E1/2+1 and
 * E1 = D1+C1, then C1 changed), a book whose file stores the values of its cycle (converged ones, others, none), a
 * precedent changed, a formula of the cycle typed again, the link broken by a value and typed back, a cycle of one
 * cell, cycles through INDIRECT, OFFSET, a defined name and a spilled array, and the same cycle in a file whose
 * calcPr asks for iterative calculation. Then what the first run left open: which cells beside a cycle Excel computes
 * (a book without a cycle whose file stores a wrong value, a formula that does not read the cycle, a dependent above
 * the cycle, on another sheet, with another precedent changed, typed after opening, a chain typed before the cycle
 * closes), and a book without iterative calculation opened after one with it. Each case reads after opening and after
 * full calculations, edits and recalculations, shows the cells' text, and some save the book (`save`) for the values
 * the file stores.
 */
import type { WorkbookCase } from './cases.js'
import { PLAIN, sheet } from './iterate.js'

type Spec = number | string

/** A case on sheet S: its cells, the rectangle read and shown, the values its file stores and its steps. */
function book(id: string, about: string, cells: Record<string, Spec>, range: string, cached: Record<string, number> | undefined, steps: WorkbookCase['steps']): WorkbookCase {
  const read = [{ sheet: 'S', range }]
  return { id, about, book: { sheets: [sheet('S', cells)] }, read, show: read, file: { ...(cached ? { cached } : {}), ...PLAIN }, steps }
}

/** A case with a circular reference, which Excel for Windows alone runs. */
const cycle = (...a: Parameters<typeof book>): WorkbookCase => ({ ...book(...a), cycle: true })

const LOOP = { A1: '=B1/2+1', B1: '=A1', C1: '=A1*10' }

export function buildWorkbookCases(): WorkbookCase[] {
  return [
    cycle('typed', 'A1 = 2 and B1 = A1, then A1 = B1/2+1 typed; then D1 = E1/2+1 and E1 = D1+C1, then C1 = 10; three full calculations; saved', { A1: 2, B1: '=A1' }, 'A1:E1', undefined, [
      'read',
      { formula: { sheet: 'S', cell: 'A1', formula: 'B1/2+1' } },
      'read',
      { formula: { sheet: 'S', cell: 'D1', formula: 'E1/2+1' } },
      'read',
      { formula: { sheet: 'S', cell: 'E1', formula: 'D1+C1' } },
      'read',
      { set: { sheet: 'S', cell: 'C1', value: 10 } },
      'read',
      'full',
      'full',
      'full',
      'read',
      'save',
    ]),
    cycle('stored', 'A1 = B1/2+1, B1 = A1, C1 = A1*10, the file storing 2, 2 and 0; saved', LOOP, 'A1:C1', { 'S!A1': 2, 'S!B1': 2, 'S!C1': 0 }, ['read', 'full', 'read', 'calc', 'read', 'save']),
    cycle('stored-other', 'the same, the file storing 50, 7 and 0', LOOP, 'A1:C1', { 'S!A1': 50, 'S!B1': 7, 'S!C1': 0 }, ['read', 'full', 'read']),
    cycle('stored-none', 'the same, the file storing no values', LOOP, 'A1:C1', undefined, ['read', 'full', 'read']),
    cycle('precedent', 'A1 = B1/2+C1, B1 = A1, D1 = A1+B1, the file storing 2, 2 and 0; then C1 from 1 to 5', { A1: '=B1/2+C1', B1: '=A1', C1: 1, D1: '=A1+B1' }, 'A1:D1', { 'S!A1': 2, 'S!B1': 2, 'S!D1': 0 }, [
      'read',
      { set: { sheet: 'S', cell: 'C1', value: 5 } },
      'read',
      'full',
      'read',
    ]),
    cycle('typed-again', 'the cycle of `stored`, then B1 = A1*3 typed', LOOP, 'A1:C1', { 'S!A1': 2, 'S!B1': 2, 'S!C1': 20 }, ['read', { formula: { sheet: 'S', cell: 'B1', formula: 'A1*3' } }, 'read', 'full', 'read']),
    cycle('broken', 'the cycle of `stored`, then 4 typed into B1, then B1 = A1 again', LOOP, 'A1:C1', { 'S!A1': 2, 'S!B1': 2, 'S!C1': 20 }, [
      'read',
      { set: { sheet: 'S', cell: 'B1', value: 4 } },
      'read',
      { formula: { sheet: 'S', cell: 'B1', formula: 'A1' } },
      'read',
      'full',
      'read',
    ]),
    cycle('self', 'A1 = A1+1 storing 5, B1 = A1*2; then C1 = C1+1 typed; saved', { A1: '=A1+1', B1: '=A1*2' }, 'A1:C1', { 'S!A1': 5, 'S!B1': 10 }, ['read', 'full', 'read', { formula: { sheet: 'S', cell: 'C1', formula: 'C1+1' } }, 'read', 'save']),
    cycle('indirect', 'A1 = INDIRECT("B1")/2+1, B1 = A1, storing 2 and 2', { A1: '=INDIRECT("B1")/2+1', B1: '=A1' }, 'A1:B1', { 'S!A1': 2, 'S!B1': 2 }, ['read', 'full', 'read']),
    cycle('offset', 'A1 = OFFSET(B1,0,0)/2+1, B1 = A1, storing 2 and 2', { A1: '=OFFSET(B1,0,0)/2+1', B1: '=A1' }, 'A1:B1', { 'S!A1': 2, 'S!B1': 2 }, ['read', 'full', 'read']),
    {
      ...cycle('name', 'A1 = nb/2+1 with the name nb for B1, B1 = A1, storing 2 and 2', { A1: '=nb/2+1', B1: '=A1' }, 'A1:B1', { 'S!A1': 2, 'S!B1': 2 }, ['read', 'full', 'read']),
      book: { sheets: [sheet('S', { A1: '=nb/2+1', B1: '=A1' })], names: [{ name: 'nb', ref: "'S'!$B$1" }] },
    },
    {
      ...cycle('spill', 'A1 = SEQUENCE(3)*B5 spilling over A1:A3, B5 = SUM(A1:A3), storing 0', { A1: '=SEQUENCE(3)*B5', B5: '=SUM(A1:A3)' }, 'A1:B5', undefined, ['read', 'full', 'read']),
      file: { cached: { 'S!A1': 0, 'S!B5': 0 }, plain: { arrays: ['S!A1'] }, mirrors: PLAIN.mirrors },
    },
    // what the first run left open
    book('control', 'no cycle: C1 = A1*10 with A1 = 2, the file storing 0 for C1', { A1: 2, C1: '=A1*10' }, 'A1:C1', { 'S!C1': 0 }, ['read', 'full', 'read']),
    cycle('unrelated', 'the cycle of `stored` and E1 = K1*2 with K1 = 3, the file storing 0 for E1; then K1 = 4', { ...LOOP, E1: '=K1*2', K1: 3 }, 'A1:K1', { 'S!A1': 2, 'S!B1': 2, 'S!C1': 0, 'S!E1': 0 }, [
      'read',
      'full',
      'read',
      { set: { sheet: 'S', cell: 'K1', value: 4 } },
      'read',
    ]),
    cycle('above', 'C3 = D3/2+1, D3 = C3, A1 = C3*10 above them, the file storing 2, 2 and 0', { A1: '=C3*10', C3: '=D3/2+1', D3: '=C3' }, 'A1:D3', { 'S!A1': 0, 'S!C3': 2, 'S!D3': 2 }, ['read', 'full', 'read']),
    {
      ...cycle('sheet', 'the cycle of `stored` on S and T!A1 = S!A1*10, the file storing 0 for it', LOOP, 'A1:C1', { 'S!A1': 2, 'S!B1': 2, 'S!C1': 0, 'T!A1': 0 }, ['read', 'full', 'read']),
      book: { sheets: [sheet('S', LOOP), sheet('T', { A1: '=S!A1*10' })] },
      read: [{ sheet: 'S', range: 'A1:C1' }, { sheet: 'T', range: 'A1' }],
      show: [{ sheet: 'S', range: 'A1:C1' }, { sheet: 'T', range: 'A1' }],
    },
    cycle('other-precedent', 'the cycle of `stored` and C1 = A1*10+K1 with K1 = 1, the file storing 0 for C1; then K1 = 5', { A1: '=B1/2+1', B1: '=A1', C1: '=A1*10+K1', K1: 1 }, 'A1:K1', { 'S!A1': 2, 'S!B1': 2, 'S!C1': 0 }, [
      'read',
      { set: { sheet: 'S', cell: 'K1', value: 5 } },
      'read',
      'full',
      'read',
    ]),
    cycle('new-dependent', 'the cycle of `stored`, then E1 = A1+100 typed', LOOP, 'A1:E1', { 'S!A1': 2, 'S!B1': 2, 'S!C1': 0 }, ['read', { formula: { sheet: 'S', cell: 'E1', formula: 'A1+100' } }, 'read', 'full', 'read']),
    cycle('typed-chain', 'A1 = 2, B1 = A1, C1 = B1*10, then A1 = B1/2+1 typed; then 3 into F1, which nothing reads', { A1: 2, B1: '=A1', C1: '=B1*10' }, 'A1:F1', undefined, [
      'read',
      { formula: { sheet: 'S', cell: 'A1', formula: 'B1/2+1' } },
      'read',
      { set: { sheet: 'S', cell: 'F1', value: 3 } },
      'read',
      'full',
      'read',
    ]),
    // last: a book without iterative calculation opened in the same Excel after one with it
    {
      ...cycle('iterate-file', 'the cycle of `stored` in a file whose calcPr asks for iterative calculation', LOOP, 'A1:C1', undefined, ['read', 'full', 'read']),
      file: { iterate: {}, cached: { 'S!A1': 2, 'S!B1': 2, 'S!C1': 0 }, ...PLAIN },
    },
    cycle('after-iterate', 'the cycle of `stored`, its file without iterative calculation, opened after `iterate-file` in the same Excel', LOOP, 'A1:C1', { 'S!A1': 2, 'S!B1': 2, 'S!C1': 0 }, ['read', 'full', 'read']),
  ]
}
