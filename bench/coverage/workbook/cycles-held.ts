/**
 * Held out: circular references without iterative calculation in cases the set `cycles` did not have, their answers
 * predicted by the engine's rules (`docs/cycles.md`) and committed before Excel ran (`cycles-held.predicted.json`): a
 * spilled array that a formula typed elsewhere brings into a cycle's way, a formula of a cycle typed again as it was, a
 * cell of a cycle cleared, a dependent typed so that it no longer reads the cycle, two cycles of which one is broken
 * under a cell that reads both, and a file storing values for some cells of a cycle and none for the others.
 */
import { existsSync, readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { WorkbookCase } from './cases.js'
import { PLAIN, sheet } from './iterate.js'

const file = new URL('./cycles-held.predicted.json', import.meta.url)
const PREDICTED = existsSync(file) ? (JSON.parse(readFileSync(file, 'utf8')) as Record<string, Record<string, Val>[]>) : {}

type Spec = number | string

/** A case with a circular reference on sheet S: its cells, the rectangle read and shown, the values its file stores and its steps. */
function held(id: string, about: string, cells: Record<string, Spec>, range: string, cached: Record<string, number> | undefined, steps: WorkbookCase['steps'], arrays?: string[]): WorkbookCase {
  const read = [{ sheet: 'S', range }]
  const c: WorkbookCase = { id, about, book: { sheets: [sheet('S', cells)] }, read, show: read, file: { ...(cached ? { cached } : {}), ...PLAIN, ...(arrays ? { plain: { arrays } } : {}) }, steps, cycle: true }
  return PREDICTED[id] ? { ...c, predicted: PREDICTED[id] } : c
}

const LOOP = { A1: '=B1/2+1', B1: '=A1', C1: '=A1*10' }
const STORED = { 'S!A1': 2, 'S!B1': 2, 'S!C1': 20 }

export function buildWorkbookCases(): WorkbookCase[] {
  return [
    held('array-enters', 'A1 = SEQUENCE(3)+D1 spilling over A1:A3 with D1 = 0, F1 = SUM(A1:A3); then D1 = A3*0 typed, a cycle through the spill', { A1: '=SEQUENCE(3)+D1', D1: 0, F1: '=SUM(A1:A3)' }, 'A1:F3', undefined, [
      'read',
      { formula: { sheet: 'S', cell: 'D1', formula: 'A3*0' } },
      'read',
      'full',
      'read',
    ], ['S!A1']),
    held('same-again', 'the cycle A1 = B1/2+1, B1 = A1, C1 = A1*10 storing 2, 2 and 20; then B1 = A1 typed again as it was', LOOP, 'A1:C1', STORED, ['read', { formula: { sheet: 'S', cell: 'B1', formula: 'A1' } }, 'read', 'full', 'read']),
    held('cleared', 'the same cycle; then B1 cleared', LOOP, 'A1:C1', STORED, ['read', { set: { sheet: 'S', cell: 'B1', value: null } }, 'read', 'full', 'read']),
    held('out-of-way', 'the same cycle and K1 = 3; then C1 = K1*2 typed, which reads the cycle no more', { ...LOOP, K1: 3 }, 'A1:K1', STORED, ['read', { formula: { sheet: 'S', cell: 'C1', formula: 'K1*2' } }, 'read', 'full', 'read']),
    held('two-cycles', 'A1 = B1/2+1, B1 = A1 storing 2 and 2, D1 = E1+1, E1 = D1 storing 5 and 4, F1 = A1+D1 storing 0; then 4 typed into B1', { A1: '=B1/2+1', B1: '=A1', D1: '=E1+1', E1: '=D1', F1: '=A1+D1' }, 'A1:F1', { 'S!A1': 2, 'S!B1': 2, 'S!D1': 5, 'S!E1': 4, 'S!F1': 0 }, [
      'read',
      { set: { sheet: 'S', cell: 'B1', value: 4 } },
      'read',
      'full',
      'read',
    ]),
    held('partly-stored', 'the cycle of `same-again`, the file storing 2 for A1 and nothing for B1 and C1', LOOP, 'A1:C1', { 'S!A1': 2 }, ['read', 'full', 'read']),
  ]
}
