/**
 * Held out after `spill-cycles-held`: two more books of arrays in a cycle through their own cells under iterative
 * calculation, where a session and a full computation of the engine part. `A4 = UNIQUE(A4#)` read by `B2 = A4#+1`, then
 * `IF(A5:A7>1,D2:D3,365)` typed into A4, whose spill A4:A6 covers the cells it reads (`unique`, alone and among the
 * cells of the book the random check found); `C3 = SUM(D1:D3)` as an array of the old kind over C3:D3, which takes in
 * D3, beside `D2 = 2` (`sum`), from the file and entered cell by cell in both orders. Each read after typing, after a
 * full calculation and after one that rebuilds the dependencies. Predicted by the engine before Excel ran it
 * (`spill-cycles-held-2.predicted.json`; a case the engine does not compute is predicted as its error).
 */
import { existsSync, readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { StoredCell } from '../stored.js'
import type { Step, WorkbookCase } from './cases.js'

const FILE = new URL('./spill-cycles-held-2.predicted.json', import.meta.url)
const PREDICTED = (existsSync(FILE) ? JSON.parse(readFileSync(FILE, 'utf8')) : {}) as Record<string, Record<string, Val>[]>

/** Read after typing, after a full calculation, after one that rebuilds the dependencies. */
const AFTER: Step[] = ['read', 'full', 'read', 'rebuild', 'read']

/** A case of sheet S: its cells in the file, the steps before the reads, iterative calculation with `iterate`. */
function held(id: string, about: string, cells: Record<string, StoredCell>, steps: Step[], range: string, iterate: { count?: number; delta?: number } = {}): WorkbookCase {
  const c: WorkbookCase = { id, about, book: { sheets: [{ name: 'S', cells }] }, read: [{ sheet: 'S', range }], steps: [...steps, ...AFTER], file: { iterate, mirrors: ['zMirrorT', 'zMirrorN', 'zMirrorX'] } }
  return PREDICTED[id] ? { ...c, predicted: PREDICTED[id] } : c
}

const typed = (cell: string, formula: string, array?: string): Step => ({ formula: { sheet: 'S', cell, formula, ...(array ? { array } : { dynamic: true }) } })

const IF = 'IF(A5:A7>1,D2:D3,365)'
const UNIQUE = { A4: { formula: 'UNIQUE(A4#)' }, B2: { formula: 'A4#+1' } }
const FOUND = { ...UNIQUE, B1: { formula: 'COUNT(E4:E4)' }, C7: { formula: '{3,-1;0.5,7}' }, B5: { value: 0 }, D5: { value: 3 } }
const SUM = { formula: 'SUM(D1:D3)', array: 'C3:D3' }

export function buildWorkbookCases(): WorkbookCase[] {
  return [
    held('unique-typed', `A4 = UNIQUE(A4#) and B2 = A4#+1, then ${IF} typed into A4`, UNIQUE, [typed('A4', IF)], 'A1:E8'),
    held('unique-typed-20', `the same with 20 iterations and no change to stop at`, UNIQUE, [typed('A4', IF)], 'A1:E8', { count: 20, delta: 0 }),
    held('unique-found-typed', `the book the random check found, then ${IF} typed into A4, 20 iterations`, FOUND, [typed('A4', IF)], 'A1:E8', { count: 20, delta: 0 }),
    held('unique-file', `${IF} in A4 and B2 = A4#+1 in the file`, { A4: { formula: IF }, B2: UNIQUE.B2 }, [], 'A1:E8'),
    held('sum-file', 'C3:D3 = SUM(D1:D3) as an array of the old kind and D2 = 2 in the file', { C3: SUM, D2: { value: 2 } }, [], 'A1:E4'),
    held('sum-typed-array-last', 'D2 = 2 in the file, then the array typed over C3:D3', { D2: { value: 2 } }, [typed('C3', SUM.formula, SUM.array)], 'A1:E4'),
    held('sum-typed-value-last', 'the array in the file, then 2 typed into D2', { C3: SUM }, [{ set: { sheet: 'S', cell: 'D2', value: 2 } }], 'A1:E4'),
  ]
}
