/**
 * Held out, Excel for Windows alone: the books of `spill-cycles-held` and `spill-cycles-held-2` (arrays in a cycle
 * through their own spill or cells) without iterative calculation, which Excel for Mac opens with a warning window.
 * Each book from its file and its formulas typed one by one in two orders, then read after typing, after a full
 * calculation and after one that rebuilds the dependencies. The ids are those of the sets with iterative calculation,
 * and `unique` adds both of its formulas typed into the empty sheet. Predicted by the engine before Excel ran it
 * (`spill-cycles-off-held.predicted.json`; a case the engine does not compute is predicted as its error).
 */
import { existsSync, readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Step, WorkbookCase } from './cases.js'
import { buildWorkbookCases as held } from './spill-cycles-held.js'
import { buildWorkbookCases as held2 } from './spill-cycles-held-2.js'

const FILE = new URL('./spill-cycles-off-held.predicted.json', import.meta.url)
const PREDICTED = (existsSync(FILE) ? JSON.parse(readFileSync(FILE, 'utf8')) : {}) as Record<string, Record<string, Val>[]>

/** The cases taken over, the iteration counts of `unique-typed-20` aside: without iterative calculation they are one case. */
const TAKEN = new Set([
  'sort-file',
  'sort-typed-B3-C2-A4',
  'sort-typed-A4-C2-B3',
  'munit-file',
  'munit-typed-C14-D11',
  'munit-typed-D11-C14',
  'sheets-file-Sheet2-Q_data',
  'sheets-file-Q_data-Sheet2',
  'sheets-typed-Sheet2.E5-Sheet2.D7-Q_data.F5-Q_data.D4',
  'sheets-typed-Q_data.F5-Q_data.D4-Sheet2.E5-Sheet2.D7',
  'wait-file',
  'wait-typed-B1-A2-C1-C2',
  'wait-typed-C2-C1-A2-B1',
  'unique-file',
  'unique-typed',
  'unique-found-typed',
  'sum-file',
  'sum-typed-array-last',
  'sum-typed-value-last',
])

/** The case without iterative calculation, its mirror sheets as before, its own predictions. */
function off({ predicted: _, ...c }: WorkbookCase): WorkbookCase {
  const { iterate: __, ...file } = c.file ?? {}
  return { ...c, about: c.about.replace(/, 20 iterations$/, ''), file, cycle: true, ...(PREDICTED[c.id] ? { predicted: PREDICTED[c.id] } : {}) }
}

/** `unique-file`'s two formulas typed into its empty sheet in `order`. */
function uniqueTyped(file: WorkbookCase, order: string[]): WorkbookCase {
  const cells = file.book.sheets[0].cells
  const steps: Step[] = order.map((cell) => ({ formula: { sheet: 'S', cell, formula: cells[cell].formula!, dynamic: true } }))
  return { ...file, id: `unique-typed-${order.join('-')}`, about: `${order.map((a) => `${a} = ${cells[a].formula}`).join(', then ')}, typed into the empty sheet`, book: { sheets: [{ name: 'S', cells: {} }] }, steps: [...steps, ...file.steps!] }
}

export function buildWorkbookCases(): WorkbookCase[] {
  const cases = [...held(), ...held2()].filter((c) => TAKEN.has(c.id))
  const file = cases.find((c) => c.id === 'unique-file')!
  return [...cases, uniqueTyped(file, ['A4', 'B2']), uniqueTyped(file, ['B2', 'A4'])].map(off)
}
