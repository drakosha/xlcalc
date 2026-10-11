/**
 * Held out: rules at the sheet's last row and column under lines inserted and deleted, predicted by the engine before
 * Excel ran them (`edge-held.predicted.json`). The engine moves a rule of one cell up (left) with deleted lines, as a
 * reference to a cell moves, and keeps a longer range ending on the last row (column) ending there; its conditional
 * formats stretch a cell there instead (the set `cond/edge-held` asks the same of them).
 */
import type { BookInput, CellInput, DataValidation, StructureOp } from '../../../src/contract.js'
import type { ValidCase, ValidCases } from './cases.js'

/** A rule told apart by its input message, whole numbers from 1 to 10. */
const rule = (range: string, prompt: string): DataValidation => ({ range, type: 'whole', formula1: '1', formula2: '10', prompt })

/** The rules of S, in its bottom-right corner: single cells, and ranges of two along and across the last row and column. */
export const EDGE_RULES: [range: string, about: string][] = [
  ['XFD1048576', 'corner'],
  ['XFB1048576', 'last row'],
  ['XFD1048570', 'last column'],
  ['XEZ1048576:XFA1048576', 'one row of two on the last row'],
  ['XFC1048575:XFC1048576', 'two rows to the last row'],
  ['XFD1048565:XFD1048566', 'one column of two on the last column'],
  ['XFC1048568:XFD1048568', 'two columns to the last column'],
]

/** S: the rules, each of their cells holding 5 (50 at the corner, which its rule refuses). */
function book(): BookInput {
  const validations = EDGE_RULES.map(([range, about]) => rule(range, about))
  const cells: Record<string, CellInput> = { A1: { value: 1 } }
  for (const [range] of EDGE_RULES) for (const a of range.split(':')) cells[a] = { value: 5 }
  cells.XFD1048576 = { value: 50 }
  return { sheets: [{ name: 'S', cells, validations }] }
}

/** The structure operations of both edge sets: lines deleted and inserted far before the rules, and the last line deleted. */
export const EDGE_OPS: [id: string, op: StructureOp][] = [
  ['delete-row-5', { op: 'deleteRows', sheet: 'S', row: 5 }],
  ['delete-rows-5x3', { op: 'deleteRows', sheet: 'S', row: 5, count: 3 }],
  ['delete-last-row', { op: 'deleteRows', sheet: 'S', row: 1048576 }],
  ['delete-col-b', { op: 'deleteColumns', sheet: 'S', column: 'B' }],
  ['delete-cols-bx3', { op: 'deleteColumns', sheet: 'S', column: 'B', count: 3 }],
  ['delete-last-col', { op: 'deleteColumns', sheet: 'S', column: 'XFD' }],
  ['insert-row-5', { op: 'insertRows', sheet: 'S', row: 5 }],
  ['insert-col-b', { op: 'insertColumns', sheet: 'S', column: 'B' }],
]

/** The rectangle both edge sets read: the corner, with room for three lines deleted. */
export const EDGE_WINDOW = 'XEV1048558:XFD1048576'

export function buildValidCases(): ValidCases {
  const cases: ValidCase[] = EDGE_OPS.map(([id, op]) => ({ id, book: 'edge', op }))
  return { books: { edge: book() }, cases, window: EDGE_WINDOW }
}
