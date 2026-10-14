/**
 * Held out after `point-caret-held`, predicted by the revised `refAt` before Excel ran it
 * (`point-caret-held-2.predicted.json`): the caret at the start of other formulas and of values, where the arrow puts in
 * an = of its own and the reference; formulas whose brackets are left open, entered as Return enters them, with the
 * caret at the end and before an operator, a separator or a closing bracket inside them. Steps as in `point-caret-held`:
 * F2, the caret moved, F2, then Right and Return; a case with no move presses Right right after the text.
 */
import { collector, type PointStep, type UiCase } from './cases.js'
import type { CellInput } from '../../../src/contract.js'

export const POINT_BY = 'refAt'

const CELLS: Record<string, CellInput> = { A1: { value: 2 }, D3: { value: 10 }, D4: { value: 20 } }

/** At the start of the text. */
const START = ['=SUM(1,2)', '-5', '=A1+B1', '=1/0', '= A1', '5', 'abc']

/** A text and how many characters the caret goes back from its end. */
const INSIDE: [string, number][] = [
  ['=SUM(A1,2', 2],
  ['=(1+2', 2],
  ['=A1 +1', 3],
  ['=A1)', 1],
  ['=SUM("a",', 1],
  ['="a"&', 1],
  ['=SUM(1+2)*', 1],
]

/** Right pressed right after the text: brackets left open at the end. */
const END = ['=SUM(A1', '=(1+2', '=((1', '=SUM(A1,2', '=MAX(1;2']

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const point = (group: string, text: string, moves: PointStep[] | null) =>
    add(group, { cells: CELLS, act: { do: 'point', at: 'C3', steps: [{ type: text }, ...(moves ? [{ key: 'f2' } as const, ...moves, { key: 'f2' } as const] : []), { key: 'right' }, { key: 'return' }] }, read: ['C3', 'C4'] })
  for (const t of START) point('start', t, [{ key: 'home' }])
  for (const [t, back] of INSIDE) point('inside', t, Array.from({ length: back }, (): PointStep => ({ key: 'left' })))
  for (const t of END) point('end', t, null)
  return out
}
