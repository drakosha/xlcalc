/**
 * Held out after `point-caret-held-2`, predicted by the revised `refAt` before Excel ran it
 * (`point-caret-held-3.predicted.json`): brackets left open, a function's (closed as Return closes them) and a group's
 * (left open, so the formula does not read), alone and nested in each other, at the end of the text and with the caret
 * before an operator inside it. Steps as in `point-caret-held-2`.
 */
import { collector, type PointStep, type UiCase } from './cases.js'
import type { CellInput } from '../../../src/contract.js'

export const POINT_BY = 'refAt'

const CELLS: Record<string, CellInput> = { A1: { value: 2 }, D3: { value: 10 }, D4: { value: 20 } }

/** Right pressed right after the text. */
const END = ['=SUM((1', '=(SUM(1', '=MAX(1,(2', '=ROUND(1.5', '=IF(A1>0,1', '=-(A1', '=SUM(A1)+(2', '=SUM(1)+MAX(2', '=ABS(-(1)']

/** A text and how many characters the caret goes back from its end. */
const INSIDE: [string, number][] = [
  ['=SUM(A1+2', 2],
  ['=(A1+2', 2],
  ['=MAX(A1,2)+(1', 4],
  ['=ROUND(A1*2', 2],
]

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const point = (group: string, text: string, moves: PointStep[] | null) =>
    add(group, { cells: CELLS, act: { do: 'point', at: 'C3', steps: [{ type: text }, ...(moves ? [{ key: 'f2' } as const, ...moves, { key: 'f2' } as const] : []), { key: 'right' }, { key: 'return' }] }, read: ['C3', 'C4'] })
  for (const t of END) point('end', t, null)
  for (const [t, back] of INSIDE) point('inside', t, Array.from({ length: back }, (): PointStep => ({ key: 'left' })))
  return out
}
