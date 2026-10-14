/**
 * Held out after `point-third-held`, predicted by the revised `refAt` before Excel ran it (`point-caret-held.predicted.json`):
 * an arrow in Point mode with the caret elsewhere than at the end, which `refAt` had not been measured at. A text typed
 * over C3, F2 to Edit mode, the caret moved (Home to the start, arrows by one), F2 back to Point mode, then Right and
 * Return: the caret before the formula's first character, after it, and before an operator inside the formula, where
 * the text before the caret reads and the whole text does not, or the other way round; F2 twice with the caret left at
 * the end, as the arrow at the end was measured, for a control.
 */
import { collector, type PointStep, type UiCase } from './cases.js'
import type { CellInput } from '../../../src/contract.js'

export const POINT_BY = 'refAt'

const CELLS: Record<string, CellInput> = { A1: { value: 2 }, D3: { value: 10 }, D4: { value: 20 } }

/** Texts with the caret at 0, then at 1. */
const START = ['=A1', '=5', '=(', '=1+', '+1', '=', '=A1)', '="a"', '=SUM(A1', '-A1']
const AFTER_FIRST = ['=A1', '=5', '+1', '=(']

/** A text and how many characters the caret goes back from its end: before an operator, inside the formula. */
const INSIDE: [string, number][] = [
  ['=A1+(', 2],
  ['=A1+B1', 3],
  ['=A1+', 1],
  ['=A1)+1', 3],
  ['=5*', 1],
  ['=SUM(A1,', 1],
]

/** F2 twice, the caret left at the end. */
const CONTROL = ['=A1', '=1+', '=A1)', '=(']

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const point = (group: string, text: string, moves: PointStep[]) => add(group, { cells: CELLS, act: { do: 'point', at: 'C3', steps: [{ type: text }, { key: 'f2' }, ...moves, { key: 'f2' }, { key: 'right' }, { key: 'return' }] }, read: ['C3', 'C4'] })
  for (const t of START) point('start', t, [{ key: 'home' }])
  for (const t of AFTER_FIRST) point('first', t, [{ key: 'home' }, { key: 'right' }])
  for (const [t, back] of INSIDE) point('inside', t, Array.from({ length: back }, (): PointStep => ({ key: 'left' })))
  for (const t of CONTROL) point('control', t, [])
  return out
}
