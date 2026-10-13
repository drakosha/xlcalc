/**
 * Held out, predicted before Excel ran it by the sheet component's own model of Point mode (`point-model.ts`, after
 * `xlcalc-sheet/src/core/point.ts`): a formula typed over C3 (Enter mode) and arrows pressed, where a reference may stand
 * and where it may not (after `=`, an operator, `(`, `:`, the list separator, a space after them; after a number, a
 * reference typed, a closing bracket, an open string, a word), then keys pressed while pointing (an arrow again, Shift
 * and an arrow, Command and an arrow, Home, Page Down, a text typed and an arrow), and F2 between Enter and Edit mode;
 * each case ends with Return. Both locales' list separators are in the set, which runs in each: `,` separates arguments
 * in en-US and `;` in ru-RU. Typed on the keyboard: ASCII only. Mouse presses and a sheet's tab are not measured here.
 */
import { collector, type PointStep, type UiCase } from './cases.js'
import type { CellInput } from '../../../src/contract.js'

/** Cells around C3: a column of numbers below D3 for Command and an arrow to stop at, a text beside, a number above. */
const CELLS: Record<string, CellInput> = { A1: { value: 2 }, C1: { value: 1 }, D3: { value: 10 }, D4: { value: 20 }, D5: { value: 30 }, D6: { value: 40 }, B6: { value: 5 }, E3: { value: 'x' } }

const t = (type: string): PointStep => ({ type })
const RIGHT: PointStep = { key: 'right' }
const RETURN: PointStep = { key: 'return' }

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const point = (group: string, steps: PointStep[]) => add(group, { cells: CELLS, act: { do: 'point', at: 'C3', steps: [...steps, RETURN] }, read: ['C4', 'D3'] })

  // (a) an arrow where a reference may stand, by the component's rule
  for (const text of ['=', '=A1+', '=SUM(', '=SUM(A1,', '=SUM(A1;', '=A1:', '=SUM( ', '="x"&', '=-', '=A1>', '=(', '+', '-']) point('fits', [t(text), RIGHT])
  // (a) and where it may not
  for (const text of ['=5', '=A1', '=A1 ', '="a', '=TRUE', '=SUM(A1)', 'abc', '=SUM', '=1.']) point('not', [t(text), RIGHT])
  // (c) keys while pointing, after `=` and Right put D3
  point('move', [t('='), RIGHT, RIGHT])
  point('move', [t('='), RIGHT, { key: 'down' }])
  point('move', [t('='), { key: 'up' }])
  point('move', [t('='), { key: 'left' }, { key: 'left' }, { key: 'left' }])
  point('shift', [t('='), RIGHT, { key: 'right', shift: true }, { key: 'down', shift: true }])
  point('shift', [t('='), { key: 'right', shift: true }])
  // (d) a plain arrow after a range was extended: from its first cell or from the far corner
  point('after', [t('='), RIGHT, { key: 'down', shift: true }, { key: 'down' }])
  point('after', [t('='), RIGHT, { key: 'right', shift: true }, RIGHT])
  point('jump', [t('='), RIGHT, { key: 'down', cmd: true }])
  point('jump', [t('='), RIGHT, { key: 'down', cmd: true, shift: true }])
  point('jump', [t('='), RIGHT, { key: 'home' }])
  point('jump', [t('='), RIGHT, { key: 'pagedown' }])
  point('typed', [t('='), RIGHT, t('+'), RIGHT])
  point('typed', [t('='), RIGHT, t('+1'), RIGHT])
  // (e) F2 between Enter and Edit mode
  point('f2', [t('='), { key: 'f2' }, RIGHT])
  point('f2', [t('=A1+'), { key: 'f2' }, RIGHT])
  point('f2', [t('='), { key: 'f2' }, { key: 'f2' }, RIGHT])
  point('f2', [t('='), RIGHT, { key: 'f2' }, RIGHT])
  return out
}
