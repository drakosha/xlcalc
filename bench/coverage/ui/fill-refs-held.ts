/**
 * Held out from the entry sets `refs` and `refs-held`, predicted by the session before Excel ran it: a formula entered
 * into a selection (Control+Return) takes the format the active cell's formula takes from what it refers to, in every
 * cell, or each cell its own; AutoSum over a column whose first cell is a percent and the rest of other formats; a
 * formula typed into one cell of a column formatted as a percent. Typed on the keyboard: ASCII only.
 */
import type { CellInput } from '../../../src/contract.js'
import { collector, type UiCase } from './cases.js'

/** A1 and A2 percents, A3 2 in General, A4 a date, A5 money. */
const CELLS: Record<string, CellInput> = { A1: { value: 0.1, numberFormat: 9 }, A2: { value: 0.2, numberFormat: 9 }, A3: { value: 2 }, A4: { value: 46096, numberFormat: 14 }, A5: { value: 100, numberFormat: 7 } }

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const fill = (range: string, text: string, more: Partial<Omit<UiCase, 'id' | 'act'>> = {}) => add('fill', { ...more, cells: { ...CELLS, ...more.cells }, act: { do: 'fill', range, text } })
  fill('B1:B3', '=A1+A3')
  fill('B3:B5', '=A3+A1')
  fill('B1:C2', '=A3')
  fill('B1:B5', '=A1')
  fill('B1', '=A5-A3')
  fill('B2:B4', '=A3-A2')
  // a cell of a column formatted as money, typed into
  fill('B1', '=A1', { columns: { B: { numberFormat: 7 } } })
  add('sum', { cells: CELLS, act: { do: 'press', range: 'A6', buttons: ['AutoSum'], enter: true } })
  add('sum', { cells: { ...CELLS, A1: { value: 0.1 } }, act: { do: 'press', range: 'A6', buttons: ['AutoSum'], enter: true } })
  return out
}
