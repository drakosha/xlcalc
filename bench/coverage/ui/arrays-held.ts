/**
 * Held out, predicted by `ViewportCell.editText` before Excel ran it: what the formula bar shows for a cell of an array,
 * selected and once its editing began (F2). A dynamic array's anchor and the cells it spilled into (`SEQUENCE`, a range
 * times a number), one blocked (`#SPILL!`) with the empty cell and the value in its way, formulas of the old kind over an
 * area (a sum in every cell of G1:H2, a range times a number down I1:I3) and over one cell, and for comparison plain
 * formulas, a reference to a spill and a value. The same book in every case, each case its cells.
 */
import type { CellInput } from '../../../src/contract.js'
import { collector, type UiCase } from './cases.js'

const f = (formula: string, array?: string): CellInput => ({ formula, ...(array ? { array } : {}) })

const CELLS: Record<string, CellInput> = {
  C1: { value: 1 }, C2: { value: 2 }, C3: { value: 3 },
  A1: f('SEQUENCE(3)'), B1: f('C1:C3*2'),
  E1: f('SEQUENCE(3)'), E3: { value: 'x' },
  G1: f('SUM(C1:C3)', 'G1:H2'), I1: f('C1:C3*2', 'I1:I3'), K1: f('SUM(C1:C3*2)', 'K1'),
  M1: f('C1*2'), N1: f('SUM(C1:C3)'), O1: f('SUM(A1#)'), P1: { value: 1234.5 },
}

const GROUPS: [string, string[]][] = [
  ['spill', ['A1', 'A2', 'A3', 'B1', 'B2']],
  ['blocked', ['E1', 'E2', 'E3']],
  ['legacy', ['G1', 'H1', 'G2', 'H2', 'I1', 'I2', 'I3', 'K1']],
  ['plain', ['M1', 'N1', 'O1', 'P1']],
]

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  for (const [group, cells] of GROUPS) add(group, { cells: CELLS, act: { do: 'bars', cells } })
  return out
}
