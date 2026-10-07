/**
 * Held-out array formulas of the old kind (`run.ts spill cse-held`): answers predicted from the rules of the set `cse`
 * and committed before Excel computed them (`PREDICTED`).
 */
import type { CellInput } from '../../../src/contract.js'
import { one, type CellSpec, type SpillCase } from './cases.js'
import { PREDICTED } from './cse-held.predicted.js'

const DATA: Record<string, CellSpec> = { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, B1: 10, B2: 20, B3: 30, B4: 40, B5: 50, C1: 100 }
const arr = (formula: string, area: string): CellInput => ({ formula, array: area })

const held = (id: string, about: string, cells: Record<string, CellSpec>, ranges: string[], extra: Partial<SpillCase> = {}): SpillCase =>
  one(id, about, { ...DATA, ...cells }, ranges, { ...extra, predicted: PREDICTED[id] })

export function buildSpillCases(): SpillCase[] {
  return [
    held('column-longer', 'a column over a longer area', { F1: arr('B1:B3/10', 'F1:F4') }, ['F1:F4']),
    held('outer-product', 'a column times a row', { F6: arr('A1:A2*B1:C1', 'F6:G7') }, ['F6:G7']),
    held('sum-row', 'one value over a row', { F9: arr('SUM(A1:A5)', 'F9:G9') }, ['F9:G9']),
    held('transpose-longer', 'TRANSPOSE over a longer row', { H1: arr('TRANSPOSE(B1:B3)', 'H1:K1') }, ['H1:K1']),
    held('row-times', 'ROW() times ten over a column', { H3: arr('ROW()*10', 'H3:H5') }, ['H3:H5']),
    held('single-max', 'one cell over products', { H7: arr('MAX(A1:A5*B1:B5)', 'H7') }, ['H7']),
    held('readers', 'formulas reading an area', { F1: arr('B1:B3/10', 'F1:F4'), I1: '=SUM(F1:F3)', I2: '=COUNT(F1:F4)' }, ['I1:I2']),
    held('precedent-edit', 'a value the array reads, edited', { J1: arr('A1:A3+1', 'J1:J3') }, ['J1:J3'], { edits: [{ sheet: 'S', cell: 'A2', value: 10 }] }),
    held('part-edit', 'a value typed into part of the area', { J5: arr('A1:A2', 'J5:J6') }, ['J5:J6'], { edits: [{ sheet: 'S', cell: 'J6', value: 9 }] }),
  ]
}
