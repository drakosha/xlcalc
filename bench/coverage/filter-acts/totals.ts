/**
 * A total row under an AutoFilter on A1:C21, or a table's in its place: right under it with SUM, SUBTOTAL(109) and
 * SUBTOTAL(9) or a label alone, and past an empty row with SUM and SUBTOTAL(109). Each with criteria on column A, and on
 * a copy of its own sorted by column C, through the object model and on the column's button. And without a filter, the
 * range Excel takes for one: a cell of the header, and the range without the total row selected.
 */
import { sales, total, type ActCase, type ActStep } from './cases.js'

const SUM = 'SUM(C2:C21)'
const VARIANTS: [string, number, string | undefined][] = [
  ['sum', 22, SUM],
  ['sub109', 22, 'SUBTOTAL(109,C2:C21)'],
  ['sub9', 22, 'SUBTOTAL(9,C2:C21)'],
  ['label', 22, undefined],
  ['gap-sum', 23, SUM],
  ['gap-sub109', 23, 'SUBTOTAL(109,C2:C21)'],
]
const ACTIONS: [string, ActStep][] = [
  ['criteria-model', { criteria: 'Север', column: 'A', via: 'model' }],
  ['criteria-button', { criteria: 'Север', column: 'A', via: 'button' }],
  ['sort-model', { sort: 'C', via: 'model' }],
  ['sort-button', { sort: 'C', via: 'button' }],
]

export function buildActCases(): ActCase[] {
  const out: ActCase[] = []
  for (const [name, row, formula] of VARIANTS) {
    const cells = { ...sales(), ...total(row, formula) }
    for (const [act, step] of ACTIONS) out.push({ id: `${name}-${act}`, cells, filter: 'A1:C21', steps: [step] })
  }
  for (const [act, step] of ACTIONS) out.push({ id: `table-${act}`, cells: sales(), table: 'A1:C21', steps: [{ totals: true }, { set: 'Итого', at: 'A22' }, step] })
  for (const [name, row, formula] of VARIANTS.filter(([n]) => n === 'sum' || n === 'sub109' || n === 'gap-sum')) {
    const cells = { ...sales(), ...total(row, formula) }
    out.push({ id: `pick-${name}-cell-button`, cells, steps: [{ over: 'A1', via: 'button' }] })
    out.push({ id: `pick-${name}-cell-model`, cells, steps: [{ over: 'A1', via: 'plain' }] })
    out.push({ id: `pick-${name}-range-button`, cells, steps: [{ over: 'A1:C21', via: 'button' }] })
  }
  return out
}
