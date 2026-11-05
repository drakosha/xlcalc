/**
 * An AutoFilter on A1:C21 with criteria on column A, removed: `AutoFilterMode = false` and Data > Filter. Over a row
 * hidden before the criteria that meets them and one that does not (`a`), the same with a height of 0 (`b`), the total
 * row right under it (`c`), a row hidden past empty rows (`d`), a row hidden without criteria (`e`), and criteria the
 * file holds over rows it hides, with the total row under them shown or not (`f`).
 */
import { regionOf, sales, total, type ActCase, type ActStep, type Via } from './cases.js'

const CRITERIA: ActStep = { criteria: 'Север', column: 'A', via: 'model' }
const FILE = '<filterColumn colId="0"><filters><filter val="Север"/></filters></filterColumn>'
const BELOW = (row: number) => ({ [`A${row}`]: { value: 'Прочее' } })
const failing = Array.from({ length: 20 }, (_, k) => k + 2).filter((r) => regionOf(r) !== 'Север')

export function buildActCases(): ActCase[] {
  const out: ActCase[] = []
  const filter = 'A1:C21'
  for (const via of ['model', 'button'] as Via[]) {
    const off: ActStep = { off: via }
    const steps = [CRITERIA, off]
    // row 5 does not meet the criteria, row 6 does
    out.push({ id: `a-fail-${via}`, cells: sales(), filter, hidden: [5], watch: [5], steps })
    out.push({ id: `a-pass-${via}`, cells: sales(), filter, hidden: [6], watch: [6], steps })
    out.push({ id: `b-fail-${via}`, cells: sales(), filter, zero: [7], watch: [7], steps })
    out.push({ id: `b-pass-${via}`, cells: sales(), filter, zero: [6], watch: [6], steps })
    out.push({ id: `c-${via}`, cells: { ...sales(), ...total(22, 'SUM(C2:C21)') }, filter, watch: [22], steps })
    out.push({ id: `d-25-${via}`, cells: { ...sales(), ...BELOW(25) }, filter, hidden: [25], watch: [25], steps })
    out.push({ id: `d-23-${via}`, cells: { ...sales(), ...BELOW(23) }, filter, hidden: [23], watch: [23], steps })
    // no criteria: row 2 would meet them, row 3 would not
    out.push({ id: `e-pass-${via}`, cells: sales(), filter, hidden: [2], watch: [2], steps: [off] })
    out.push({ id: `e-fail-${via}`, cells: sales(), filter, hidden: [3], watch: [3], steps: [off] })
    out.push({ id: `f-file-${via}`, cells: sales(), filter, hidden: failing, fileCriteria: FILE, watch: [3], steps: [off] })
    out.push({ id: `f-total-${via}`, cells: { ...sales(), ...total(22, 'SUM(C2:C21)') }, filter, hidden: failing, fileCriteria: FILE, watch: [3, 22], steps: [off] })
  }
  return out
}
