/**
 * Held out after `removal` and `totals-held`: a filter removed over the row right under its range, hidden. Criteria the
 * file holds over rows it hides, its AutoFilter on A1:C21 and row 22 hidden too, with a SUM and with a SUBTOTAL; criteria
 * Excel applies over a row 22 of SUBTOTAL the book hides. Through `AutoFilterMode = false` and Data > Filter.
 */
import { regionOf, sales, total, type ActCase, type PredictedAct, type Via } from './cases.js'

const FILE = '<filterColumn colId="0"><filters><filter val="Север"/></filters></filterColumn>'
const failing = Array.from({ length: 20 }, (_, k) => k + 2).filter((r) => regionOf(r) !== 'Север')
const SUB = 'SUBTOTAL(109,C2:C21)'
/** Whether row 22 is out of what the filter does, in the model. */
const OUT: Record<string, boolean> = { 'file-sum': false, 'file-sub': true, 'criteria-sub': true }

export function buildActCases(): ActCase[] {
  const out: ActCase[] = []
  const filter = 'A1:C21'
  for (const via of ['model', 'button'] as Via[]) {
    out.push({ id: `file-sum-${via}`, cells: { ...sales(), ...total(22, 'SUM(C2:C21)') }, filter, hidden: [...failing, 22], fileCriteria: FILE, watch: [22], steps: [{ off: via }] })
    out.push({ id: `file-sub-${via}`, cells: { ...sales(), ...total(22, SUB) }, filter, hidden: [...failing, 22], fileCriteria: FILE, watch: [22], steps: [{ off: via }] })
    out.push({ id: `criteria-sub-${via}`, cells: { ...sales(), ...total(22, SUB) }, filter, hidden: [22], watch: [22], steps: [{ criteria: 'Север', column: 'A', via: 'model' }, { off: via }] })
  }
  return out
}

/** The model of `totals-held`: removing a filter shows the rows of the range it reads but a last row of SUBTOTAL, which stays as it was. */
export function predictAct(c: ActCase): PredictedAct {
  const left = OUT[c.id.replace(/-(model|button)$/, '')] ? [22] : []
  return { opens: true, saved: { ref: null, hidden: left }, after: [...(c.steps.length > 1 ? [{}] : []), { range: '', hidden: left }] } as unknown as PredictedAct
}
