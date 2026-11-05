/**
 * Held out after `totals-held`: a last row of SUBTOTAL inside the AutoFilter's own range, A1:C22, in the file and as
 * Data > Filter on A1 takes it, under criteria and a sort, through the object model.
 */
import { regionOf, sales, total, type ActCase, type ActStep, type PredictedAct } from './cases.js'

const CRITERIA: ActStep = { criteria: 'Север', column: 'A', via: 'model' }
const SORT: ActStep = { sort: 'C', via: 'model' }
const cells = { ...sales(), ...total(22, 'SUBTOTAL(109,C2:C21)') }

export function buildActCases(): ActCase[] {
  const pick: ActStep = { over: 'A1', via: 'button' }
  return [
    { id: 'file-criteria', cells, filter: 'A1:C22', steps: [CRITERIA] },
    { id: 'file-sort', cells, filter: 'A1:C22', steps: [SORT] },
    { id: 'picked-criteria', cells, steps: [pick, CRITERIA] },
    { id: 'picked-sort', cells, steps: [pick, SORT] },
  ]
}

/** The model of `totals-held`: the row stays out of what the filter does whatever the AutoFilter's own range; criteria save the range they filtered, a sort leaves the own range. */
export function predictAct(c: ActCase): PredictedAct {
  const sort = c.id.endsWith('sort')
  const hidden = sort ? [] : Array.from({ length: 20 }, (_, k) => k + 2).filter((r) => regionOf(r) !== 'Север')
  const end = { range: 'A1:C22', hidden, total: { row: 22, text: sort ? '46030' : '10480' } }
  return { opens: true, after: [...(c.steps.length > 1 ? [{ range: 'A1:C22' }] : []), end], saved: { ref: sort ? 'A1:C22' : 'A1:C21', hidden, ...(sort ? { sortRef: 'A2:C21' } : {}) } } as unknown as PredictedAct
}
