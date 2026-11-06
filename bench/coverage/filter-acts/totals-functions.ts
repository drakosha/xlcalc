/**
 * Held out after `totals`: AVERAGE and COUNT in a total row under an AutoFilter on A1:C21, right under it and past an empty row, under
 * criteria and a sort of the object model. The read of Excel for Windows's code had it look for SUBTOTAL alone; `totals` measured SUM.
 *
 * The model, from SUM in `totals`: a row of these under the range is part of it, hidden by criteria, which save the range read, and sorted
 * with the rest, by its value, which a sort saves over the AutoFilter's own range. So AVERAGE, 2301.5, goes between the amounts, to row 12,
 * its reference moved up by ten rows and stopped at row 1, and COUNT, 20, stays last. Past an empty row it is out of everything.
 */
import { regionOf, sales, total, type ActCase, type ActStep, type PredictedAct } from './cases.js'

const FUNCTIONS: [string, string, string][] = [
  ['average', 'AVERAGE(C2:C21)', '2301.5'],
  ['count', 'COUNT(C2:C21)', '20'],
]
const CRITERIA: ActStep = { criteria: 'Север', column: 'A', via: 'model' }
const SORT: ActStep = { sort: 'C', via: 'model' }

export function buildActCases(): ActCase[] {
  return FUNCTIONS.flatMap(([name, formula]) =>
    [
      ['next', 22],
      ['gap', 23],
    ].flatMap(([where, row]) => [
      { id: `${name}-${where}-criteria`, cells: { ...sales(), ...total(row as number, formula) }, filter: 'A1:C21', steps: [CRITERIA] },
      { id: `${name}-${where}-sort`, cells: { ...sales(), ...total(row as number, formula) }, filter: 'A1:C21', steps: [SORT] },
    ]),
  )
}

export function predictAct(c: ActCase): PredictedAct {
  const [, , text] = FUNCTIONS.find(([name]) => c.id.startsWith(name))!
  const next = c.id.includes('-next-')
  const range = next ? 'A1:C22' : 'A1:C21'
  const row = next ? 22 : 23
  if (c.id.endsWith('sort')) {
    const moved = next && c.id.startsWith('average')
    return { opens: true, after: [{ range, filterMode: false, hidden: [], total: { row: moved ? 12 : row, text: moved ? '2986.5' : text } }], saved: { ref: 'A1:C21', criteria: {}, sortRef: next ? 'A2:C22' : 'A2:C21', hidden: [] } } as unknown as PredictedAct
  }
  const hidden = Array.from({ length: next ? 21 : 20 }, (_, k) => k + 2).filter((r) => r > 21 || regionOf(r) !== 'Север')
  return { opens: true, after: [{ range, filterMode: true, hidden, total: { row, text } }], saved: { ref: range, criteria: { '0': 'Север' }, hidden } } as unknown as PredictedAct
}
