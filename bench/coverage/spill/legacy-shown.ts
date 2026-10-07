/**
 * Formulas whose old kind and the dynamic one Excel 365 shows for it part in Excel (`run.ts spill legacy-shown`): the
 * cases of the `legacy` sets where the engine, computing the text Excel shows, does not give the old formula's value.
 * Here that text is computed as a dynamic formula, at the row the old one stood in: Excel's answer tells whether the
 * two kinds part in Excel itself.
 */
import { dataCells } from '../data.js'
import { fromLegacy } from '../../../src/index.js'
import { sheet, type CellSpec, type SpillCase } from './cases.js'
import NESTING from '../../../test/excel/legacy/nesting.json' with { type: 'json' }
import PARAMS from '../../../test/excel/legacy/params.json' with { type: 'json' }

const PARTED: Record<string, string[]> = {
  params: ['GROWTH#r4', 'MATCH#r3', 'NETWORKDAYS.INTL#r3', 'NETWORKDAYS.INTL#e3', 'TREND#r4', 'XIRR#r3'],
  nesting: ['ACCRINT#array-e1', 'ACCRINT#array-e2', 'CUMIPMT#array-e1', 'GROWTH#array-r4', 'MATCH#array-r3', 'PRICE#array-e1', 'TREND#array-r4', 'XIRR#array-r3'],
}

/** The data sheet D as the case books have it: errors and text as their formulas. */
function data(): Record<string, CellSpec> {
  const out: Record<string, CellSpec> = {}
  for (const [addr, v] of dataCells()) out[addr] = typeof v !== 'object' ? v : 'formula' in v ? `=${v.formula}` : v.date
  return out
}

export function buildSpillCases(): SpillCase[] {
  const sets = { params: PARAMS, nesting: NESTING } as Record<string, { cases: { id: string; formula: string }[]; rows: Record<string, number> }>
  return Object.entries(PARTED).map(([set, ids]) => {
    const { cases, rows } = sets[set]
    const cells = Object.fromEntries(ids.map((id) => [`A${rows[id]}`, `=${fromLegacy(cases.find((c) => c.id === id)!.formula.replace(/_xlfn\.(_xlws\.)?|_xlpm\./g, ''))}`]))
    return { id: set, about: `the dynamic text of the ${set} cases that part`, book: { sheets: [sheet('D', data()), sheet('C', cells)] }, read: ids.map((id) => ({ sheet: 'C', range: `A${rows[id]}` })) }
  })
}
