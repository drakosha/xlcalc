/**
 * Held out: conditional formats at the sheet's last row and column under lines inserted and deleted, predicted by the
 * engine before Excel ran them (`edge-held.predicted.json`), on the rules and operations of the set `valid/edge-held`.
 * The engine stretches a rule of one cell on the last row (column) over the line before it when lines are deleted
 * above (left of) it, as a range ending there keeps ending there, where its data validation moves such a cell.
 */
import type { BookInput, CellInput, ConditionalFormat } from '../../../src/contract.js'
import { EDGE_OPS, EDGE_RULES, EDGE_WINDOW } from '../valid/edge-held.js'
import type { CondCases } from './cases.js'

/** A fill of its own for each rule, so a cell shows which rule it is under. */
const FILLS = ['#FFC7CE', '#C6EFCE', '#FFEB9C', '#BDD7EE', '#F8CBAD', '#D9D9D9', '#E2EFDA']

/** S: a rule over each range of `EDGE_RULES` for a number above 0, each of their cells holding 5. */
function book(): BookInput {
  const conditionalFormats: ConditionalFormat[] = EDGE_RULES.map(([range], i) => ({ type: 'cellIs', ranges: [range], operator: 'greaterThan', formulas: ['0'], style: { fill: { color: FILLS[i] } } }))
  const cells: Record<string, CellInput> = { A1: { value: 1 } }
  for (const [range] of EDGE_RULES) for (const a of range.split(':')) cells[a] = { value: 5 }
  return { sheets: [{ name: 'S', cells, conditionalFormats }] }
}

export function buildCondCases(): CondCases {
  return { books: { edge: book() }, cases: EDGE_OPS.map(([id, op]) => ({ id, book: 'edge', op, read: [['S', EDGE_WINDOW]] })) }
}
