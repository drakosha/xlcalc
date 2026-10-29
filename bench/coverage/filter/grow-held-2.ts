/**
 * Held out after the current region was fitted to `grow-held` (`getAutoFilterRange`): a diagonal two columns and
 * three columns out from the AutoFilter, a block to its right that touches it at a corner alone and one that does not,
 * data above its header row that the region reaches and leads down a column further right, and a column beside it, or
 * one column further, that goes lower than it. A case whose data are there before it takes one step far from them (J20).
 */
import type { CellInput } from '../../../src/contract.js'
import { collector, data, type FilterCase, type FilterStep } from './cases.js'

/** Cells of `range` (`D6:E9`) each holding its row's number. */
function block(range: string): Record<string, CellInput> {
  const [a, b] = range.split(':')
  const [c1, r1, c2, r2] = [a.charCodeAt(0), Number(a.slice(1)), b.charCodeAt(0), Number(b.slice(1))]
  const out: Record<string, CellInput> = {}
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out[`${String.fromCharCode(c)}${r}`] = { value: r }
  return out
}

/** Cells moved `n` rows down. */
const lower = (cells: Record<string, CellInput>, n: number) => Object.fromEntries(Object.entries(cells).map(([k, v]) => [k.replace(/\d+$/, (r) => String(Number(r) + n)), v]))

export function buildFilterCases(): FilterCase[] {
  const { out, add } = collector()
  const cells = data('ABC', 5)
  const filter = 'A1:C5'
  const away: FilterStep[] = [{ type: '9', at: 'J20' }]
  // a diagonal two and three columns out, one with a column missing
  add('diagonal', { cells, filter, steps: [{ type: '1', at: 'D6' }, { type: '2', at: 'E7' }, { type: '3', at: 'F8' }] })
  add('diagonal', { cells, filter, steps: [{ type: '1', at: 'D6' }, { type: '2', at: 'E7' }, { type: '3', at: 'F8' }, { type: '4', at: 'G9' }] })
  add('diagonal', { cells, filter, steps: [{ type: '1', at: 'D6' }, { type: '3', at: 'F7' }] })
  // a block to its right touching it at a corner alone, and one a column further
  add('island', { cells: { ...cells, ...block('D6:E9') }, filter, steps: away })
  add('island', { cells: { ...cells, ...block('E6:F9') }, filter, steps: away })
  // data above its header row: a row right above, then a block above that leads down column E
  add('above', { cells: { ...block('A1:C1'), ...lower(cells, 1) }, filter: 'A2:C6', steps: [{ type: '7', at: 'A7' }] })
  add('above', { cells: { ...block('A1:E2'), ...block('E3:E10'), ...lower(cells, 2) }, filter: 'A3:C7', steps: away })
  // a column beside it, or one further, going lower than it
  add('column', { cells: { ...cells, ...block('D1:D9') }, filter, steps: away })
  add('column', { cells: { ...cells, ...block('E1:E9') }, filter, steps: away })
  add('column', { cells: { ...data('BCD', 5, 'B'), ...block('A1:A9') }, filter: 'B1:D5', steps: away })
  return out
}
