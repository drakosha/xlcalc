/**
 * Held out after the rule of `grow` was fitted (`getAutoFilterRange`): the AutoFilter over A1:C5 under what is entered
 * further from it, a column away and diagonally, a formula that shows nothing and one that spills, data under it before
 * any step, a row emptied between rows that hold something, a hidden row, a paste beside it, and criteria applied after
 * rows were entered under it, which `grow` did not try.
 */
import { collector, data, type FilterCase } from './cases.js'

export function buildFilterCases(): FilterCase[] {
  const { out, add } = collector()
  const cells = data('ABC', 5)
  const filter = 'A1:C5'
  // a column away on either side, and diagonally beside the column on its right
  add('far', { cells, filter, steps: [{ type: '7', at: 'E6' }] })
  add('far', { cells: data('ABC', 5, 'C'), filter: 'C1:E5', steps: [{ type: '7', at: 'A6' }] })
  add('far', { cells, filter, steps: [{ type: '7', at: 'D6' }, { type: '8', at: 'E7' }] })
  // a formula that shows nothing, one that spills down
  add('formula', { cells, filter, steps: [{ type: '=""', at: 'A6' }] })
  add('formula', { cells, filter, steps: [{ type: '=SEQUENCE(3)', at: 'B6' }] })
  // rows under it that hold data before anything is entered
  add('below', { cells: { ...data('ABC', 7) }, filter, steps: [{ type: '9', at: 'H1' }] })
  // a row emptied between two that hold something; a hidden row
  add('emptied', { cells, filter, steps: [{ type: '7', at: 'A6' }, { type: '8', at: 'A7' }, { clear: 'A6' }] })
  add('hidden', { cells, filter, hidden: [6], steps: [{ type: '7', at: 'A6' }] })
  // pasted into the column on its right alone
  add('beside', { cells, filter, steps: [{ paste: '6\n7\n8', at: 'D6' }] })
  // criteria after rows entered under it: what the file then holds
  add('criteria', { cells, filter, steps: [{ type: 'North', at: 'A6' }, { type: 'South', at: 'A7' }, { criteria: { field: 1, value: 'North' } }] })
  add('criteria', { cells, filter, steps: [{ type: '7', at: 'D6' }, { criteria: { field: 2, value: '7' } }] })
  return out
}
