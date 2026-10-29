/**
 * The AutoFilter over A1:C5 (a header and four rows, `data`) under what is entered around it: typed right under it, in
 * its columns and beside them, a row further down, pasted, copied, filled down, a formula and nothing, with criteria
 * and hidden rows, over the whole height of the sheet, cleared and undone, and written through the object model.
 */
import { collector, data, type FilterCase } from './cases.js'

export function buildFilterCases(): FilterCase[] {
  const { out, add } = collector()
  const cells = data('ABC', 5)
  const filter = 'A1:C5'
  // right under it, in each of its columns and beside them
  for (const at of ['A6', 'B6', 'C6', 'D6']) add('under', { cells, filter, steps: [{ type: '7', at }] })
  add('under', { cells: data('ABC', 5, 'B'), filter: 'B1:D5', steps: [{ type: '7', at: 'A6' }] })
  add('under', { cells, filter, steps: [{ type: 'North', at: 'A6' }, { type: 'South', at: 'A7' }] })
  add('under', { cells, filter, steps: [{ type: '7', at: 'D6' }, { type: '8', at: 'C6' }] })
  // a row between
  add('gap', { cells, filter, steps: [{ type: '7', at: 'A7' }] })
  add('gap', { cells, filter, steps: [{ type: '7', at: 'A7' }, { type: '8', at: 'A6' }] })
  // pasted, copied and filled down
  add('paste', { cells, filter, steps: [{ paste: 'North\t6\t6', at: 'A6' }] })
  add('paste', { cells, filter, steps: [{ paste: 'North\t6\t6\nSouth\t7\t7', at: 'A6' }] })
  add('paste', { cells, filter, steps: [{ paste: 'North\t6\t6', at: 'A7' }] })
  add('paste', { cells, filter, steps: [{ paste: '6', at: 'D6' }] })
  add('copy', { cells, filter, steps: [{ copy: 'A2:C2', to: 'A6' }] })
  add('copy', { cells, filter, steps: [{ copy: 'A2:C3', to: 'A6' }] })
  add('copy', { cells, filter, steps: [{ copy: 'B2', to: 'B6' }] })
  add('fill', { cells, filter, steps: [{ fill: 'A5:C6' }] })
  add('fill', { cells, filter, steps: [{ fill: 'A5:C7' }] })
  add('fill', { cells, filter, steps: [{ fill: 'B5:B6' }] })
  // a formula, a space and nothing typed
  add('entry', { cells, filter, steps: [{ type: '=B2*2', at: 'A6' }] })
  add('entry', { cells, filter, steps: [{ type: ' ', at: 'A6' }] })
  add('entry', { cells, filter, steps: [{ type: '', at: 'A6' }] })
  // criteria applied (North: rows 3 and 5 hidden), and rows hidden by hand
  const criteria = { field: 1, value: 'North', hides: [3, 5] }
  add('criteria', { cells, filter, criteria, steps: [{ type: 'North', at: 'A6' }] })
  add('criteria', { cells, filter, criteria, steps: [{ type: 'South', at: 'A6' }] })
  add('criteria', { cells, filter, hidden: [5], steps: [{ type: '7', at: 'A6' }] })
  // over the whole height of the sheet
  add('full', { cells, filter: 'A1:C1048576', steps: [{ type: '7', at: 'A6' }, { type: '8', at: 'D6' }] })
  // a header alone
  add('header', { cells: data('ABC', 1), filter: 'A1:C1', steps: [{ type: '7', at: 'B2' }] })
  // cleared: the last row, a row it grew over
  add('clear', { cells, filter, steps: [{ clear: 'A5:C5' }] })
  add('clear', { cells, filter, steps: [{ type: '7', at: 'A6' }, { clear: 'A6' }] })
  add('clear', { cells, filter, steps: [{ clear: 'A5:C5' }, { type: '7', at: 'A6' }] })
  // undone
  add('undo', { cells, filter, steps: [{ type: '7', at: 'A6' }, { undo: true }] })
  add('undo', { cells, filter, steps: [{ paste: 'North\t6\t6\nSouth\t7\t7', at: 'A6' }, { undo: true }] })
  add('undo', { cells, filter, steps: [{ type: '7', at: 'A6' }, { type: '8', at: 'A7' }, { undo: true }] })
  // through the object model
  add('set', { cells, filter, steps: [{ set: '7', at: 'A6' }] })
  add('set', { cells, filter, steps: [{ set: '7', at: 'A6' }, { set: '8', at: 'A7' }] })
  return out
}
