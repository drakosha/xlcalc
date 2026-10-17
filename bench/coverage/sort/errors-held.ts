/**
 * Held out, predicted before Excel ran it (`errors-held.predicted.json`) by Microsoft's words on the Sort command,
 * "All error values, such as #NUM! and #REF!, are equal": errors alone, in the order of ERROR.TYPE's numbers and
 * against it, and errors among numbers, text, booleans and a blank, ascending and descending. Two cells give #N/A by
 * different formulas, so an order between equal errors shows too.
 */
import type { SortCase, SortSet } from './cases.js'

const ERRORS = ['=NA()', '=1/0', '=#REF!', '=#N/A', '=#VALUE!', '=SQRT(-1)', '=#NAME?', '=#NULL!']
const MIXED = ['=NA()', '10', 'b', 'TRUE', '=1/0', '2', 'A', 'FALSE', '', '=#REF!']

const both = (id: string, cells: string[]): SortCase[] => [
  { id: `${id}-asc`, cells },
  { id: `${id}-desc`, cells, descending: true },
]

export const set: SortSet = {
  cases: [...both('errors', ERRORS), ...both('errors-reversed', [...ERRORS].reverse()), ...both('mixed', MIXED)],
}
