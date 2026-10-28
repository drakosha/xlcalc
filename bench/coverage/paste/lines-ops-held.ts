/**
 * Held out after the set `lines-ops` showed that Paste Special lays a copy that left out rows a filter hid over the
 * target whole, hidden rows too, and skips no blank: predicted by the engine before Excel ran them
 * (`lines-ops-held.predicted.json`), in the book of `lines-ops`. Such a copy with values, formats and all alone over
 * hidden rows, transposed, with Skip Blanks onto a sheet without a filter, onto one cell of the filtered sheet, a whole
 * column with Skip Blanks; and a plain copy over hidden rows, which Excel refuses.
 */
import type { PasteCase, PasteCases, PasteOperation, PasteWhat } from './cases.js'
import { opsBook } from './lines-ops.js'

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, what: PasteWhat, o: { operation?: PasteOperation; skipBlanks?: boolean; transpose?: boolean } = {}, toSheet = 'F') =>
    cases.push({ id, book: 'ops', from: { sheet: 'F', range: from }, to: { sheet: toSheet, range: to }, what, ...o })
  add('values-over-hidden', 'A2:A4', 'C2:C7', 'values')
  add('formats-over-hidden', 'A2:A4', 'C2:C7', 'formats')
  add('all-over-hidden', 'A2:A4', 'C2:C7', 'all')
  add('copy-over-hidden', 'A2:A4', 'C2:C7', 'copy')
  add('values-transposed', 'A2:A8', 'H1', 'values', { transpose: true }, 'S')
  add('values-skip', 'A2:A8', 'C2', 'values', { skipBlanks: true }, 'S')
  add('values-one-cell', 'A2:A8', 'E2', 'values')
  add('column-skip', 'A:A', 'C:C', 'all', { skipBlanks: true })
  return { books: { ops: opsBook() }, cases, plain: { ops: ['A5', 'C2', 'C3', 'C4', 'C5', 'D2', 'F3'] } }
}
