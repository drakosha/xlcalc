/**
 * A font size or a rotation set on whole columns through the object model, each case in a book of its own: whether
 * Excel fits the rows that hold no cell of those columns, as it fits the rows that hold one, and a row set by hand.
 */
import { textOf } from '../fit/autofit.js'
import { collector, one, type RowCase } from './cases.js'

/** The rows read beside the case's (2): the first, the next, far down and the last. */
const MORE = [1, 3, 50, 1000, 1048576]

export function buildRowCases(): RowCase[] {
  const { out, add } = collector()
  add('size', { cells: one('a'), patched: { cols: [0, 0], size: 20 }, more: MORE })
  add('size', { patched: { cols: [0, 0], size: 20 }, more: MORE })
  add('size', { cells: one('a'), patched: { cols: [0, 1], size: 20 }, more: MORE, fitAfter: true })
  add('size', { row: { height: 30 }, cells: one('a'), patched: { cols: [0, 0], size: 20 }, more: MORE })
  add('rotation', { cells: one(textOf(46)), patched: { cols: [0, 0], rotation: 90 }, more: MORE })
  add('rotation', { patched: { cols: [0, 0], rotation: 90 }, more: MORE })
  add('rotation', { cells: one('bbbbb'), patched: { cols: [0, 0], rotation: 'vertical' }, more: MORE })
  return out
}
