/**
 * Held out after `sheets-held`, whose sheet moved before a hidden one went after it: the rule taken from that, a sheet
 * moved before a hidden sheet goes before the next visible one (or to the end), against moves before two hidden sheets
 * in a row, a very hidden one and a hidden last sheet, a hidden sheet moved, and sheets inserted before hidden ones and
 * after a very hidden last sheet. References across sheets show where the sheets went. `sheets-held-2.predicted.json`,
 * the engine's answers, was committed before Excel ran the set.
 */
import type { BookInput, SheetInput } from '../../../src/contract.js'
import type { HiddenCase, HiddenCases } from './cases.js'

const NAMES = ['Jan', 'Feb', 'Mar', 'Apr'] as const

const SUM: SheetInput = {
  name: 'Sum',
  cells: {
    A1: { formula: "IFERROR(SUM('Jan:Mar'!B2),\"ref\")" },
    A2: { formula: "IFERROR(SUM('Feb:Apr'!B2),\"ref\")" },
    A3: { formula: "IFERROR(COUNT('Jan:Apr'!B2),\"ref\")" },
  },
}

/** Jan to Apr (B2 1, 10, 100, 1000) and Sum, each sheet's `hidden` as given. */
function months(hidden: Partial<Record<(typeof NAMES)[number] | 'Sum', true | 'very'>> = {}): BookInput {
  const h = (name: keyof typeof hidden) => (hidden[name] ? { hidden: hidden[name] } : {})
  return { sheets: [...NAMES.map((name, i) => ({ name, cells: { B2: { value: 10 ** i } }, ...h(name) })), { ...SUM, ...h('Sum') }] }
}

const BOOKS: Record<string, BookInput> = {
  two: months({ Feb: true, Mar: true }),
  very: months({ Feb: 'very' }),
  feb: months({ Feb: true }),
  sumHidden: months({ Sum: true }),
  sumVery: months({ Sum: 'very' }),
}

const CASES: HiddenCase[] = [
  { id: 'move-before-two', book: 'two', op: { op: 'moveSheet', sheet: 'Sum', before: 'Feb' } },
  { id: 'move-next-visible', book: 'two', op: { op: 'moveSheet', sheet: 'Apr', before: 'Feb' } },
  { id: 'move-before-second', book: 'two', op: { op: 'moveSheet', sheet: 'Jan', before: 'Mar' } },
  { id: 'move-hidden-before-hidden', book: 'two', op: { op: 'moveSheet', sheet: 'Feb', before: 'Mar' } },
  { id: 'move-before-very', book: 'very', op: { op: 'moveSheet', sheet: 'Sum', before: 'Feb' } },
  { id: 'move-hidden-first', book: 'feb', op: { op: 'moveSheet', sheet: 'Feb', before: 'Jan' } },
  { id: 'move-before-hidden-last', book: 'sumHidden', op: { op: 'moveSheet', sheet: 'Jan', before: 'Sum' } },
  { id: 'move-end-after-very', book: 'sumVery', op: { op: 'moveSheet', sheet: 'Jan' } },
  { id: 'insert-before-second', book: 'two', op: { op: 'insertSheet', name: 'New', before: 'Mar' } },
  { id: 'insert-before-very', book: 'very', op: { op: 'insertSheet', name: 'New', before: 'Feb' } },
  { id: 'insert-end-after-very', book: 'sumVery', op: { op: 'insertSheet', name: 'New' } },
]

export function buildHiddenCases(): HiddenCases {
  return { books: BOOKS, cases: CASES, read: { sheet: 'Sum', cells: ['A1', 'A2', 'A3'] } }
}
