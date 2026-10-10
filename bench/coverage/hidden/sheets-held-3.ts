/**
 * Held out after `sheets-held-2`, whose misses gave two rules: a moved sheet is placed by the visible sheets (after the
 * last visible one where no visible sheet follows the one it goes before, or at the end), and an end of a reference
 * across sheets that moves stays its end until it passes the other end. Here the ends of `'Jan:Mar'`, `'Feb:Apr'` and
 * `'Jan:Apr'` move inward, outward and past each other with every sheet visible, and sheets move past hidden ones at the
 * start and the end. `sheets-held-3.predicted.json`, the engine's answers, was committed before Excel ran the set.
 */
import type { BookInput, SheetInput } from '../../../src/contract.js'
import type { HiddenCase, HiddenCases } from './cases.js'

const NAMES = ['Jan', 'Feb', 'Mar', 'Apr'] as const

const SUM: SheetInput = {
  name: 'Sum',
  cells: {
    A1: { formula: "IFERROR(SUM('Jan:Mar'!B2),\"ref\")" },
    A2: { formula: "IFERROR(SUM('Feb:Apr'!B2),\"ref\")" },
    A3: { formula: "IFERROR(SUM('Jan:Apr'!B2),\"ref\")" },
  },
}

/** Jan to Apr (B2 1, 10, 100, 1000) and Sum, each sheet's `hidden` as given. */
function months(hidden: Partial<Record<(typeof NAMES)[number] | 'Sum', true | 'very'>> = {}): BookInput {
  const h = (name: keyof typeof hidden) => (hidden[name] ? { hidden: hidden[name] } : {})
  return { sheets: [...NAMES.map((name, i) => ({ name, cells: { B2: { value: 10 ** i } }, ...h(name) })), { ...SUM, ...h('Sum') }] }
}

const BOOKS: Record<string, BookInput> = {
  plain: months(),
  tail: months({ Apr: true, Sum: true }),
  head: months({ Jan: true }),
  middle: months({ Mar: true, Apr: true }),
  middleTail: months({ Mar: true, Apr: true, Sum: true }),
  sumVery: months({ Sum: 'very' }),
}

const CASES: HiddenCase[] = [
  { id: 'end-high-inward', book: 'plain', op: { op: 'moveSheet', sheet: 'Mar', before: 'Feb' } },
  { id: 'end-low-outward', book: 'plain', op: { op: 'moveSheet', sheet: 'Feb', before: 'Jan' } },
  { id: 'end-high-past-low', book: 'plain', op: { op: 'moveSheet', sheet: 'Apr', before: 'Jan' } },
  { id: 'end-low-to-end', book: 'plain', op: { op: 'moveSheet', sheet: 'Jan' } },
  { id: 'end-high-between', book: 'plain', op: { op: 'moveSheet', sheet: 'Apr', before: 'Feb' } },
  { id: 'end-past-hidden-tail', book: 'tail', op: { op: 'moveSheet', sheet: 'Jan' } },
  { id: 'before-hidden-first', book: 'head', op: { op: 'moveSheet', sheet: 'Apr', before: 'Jan' } },
  { id: 'hidden-to-end', book: 'middle', op: { op: 'moveSheet', sheet: 'Mar' } },
  { id: 'hidden-to-end-hidden-tail', book: 'middleTail', op: { op: 'moveSheet', sheet: 'Mar' } },
  { id: 'before-very-last', book: 'sumVery', op: { op: 'moveSheet', sheet: 'Feb', before: 'Sum' } },
]

export function buildHiddenCases(): HiddenCases {
  return { books: BOOKS, cases: CASES, read: { sheet: 'Sum', cells: ['A1', 'A2', 'A3'] } }
}
