/**
 * Hidden and very hidden sheets, held out: books written with them (the first sheet hidden too, where the writer makes
 * the first visible sheet active), sheets hidden and shown by Excel's object model, the last visible sheet hidden and
 * deleted, hidden sheets renamed, deleted and moved, sheets inserted and moved beside them; a reference across sheets
 * over a hidden one and SUBTOTAL and AGGREGATE over a hidden sheet's hidden row. Two files are changed by hand: one whose
 * active tab is its hidden first sheet, one whose every sheet is hidden. `sheets-held.predicted.json`, the engine's
 * answers, was committed before Excel ran the set, with `RAW_GUESS` for the files changed by hand.
 */
import type { BookInput, SheetInput } from '../../../src/contract.js'
import type { HiddenCase, HiddenCases } from './cases.js'

const month = (name: string, k: number, extra: Partial<SheetInput> = {}): SheetInput => ({ name, cells: Object.fromEntries([2, 3, 4, 5].map((r) => [`B${r}`, { value: k * r }])), ...extra })

/** The formulas read: each gives text where an operation lost its reference, as a read of an error stops Excel for Mac. */
const SUM: SheetInput = {
  name: 'Sum',
  cells: {
    A1: { formula: "IFERROR(SUM('Jan:Mar'!B2),\"ref\")" },
    A2: { formula: 'IFERROR(SUBTOTAL(9,Feb!B2:B5),"ref")' },
    A3: { formula: 'IFERROR(SUBTOTAL(109,Feb!B2:B5),"ref")' },
    A4: { formula: 'IFERROR(Feb!B2*2,"ref")' },
    A5: { formula: 'IFERROR(AGGREGATE(9,5,Feb!B2:B5),"ref")' },
    A6: { formula: "IFERROR(COUNT('Jan:Mar'!B2:B5),\"ref\")" },
  },
}

/** Jan, Feb (its row 3 hidden), Mar and Sum, each sheet's `hidden` as given. */
function months(hidden: Partial<Record<'Jan' | 'Feb' | 'Mar' | 'Sum', true | 'very'>> = {}): BookInput {
  const h = (name: keyof typeof hidden) => (hidden[name] ? { hidden: hidden[name] } : {})
  return { sheets: [month('Jan', 1, h('Jan')), month('Feb', 10, { rows: { 3: { hidden: true } }, ...h('Feb') }), month('Mar', 100, h('Mar')), { ...SUM, ...h('Sum') }] }
}

const BOOKS: Record<string, BookInput> = {
  plain: months(),
  feb: months({ Feb: true }),
  febVery: months({ Feb: 'very' }),
  first: months({ Jan: true }),
  firstTwo: months({ Jan: 'very', Feb: true }),
  onlyFeb: months({ Jan: true, Mar: true, Sum: 'very' }),
  lastHidden: months({ Sum: true }),
}

const CASES: HiddenCase[] = [
  // written by writeXlsx
  { id: 'w-plain', book: 'plain' },
  { id: 'w-hidden', book: 'feb' },
  { id: 'w-very', book: 'febVery' },
  { id: 'w-first', book: 'first' },
  { id: 'w-first-two', book: 'firstTwo' },
  { id: 'w-only-feb', book: 'onlyFeb' },
  { id: 'w-last-hidden', book: 'lastHidden' },
  // changed by hand
  { id: 'raw-active-hidden', book: 'first', raw: 'activeHidden' },
  { id: 'raw-all-hidden', book: 'plain', raw: 'allHidden' },
  // hidden and shown by Excel's object model
  { id: 'm-hide', book: 'plain', make: [{ sheet: 'Feb', visible: 0 }] },
  { id: 'm-very', book: 'plain', make: [{ sheet: 'Feb', visible: 2 }] },
  { id: 'm-unhide', book: 'feb', make: [{ sheet: 'Feb', visible: -1 }] },
  { id: 'm-unhide-very', book: 'febVery', make: [{ sheet: 'Feb', visible: -1 }] },
  { id: 'm-very-to-hidden', book: 'febVery', make: [{ sheet: 'Feb', visible: 0 }] },
  { id: 'm-hide-active', book: 'plain', make: [{ sheet: 'Jan', visible: 0 }] },
  { id: 'm-hide-three', book: 'plain', make: [{ sheet: 'Jan', visible: 0 }, { sheet: 'Feb', visible: 2 }, { sheet: 'Mar', visible: 0 }] },
  { id: 'm-hide-last', book: 'onlyFeb', make: [{ sheet: 'Feb', visible: 0 }] },
  { id: 'm-very-last', book: 'onlyFeb', make: [{ sheet: 'Feb', visible: 2 }] },
  // sheet operations beside hidden sheets
  { id: 'o-rename', book: 'feb', op: { op: 'renameSheet', sheet: 'Feb', name: 'February' } },
  { id: 'o-rename-very', book: 'febVery', op: { op: 'renameSheet', sheet: 'Feb', name: 'February' } },
  { id: 'o-delete-hidden', book: 'feb', op: { op: 'deleteSheet', sheet: 'Feb' } },
  { id: 'o-delete-very', book: 'febVery', op: { op: 'deleteSheet', sheet: 'Feb' } },
  { id: 'o-delete-last', book: 'onlyFeb', op: { op: 'deleteSheet', sheet: 'Feb' } },
  { id: 'o-delete-visible', book: 'feb', op: { op: 'deleteSheet', sheet: 'Jan' } },
  { id: 'o-move-hidden', book: 'feb', op: { op: 'moveSheet', sheet: 'Feb' } },
  { id: 'o-move-very', book: 'febVery', op: { op: 'moveSheet', sheet: 'Feb', before: 'Jan' } },
  { id: 'o-move-before-hidden', book: 'feb', op: { op: 'moveSheet', sheet: 'Sum', before: 'Feb' } },
  { id: 'o-insert-before-hidden', book: 'feb', op: { op: 'insertSheet', name: 'New', before: 'Feb' } },
  { id: 'o-insert-end-hidden', book: 'lastHidden', op: { op: 'insertSheet', name: 'New' } },
]

export function buildHiddenCases(): HiddenCases {
  return { books: BOOKS, cases: CASES, read: { sheet: 'Sum', cells: ['A1', 'A2', 'A3', 'A4', 'A5', 'A6'] } }
}

/**
 * What Excel was expected to do with the files changed by hand, written before it ran them: open the one whose active tab
 * is a hidden sheet without repairing it, showing its first visible sheet, and repair the one whose every sheet is hidden.
 */
export const RAW_GUESS = { 'raw-active-hidden': { repaired: false, opened: 'Feb' }, 'raw-all-hidden': { repaired: true } } as const

/**
 * The predictions Excel for Mac and for Windows did not confirm: Excel deletes and moves no very hidden sheet (the object
 * model's error; Excel for Mac deletes nothing and answers nothing); a sheet moved before a hidden one goes after it,
 * before the next visible sheet; a file whose active tab is a hidden sheet opens with that sheet shown.
 */
export const MISSED: readonly string[] = ['raw-active-hidden', 'o-delete-very', 'o-move-very', 'o-move-before-hidden']
