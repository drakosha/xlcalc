/**
 * Held out, predicted by `Session.pasteText` (no percent editor) before Excel ran it: a cell edited in place, as a person
 * edits it after F2, a double click or in the formula bar (the cell's editing begun with F2, what it opened with erased
 * or only a digit changed, then Return). Excel's percent editor adds `%` to a number typed over a percent cell
 * (`docs/entry.md`, "A percent cell"); the question is whether it does so in an edit: a percent cell retyped with its
 * sign and without it, with only its digit changed, with a decimal in both locales' forms, a negative number; an empty
 * percent cell and a percent cell with a formula or text; a cell in General as the control; a number with a sign or in
 * parentheses (`-5` took the percent in the first run). Typed on the keyboard:
 * ASCII only. Both locales' forms are in the set, which runs in each.
 */
import { collector, type UiCase } from './cases.js'
import type { CellInput } from '../../../src/contract.js'

const PCT: CellInput = { value: 0.05, numberFormat: '0%' }
const DEC: CellInput = { value: 0.125, numberFormat: '0.0%' }

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const edit = (group: string, cell: CellInput, text: string, keys?: { left: number; back: number; type: string }) =>
    add(group, { cells: { B2: cell }, act: { do: 'edit', at: 'B2', text, ...(keys ? { keys } : {}) } })

  edit('percent', PCT, '6%')
  edit('percent', PCT, '6')
  // `5%` opens with the cursor after `%`: Left, Backspace and `6` leave `6%`
  edit('percent', PCT, '6%', { left: 1, back: 1, type: '6' })
  edit('percent', PCT, '0.5')
  edit('percent', PCT, '-5')
  edit('decimal', DEC, '7.5%')
  edit('decimal', DEC, '7,5%')
  edit('decimal', DEC, '7.5')
  edit('decimal', DEC, '7,5')
  edit('empty', { numberFormat: '0%' }, '6')
  edit('formula', { formula: '0.05', numberFormat: '0%' }, '6')
  edit('text', { value: 'abc', numberFormat: '0%' }, '6')
  edit('general', { value: 5 }, '6%')
  // `-5` took the percent in a first run, `6` did not: a sign and the forms around it
  edit('sign', PCT, '-0.5')
  edit('sign', PCT, '-50')
  edit('sign', PCT, '+5')
  edit('sign', PCT, '(5)')
  edit('sign', PCT, '-5%')
  edit('sign', { value: 5 }, '-5')
  edit('sign', { numberFormat: '0%' }, '-5')
  return out
}
