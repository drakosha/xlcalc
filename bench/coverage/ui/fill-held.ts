/**
 * Held out, predicted by `Session.setCellTexts` and `previewCellTexts` before Excel ran it: a text typed into a selection
 * and entered into each of its cells (Control+Return), the selection's first cell active. A number, a long number (does
 * the column widen), formulas with relative and absolute references over a column and a rectangle (are they shifted),
 * dates and percents of both locales, text after an apostrophe, a selection whose cells have different formats (a percent
 * cell active and not), rules of data validation of the styles information and stop over the selection and over a cell of
 * it that is not the active one, and a formula that does not parse. Typed on the keyboard: ASCII only. Both locales' forms
 * are in the set, which runs in each.
 */
import { collector, DATA, type UiCase } from './cases.js'

const RULE = { type: 'whole', operator: 'greaterThan', formula1: '0' } as const

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const fill = (group: string, range: string, text: string, more: Partial<Omit<UiCase, 'id' | 'act'>> = {}) =>
    add(group, { ...more, cells: { ...DATA, ...more.cells }, act: { do: 'fill', range, text } })

  fill('number', 'B1:B3', '5')
  fill('wide', 'B1:B3', '1234567890123')
  fill('formula', 'B1:B3', '=A1*2')
  fill('formula', 'B1:C2', '=$A$1+A1')
  fill('formula', 'B1:B2', '=A1+A2')
  fill('formula', 'B1:B3', '-A1')
  fill('date', 'B1:B2', '1/31/2025')
  fill('date', 'B1:B2', '31.01.2025')
  fill('percent', 'B1:B2', '12.5%')
  fill('percent', 'B1:B2', '12,5%')
  fill('text', 'B1:B2', "'007")
  fill('text', 'B1:B2', 'abc')
  // the percent cell's editor: the active cell's format, or each cell's
  fill('formats', 'B1:B2', '5', { cells: { B1: { numberFormat: '0%' } } })
  fill('formats', 'B1:B2', '5', { cells: { B2: { numberFormat: '0%' } } })
  fill('valid', 'B1:B3', '-2', { validations: [{ range: 'B1:B3', ...RULE, errorStyle: 'information' }] })
  fill('valid', 'B1:B3', '-2', { validations: [{ range: 'B1:B3', ...RULE }] })
  fill('valid', 'B1:B3', '-2', { validations: [{ range: 'B2', ...RULE }] })
  fill('bad', 'B1:B2', '=1+')
  return out
}
