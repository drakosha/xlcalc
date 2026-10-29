/**
 * The cases of the rows sets (`test/excel/rows/`, `test/excel/win/rows/`) where the engine knowingly differs from Excel,
 * in the row's height or in its cells, by `<tag>:<set>:<id>`, with the reason.
 */

import { knownCases as cases } from './known-cases.js'

export const ROWS_REASONS = {
  font: 'a font other than Calibri (Arial, Times New Roman, Courier New): the engine fits a line of any font as a line of Calibri of its size',
  formulaText: 'a formula typed or pasted over lines: Excel keeps its line breaks in the formula, the engine its canonical text without them (the value is the same)',
  macDate: "Excel for Mac shows the built-in short date by its own code (m/d/yy, 3/15/26), the engine as Excel for Windows does (3/15/2026), so the turned text it fits the row to is longer",
  turnedWrap:
    'wrapped text turned: Excel wraps text turned by 90 degrees at the height the row had (409.5 points as the book opened, 401 after AutoFit) and gives wrapped letters one above the other two points more; the engine fits wrapped text turned up or down to nothing and its letters by its longest line',
  patchedEmpty:
    "a font size set on whole columns: Excel fits every row of the sheet to the font, those without a cell too, and saves none of them; the engine fits the rows that hold a cell, as a height over every row would be written into the file row by row",
} as const

export type RowsReason = keyof typeof ROWS_REASONS

/** Against Excel for Mac. */
export const ROWS_KNOWN: ReadonlyMap<string, RowsReason> = new Map([
  ...cases('font', { 'en-US:autofit': ['font-7', 'lines-11'], 'ru-RU:autofit': ['font-7', 'lines-11'], 'en-US:held': ['fonts-1', 'fonts-2', 'fonts-6'], 'ru-RU:held': ['fonts-1', 'fonts-2', 'fonts-6'] }),
  ...cases('formulaText', { 'en-US:held': ['typed-2'], 'ru-RU:held': ['typed-2'], 'en-US:keys': ['keys-7'], 'en-US:quotes': ['paste-12'] }),
  ...cases('macDate', { 'en-US:rotation': ['value-2', 'value-6'] }),
  ...cases('turnedWrap', { 'en-US:rotation': ['wrap-2', 'wrap-3', 'wrap-7', 'wrap-10'] }),
  ...cases('patchedEmpty', { 'en-US:patched': ['size-1', 'size-2', 'size-3', 'size-4'] }),
])

/** Against Excel for Windows, the cells alone. */
export const ROWS_WIN_KNOWN: ReadonlyMap<string, RowsReason> = new Map([...cases('formulaText', { 'en-US:held': ['typed-2'], 'en-US:quotes': ['paste-12'], 'ru-RU:quotes-ru': ['paste-12'] })])
