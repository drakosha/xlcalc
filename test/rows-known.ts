/**
 * The cases of the rows sets (`test/excel/rows/`, `test/excel/win/rows/`) where the engine knowingly differs from Excel,
 * in the row's height or in its cells, by `<tag>:<set>:<id>`, with the reason.
 */

export const ROWS_REASONS = {
  font: 'a font other than Calibri (Arial, Times New Roman, Courier New): the engine fits a line of any font as a line of Calibri of its size',
  formulaText: 'a formula typed or pasted over lines: Excel keeps its line breaks in the formula, the engine its canonical text without them (the value is the same)',
} as const

export type RowsReason = keyof typeof ROWS_REASONS

/** `<tag>:<set>` and its case ids → the same reason. */
function cases(reason: RowsReason, bySet: Record<string, string[]>): [string, RowsReason][] {
  return Object.entries(bySet).flatMap(([set, ids]) => ids.map((id): [string, RowsReason] => [`${set}:${id}`, reason]))
}

/** Against Excel for Mac. */
export const ROWS_KNOWN: ReadonlyMap<string, RowsReason> = new Map([
  ...cases('font', { 'en-US:autofit': ['font-7', 'lines-11'], 'ru-RU:autofit': ['font-7', 'lines-11'], 'en-US:held': ['fonts-1', 'fonts-2', 'fonts-6'], 'ru-RU:held': ['fonts-1', 'fonts-2', 'fonts-6'] }),
  ...cases('formulaText', { 'en-US:held': ['typed-2'], 'ru-RU:held': ['typed-2'], 'en-US:keys': ['keys-7'], 'en-US:quotes': ['paste-12'] }),
])

/** Against Excel for Windows, the cells alone. */
export const ROWS_WIN_KNOWN: ReadonlyMap<string, RowsReason> = new Map([...cases('formulaText', { 'en-US:held': ['typed-2'], 'en-US:quotes': ['paste-12'], 'ru-RU:quotes-ru': ['paste-12'] })])
