/**
 * The cases of the fit sets (`test/excel/fit/`) where the engine knowingly differs from Excel for Mac, in the column's
 * width or in what its cells show, by `<tag>:<set>:<id>`, with the reason.
 */

export const FIT_REASONS = {
  macRegional:
    "Excel for Mac shows the built-in short date by its own code (en-US m/d/yy, 3/15/26) and ru-RU's month names in lower case and its own short forms (март); the engine follows Excel for Windows (3/15/2026, Март, мар), so the text it fits or widens a column for differs",
  font: "a cell's own font (bold, size 16): Excel fits the text in it, the engine lays out every cell in Calibri 11",
  script: 'characters outside Latin, Cyrillic and the common signs (中文) count as a digit here; Excel lays them out in a font that has them',
  writer:
    "a cell without a format of its own in a column with one: the model shows it by the column's format, writeXlsx writes the cell in General, and Excel shows a stored cell by its own format only",
  advances:
    "text with a character outside ASCII, which Excel for Mac measures by fractional widths: the engine takes them from Calibri's advance widths, which give Excel's width on short text and are a point off on some runs of one Cyrillic letter 600 points wide (19 of 70), on 'Итого по разделу 1' and on ru-RU's long date",
} as const

export type FitReason = keyof typeof FIT_REASONS

/** `<tag>:<set>` and its case ids → the same reason. */
function cases(reason: FitReason, bySet: Record<string, string[]>): [string, FitReason][] {
  return Object.entries(bySet).flatMap(([set, ids]) => ids.map((id): [string, FitReason] => [`${set}:${id}`, reason]))
}

/** The typed cases whose dates Excel for Mac shows by its own short date, in en-US. */
const MAC_DATES = ['date-1', 'formula-1', 'hand-2', 'fitted-2', 'turn-1', 'turn-2'].flatMap((id) => ['value', 'local', 'keys'].map((via) => `${via}-${id}`))

/** Against Excel for Mac. */
export const FIT_KNOWN: ReadonlyMap<string, FitReason> = new Map([
  ...cases('macRegional', {
    'en-US:autofit': ['format-10', 'format-12', 'narrow-3', 'formula-2'],
    'en-US:typed': [...MAC_DATES, 'keys-date-4', 'keys-fitted-1'],
    'en-US:held': ['keys-datetime-1'],
    'ru-RU:autofit': ['format-11', 'format-16', 'format-17', 'many-2'],
    'ru-RU:typed': ['value-date-2', 'local-date-2', 'keys-date-2', 'keys-turn-1'],
    'ru-RU:held': ['number-22', 'rows-1'],
  }),
  ...cases('font', { 'en-US:autofit': ['font-1', 'font-2'], 'ru-RU:autofit': ['font-1', 'font-2'] }),
  ...cases('script', { 'en-US:autofit': ['glyphs-4'], 'ru-RU:autofit': ['glyphs-4'] }),
  ...cases('writer', { 'en-US:autofit': ['colformat-1', 'colformat-2'], 'ru-RU:autofit': ['colformat-1', 'colformat-2'] }),
  ...cases('advances', {
    'en-US:glyphs': ['94', '103', '107', '109', '118', '121', '122', '123', '128', '129', '130', '131', '134', '137', '142', '143', '147', '148', '150'].map((n) => `many-${n}`),
    'en-US:held': ['text-7'],
    'ru-RU:autofit': ['format-18'],
    'ru-RU:held': ['text-7'],
  }),
])
