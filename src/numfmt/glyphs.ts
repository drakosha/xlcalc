/**
 * Widths of shown text as Excel for Mac lays out Calibri 11, the default font of a book: in points of a column, a
 * digit is 6. Measured by fills (`test/excel/display/widths.json`): a column of xlsx width `w` is
 * `trunc(6w + 21·6/256)` points wide, the padding takes 5, and text fits when the sum of its characters' widths
 * is at most the rest. Characters not in the table count as a digit (a known difference for other scripts).
 */

/** Widths in points, by characters of equal width. */
const WIDTHS: [number, string][] = [
  [2, " '"],
  [3, '.,-():{}fijlI'],
  [4, '/!rstzJгт'],
  [5, '+#^~<>=?_acegkvxyEFLSTYZавеёзксучхьэяГЕЁЗТ'],
  [6, '0123456789$₽€£¥bdhnopquABCKPRVXбдийлнпрцъАБВКРСУХЧЬЭЯ'],
  [7, 'DGHNOQUмфыДИЙЛНОПЦЪ'],
  [8, '%&wжшщюФЫ'],
  [9, 'mMЖМ'],
  [10, 'WШЩЮ'],
]

const GLYPH = new Map<string, number>(WIDTHS.flatMap(([w, cs]) => [...cs].map((c): [string, number] => [c, w])))

/** The width of a digit, which a character not in the table takes too. */
export const DIGIT = 6

/** The width of one character. */
export const glyph = (ch: string): number => GLYPH.get(ch) ?? DIGIT

/** The width of a text. */
export function textWidth(s: string): number {
  let w = 0
  for (const ch of s) w += glyph(ch)
  return w
}

/** The width of a column without its own: 8 digits and the padding (xlsx `baseColWidth` 8). */
export const DEFAULT_POINTS = 8 * DIGIT + 5

/** Points of a column of xlsx width `w` (characters of a digit with the padding); the default column without one. */
export const columnPoints = (w: number | undefined): number => (w === undefined ? DEFAULT_POINTS : Math.trunc(DIGIT * w + (21 * DIGIT) / 256))

/** The room for text in a cell of `points`: the column less its padding. */
export const room = (points: number): number => points - 5

/** Excel shows at most 255 characters of a formatted number; `#` fills the cell instead. */
export const MAX_SHOWN = 255

/** `#` across a cell that a number does not fit, one per 5 points; General shows one at least. */
export const hashes = (space: number, general = false): string => '#'.repeat(Math.min(MAX_SHOWN, Math.max(general ? 1 : 0, Math.floor(space / 5))))
