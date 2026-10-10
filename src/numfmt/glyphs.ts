/**
 * Widths of shown text as Excel for Mac lays out Calibri 11, the default font of a book: in points of a column, a
 * digit is 6. Measured by fills (`test/excel/display/widths.json`): a column of xlsx width `w` is
 * `trunc(6w + 21·6/256)` points wide, the padding takes 5, and text fits when the sum of its characters' widths
 * is at most the rest. The letters of a locale's script are in its file (`glyphWidths`) and count in books of every
 * locale, as text of any script may stand in any book. Characters not in the table count as a digit (a known
 * difference for other scripts).
 */
import { LOCALE_FILES } from '../locale/registry.js'
import { advance } from './advances.js'
import { BLANK, FILL, PAD } from './style.js'

/** Widths in points, by characters of equal width: Latin letters, digits and signs. */
const WIDTHS: [number, string][] = [
  [2, " '"],
  [3, '.,-():{}fijlI'],
  [4, '/!rstzJ'],
  [5, '+#^~<>=?_acegkvxyEFLSTYZ'],
  [6, '0123456789$€£¥bdhnopquABCKPRVX'],
  [7, 'DGHNOQU'],
  [8, '%&w'],
  [9, 'mM'],
  [10, 'W'],
]

const GLYPH = new Map<string, number>(
  [WIDTHS, ...[...LOCALE_FILES.values()].map((f) => f.glyphWidths ?? [])].flat().flatMap(([w, cs]) => [...cs].map((c): [string, number] => [c, w])),
)

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
const DEFAULT_POINTS = 8 * DIGIT + 5

/** Points of a column of xlsx width `w` (characters of a digit with the padding); the default column without one. */
export const columnPoints = (w: number | undefined): number => (w === undefined ? DEFAULT_POINTS : Math.trunc(DIGIT * w + (21 * DIGIT) / 256))

/** The room for text in a cell of `points`: the column less its padding. */
export const room = (points: number): number => points - 5

/**
 * The width of text laid out for AutoFit (`Style.layout`) by `width`, a character's: a fill takes no room, a padding `_x`
 * the room of `x`, a `?` with no digit a digit's room.
 */
function laidOut(s: string, width: (ch: string) => number): number {
  let w = 0
  let mark = ''
  for (const ch of s) {
    if (mark) w += mark === PAD ? width(ch) : 0
    else if (ch !== FILL && ch !== PAD) w += width(ch === BLANK ? '0' : ch)
    mark = !mark && (ch === FILL || ch === PAD) ? ch : ''
  }
  return w
}

/** Text of ASCII alone (the marks of laid-out text aside), which Excel for Mac measures by the widths it shows it in. */
const ASCII = /^[\x20-\x7e\uFFFE\uFFFF\u2007]*$/

/**
 * The points of the column AutoFit gives laid-out text (`test/excel/fit/`): two points past the padding, or from about 90
 * points the padding with a thirtieth more of the text, as Excel for Mac measures text for AutoFit. Text of ASCII alone
 * it measures by the widths it shows it in; other text by the font's advance widths (`advances.ts`), its 11 points
 * rounded for the first measure.
 */
export function fitText(s: string): number {
  if (ASCII.test(s)) {
    const w = laidOut(s, glyph)
    return Math.max(w + 7, Math.floor((w * 31) / 30) + 5)
  }
  const u = laidOut(s, advance)
  return Math.max(Math.floor((u * 11 + 1024) / 2048) + 7, Math.floor((u * 11 * 31) / (2048 * 30)) + 5)
}

/** The xlsx width of a column of `points`, in 256ths of a digit as Excel writes it. */
export const widthOf = (points: number): number => Math.floor((points * 256) / DIGIT) / 256

/** Excel shows at most 255 characters of a formatted number; `#` fills the cell instead. */
export const MAX_SHOWN = 255

/** `#` across a cell that a number does not fit, one per 5 points; General shows one at least. */
export const hashes = (space: number, general = false): string => '#'.repeat(Math.min(MAX_SHOWN, Math.max(general ? 1 : 0, Math.floor(space / 5))))
