/**
 * Widths of shown text as Excel for Mac lays out Calibri 11, the default font of a book: in points of a column, a
 * digit is 6. Measured by fills (`test/excel/display/widths.json`): a column of xlsx width `w` is
 * `trunc(6w + 21·6/256)` points wide, the padding takes 5, and text fits when the sum of its characters' widths
 * is at most the rest. The letters of other scripts are here too and count in books of every locale, as text of any
 * script may stand in any book, whatever locales are loaded. Characters not in the table count as a digit (a known
 * difference for other scripts).
 */
import { ADVANCES, BESIDES, DIGIT_ADVANCE } from './advances.js'
import { BLANK, FILL, PAD } from './style.js'
export { columnPoints, widthOf } from '../column-width.js'

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

/** The Russian alphabet and the ruble sign, measured by fills as the Latin letters (`test/excel/display/widths.json`). */
const CYRILLIC: [number, string][] = [
  [4, 'гт'],
  [5, 'авеёзксучхьэяГЕЁЗТ'],
  [6, '₽бдийлнпрцъАБВКРСУХЧЬЭЯ'],
  [7, 'мфыДИЙЛНОПЦЪ'],
  [8, 'жшщюФЫ'],
  [9, 'ЖМ'],
  [10, 'ШЩЮ'],
]

/**
 * Widths by character code, from lists of runs of equal width, up to the highest code they hold: the last run's for a
 * character in two, `other` for a code with none. Read by its code, a character costs no string of its own.
 */
function byCode(lists: [number, string][][], other: number, Table: Uint8ArrayConstructor | Uint16ArrayConstructor): Uint8Array | Uint16Array {
  const codes = lists.flat().map(([w, cs]): [number, number[]] => [w, [...cs].map((c) => c.charCodeAt(0))])
  const table = new Table(Math.max(...codes.flatMap(([, cs]) => cs)) + 1).fill(other)
  for (const [w, cs] of codes) for (const c of cs) table[c] = w
  return table
}

/** The width of a digit, which a character not in the table takes too. */
export const DIGIT = 6

const GLYPH = /* @__PURE__ */ byCode([WIDTHS, CYRILLIC], DIGIT, Uint8Array)

const ADVANCE = /* @__PURE__ */ byCode([ADVANCES, BESIDES], DIGIT_ADVANCE, Uint16Array)

/** The advance of the character of code point `c` in Carlito (`advances.ts`), in 1/2048 of the font's size. */
const advanceOf = (c: number): number => (c < ADVANCE.length ? ADVANCE[c] : DIGIT_ADVANCE)

/** The advance of one character, in 1/2048 of the font's size. */
export const advance = (ch: string): number => (ch.length === 1 ? advanceOf(ch.charCodeAt(0)) : DIGIT_ADVANCE)

/** The width of the character of code point `c`. */
const glyphOf = (c: number): number => (c < GLYPH.length ? GLYPH[c] : DIGIT)

/** The width of one character. */
export const glyph = (ch: string): number => (ch.length === 1 ? glyphOf(ch.charCodeAt(0)) : DIGIT)

/** The width of a text, a pair of surrogates one character. */
export function textWidth(s: string): number {
  let w = 0
  for (let i = 0; i < s.length; i++) {
    const c = s.codePointAt(i)!
    if (c > 0xffff) i++
    w += glyphOf(c)
  }
  return w
}

/** The room for text in a cell of `points`: the column less its padding. */
export const room = (points: number): number => points - 5

/** The codes of the marks of laid-out text, and of the digit a blank takes the room of. */
const FILL_CODE = /* @__PURE__ */ FILL.charCodeAt(0)
const PAD_CODE = /* @__PURE__ */ PAD.charCodeAt(0)
const BLANK_CODE = /* @__PURE__ */ BLANK.charCodeAt(0)
const ZERO_CODE = 0x30

/**
 * The width of text laid out for AutoFit (`Style.layout`) by `width`, a character's: a fill takes no room, a padding `_x`
 * the room of `x`, a `?` with no digit a digit's room.
 */
function laidOut(s: string, width: (c: number) => number): number {
  let w = 0
  let mark = 0
  for (let i = 0; i < s.length; i++) {
    const c = s.codePointAt(i)!
    if (c > 0xffff) i++
    if (mark) w += mark === PAD_CODE ? width(c) : 0
    else if (c !== FILL_CODE && c !== PAD_CODE) w += width(c === BLANK_CODE ? ZERO_CODE : c)
    mark = !mark && (c === FILL_CODE || c === PAD_CODE) ? c : 0
  }
  return w
}

/** The points AutoFit gives `w` points of text: two past the padding, or the padding and a thirtieth more of the text. */
export const padded = (w: number): number => Math.max(w + 7, Math.floor((w * 31) / 30) + 5)

/**
 * The width of laid-out text in Calibri of `size` points as Excel for Mac wraps and turns it: each character the font's
 * advance rounded to a point at its size (`test/excel/rows/held.json`, the set `scale`; `rotation-2.json`).
 */
export const sizedWidth = (s: string, size: number): number => laidOut(s, (c) => Math.round((advanceOf(c) * size) / 2048))

/** Text of ASCII alone (the marks of laid-out text aside), which Excel for Mac measures by the widths it shows it in. */
const ASCII = /^[\x20-\x7e\uFFFE\uFFFF\u2007]*$/

/**
 * The points of the column AutoFit gives laid-out text (`test/excel/fit/`): two points past the padding, or from about 90
 * points the padding with a thirtieth more of the text, as Excel for Mac measures text for AutoFit. Text of ASCII alone
 * it measures by the widths it shows it in; other text by the font's advance widths (`advances.ts`), its 11 points
 * rounded for the first measure.
 */
export function fitText(s: string): number {
  if (ASCII.test(s)) return padded(laidOut(s, glyphOf))
  const u = laidOut(s, advanceOf)
  return Math.max(Math.floor((u * 11 + 1024) / 2048) + 7, Math.floor((u * 11 * 31) / (2048 * 30)) + 5)
}

/** Excel shows at most 255 characters of a formatted number; `#` fills the cell instead. */
export const MAX_SHOWN = 255

/** `#` across a cell that a number does not fit, one per 5 points; General shows one at least. */
export const hashes = (space: number, general = false): string => '#'.repeat(Math.min(MAX_SHOWN, Math.max(general ? 1 : 0, Math.floor(space / 5))))
