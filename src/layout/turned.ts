/**
 * Text turned by `alignment.rotation` as Excel for Mac fits rows and columns to it (`test/excel/rows/rotation*.json`,
 * `test/excel/fit/rotation*.json`). A line turned by an angle takes its width's projection and its line's across the
 * other axis; its letters one above the other (`vertical`) take a line each down the row and one letter's line across
 * the column. Text is measured in its font's size, each character the advance rounded to a point (`sizedWidth`), and a
 * turned line keeps a margin of a point for each eighth of its size at each end.
 */
import type { CellStyle } from '../contract.js'
import { padded, sizedWidth } from '../numfmt/glyphs.js'

/** A turn: degrees from -90 to 90, or the letters one above the other. */
export type Turn = number | 'vertical'

/** The turn of a style, none when its text is level. */
export const turnOf = (style: CellStyle | undefined): Turn | undefined => style?.alignment?.rotation || undefined

/** The styles of a book's table that turn their text, by place; null when none does, as in most books. Read once per table. */
const TURNED = new WeakMap<readonly CellStyle[], Set<number> | null>()
export function turnedStyles(styles: readonly CellStyle[] | undefined): Set<number> | null {
  if (!styles) return null
  let x = TURNED.get(styles)
  if (x === undefined) {
    const ids = styles.flatMap((st, i) => (turnOf(st) === undefined ? [] : [i]))
    TURNED.set(styles, (x = ids.length ? new Set(ids) : null))
  }
  return x
}

/** The size of the book's default font, which a style without one shows. */
export const DEFAULT_SIZE = 11

/** The font size of a style, the default's without one. */
export const fontSize = (style: CellStyle | undefined): number => style?.font?.size ?? DEFAULT_SIZE

/** The height of one line of Calibri of `size` points, as Excel for Mac fits a row to it (`test/excel/rows/autofit.json`); a size of a half point counts as the next whole one (`held.json`). */
export const lineHeight = (size: number): number => Math.round(Math.round(size) * 1.27 + 0.96)

/** The margin a turned line keeps at each end, by its font's size: a point for each eighth, from 1 below 9 points. */
const margin = (size: number) => Math.floor((Math.round(size) + 7) / 8)

/** The width of a turned line of `size` points (`sizedWidth`), a line break as wide as a space (`rotation.json`, `wrap`). */
const lineWidth = (text: string, size: number): number => sizedWidth(text.replace(/\n/g, ' '), size)

/** A margin's projection rounded, a half up: a sine of 30 degrees is a hair under a half in floating point (`rotation-held.json`, `size`). */
const projected = (m: number, sin: number): number => Math.round(m * sin + 1e-9)

/** The sine and cosine of an angle of `deg` degrees, its sign aside: a turn up or down takes the same room. */
const trig = (deg: number): [sin: number, cos: number] => {
  const a = (Math.abs(deg) * Math.PI) / 180
  return [Math.sin(a), Math.cos(a)]
}

/**
 * The height a turned text needs of its row, whole (it stands on no line with level cells): its letters a line each
 * and two points, a line break a letter too, or when it wraps its longest line's; at 90 degrees its width as AutoFit
 * pads a column's text (`padded`), with the margins past 11 points'; at another angle the projections of its width and
 * its line, a point, and the margins as they project, two points more for a number shown by a format other than a
 * date's (`rotation-held.json`, `format`), in General too.
 */
export function turnedHeight(text: string, size: number, turn: Turn, number: boolean, wrap = false): number {
  const h = lineHeight(size)
  const m = margin(size)
  if (turn === 'vertical') return Math.max(...(wrap ? text.split('\n') : [text]).map((x) => [...x].length)) * h + 2
  const w = lineWidth(text, size)
  if (Math.abs(turn) === 90) return padded(w) + 2 * m - 4
  const [sin, cos] = trig(turn)
  return Math.round(w * sin + h * cos) + 1 + 2 * projected(m, sin) + (number ? 2 : 0)
}

/**
 * The points of the column AutoFit gives a turned text: at an angle the projections of its width and its lines padded
 * as level text (`padded`), with the margins past 11 points'; at 90 degrees its lines alone, a point less; its letters
 * one above the other one letter's line, two points less. `lines`: the text's lines across the column.
 */
export function turnedPoints(text: string, size: number, turn: Turn, lines = 1): number {
  const h = lineHeight(size) * lines
  const m = 2 * margin(size) - 4
  if (turn === 'vertical') return padded(lineHeight(size)) + m - 2
  if (Math.abs(turn) === 90) return padded(h) + m - 1
  const [sin, cos] = trig(turn)
  return padded(Math.round(lineWidth(text, size) * cos + h * sin)) + m
}
