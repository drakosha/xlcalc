/**
 * The geometry of columns and rows shared by formatting and engine-free helpers, measured against Excel for Mac: a stored
 * width's points, the characters Excel's Column Width box shows, and the sizes of rows and outlines.
 */

import type { BookInput } from './contract.js'

export type BookGeometry = readonly [digit: number, height: number, size: number]

export function fontFamily(name: string | undefined): 'sans' | 'serif' | undefined {
  switch (name?.toLowerCase()) {
    case 'arial': case 'arial cyr': return 'sans'
    case 'times new roman': return 'serif'
  }
}

/** The Normal font's digit width, row height and point size as measured in Excel for Mac. */
export function normalGeometry(font: BookInput['font']): BookGeometry | undefined {
  const family = fontFamily(font?.name)
  if (font && family) {
    if (font.size === 8) return [4, 11, 8]
    if (font.size === 10) return [family === 'serif' ? 5 : 6, 13, 10]
    if (font.size === 11) return [6, 14, 11]
  }
}

export const bookGeometry = (font: BookInput['font']): BookGeometry => normalGeometry(font) ?? [6, 15, 11]

/** The margin Excel adds to a column of at least one character. */
export const columnPadding = (digit: number): number => digit === 4 ? 3 : 5

/** A stored width's points, the default column eight characters wide. */
export const columnPoints = (width: number | undefined, digit = 6): number =>
  width === undefined ? 8 * digit + columnPadding(digit) : Math.trunc(digit * width + Math.trunc(128 / digit) * digit / 256)

/**
 * The points of the widest stored width Excel for Mac writes (`MAX_FILE_WIDTH` of `text.ts`), `columnPoints(255.83203125)`.
 * Literals only in this module: a computed or long fractional constant here keeps it whole in `xlcalc/a1`, where esbuild
 * otherwise puts `MAX_OUTLINE` in place.
 */
export const MAX_POINTS = 1535

/** A stored width in 256ths of a digit, without clipping or validation for internal callers. */
export const widthOf = (points: number, digit = 6): number => Math.floor(points * 256 / digit) / 256

/** The width Excel for Mac's Column Width box shows, to two decimals, in the book's Normal font. */
export function charsOfPoints(points: number, digit = 6): number {
  const padding = columnPadding(digit)
  const first = digit + padding
  return Math.round((points < first ? points / first : (points - padding) / digit) * 100) / 100
}

/** The whole points of a column Excel for Mac's Column Width box makes of a width in characters. */
export function pointsOfChars(chars: number, digit = 6): number {
  const padding = columnPadding(digit)
  return Math.round(chars < 1 ? chars * (digit + padding) : chars * digit + padding)
}

/** The height of a row without a height of its own, in points: one line of Calibri 11, the workbook's default font. */
export const DEFAULT_ROW_HEIGHT = 15

/** The maximum row height Excel allows, in points. */
export const MAX_ROW_HEIGHT = 409.5

/** The deepest outline of rows or columns Excel allows: seven levels of groups below the sheet's own level. */
export const MAX_OUTLINE = 7
