/**
 * The geometry of columns and rows shared by formatting and engine-free helpers, measured against Excel for Mac: a stored
 * width's points, the characters Excel's Column Width box shows, and the sizes of rows and outlines.
 */

/** A stored width's points. */
export const columnPoints = (width: number | undefined): number => width === undefined ? 53 : Math.trunc(6 * width + 126 / 256)

/** A stored width in 256ths of a digit, without clipping or validation for internal callers. */
export const widthOf = (points: number): number => Math.floor(points * 256 / 6) / 256

/**
 * The width Excel for Mac's Column Width box shows for a column of `points`, in characters of the book's default font
 * (Calibri 11), to two decimals: 6 points a character past a margin of 5, and under one character 11 points a character
 * (53 points, the default, are 8).
 */
export const charsOfPoints = (points: number): number => Math.round((points < 11 ? points / 11 : (points - 5) / 6) * 100) / 100

/** The whole points of a column Excel for Mac's Column Width box makes of a width in characters. */
export const pointsOfChars = (chars: number): number => Math.round(chars < 1 ? chars * 11 : chars * 6 + 5)

/** The height of a row without its own, in points: a line of Calibri 11, the book's default font. */
export const DEFAULT_ROW_HEIGHT = 15

/** The highest row Excel makes, in points. */
export const MAX_ROW_HEIGHT = 409.5

/** The deepest outline of rows or columns Excel makes: seven levels of groups below the sheet's own. */
export const MAX_OUTLINE = 7
