/** The stored-width geometry shared by formatting and engine-free helpers, measured against Excel for Mac. */
export const columnPoints = (width: number | undefined): number => width === undefined ? 53 : Math.trunc(6 * width + 126 / 256)

/** A stored width in 256ths of a digit, without clipping or validation for internal callers. */
export const widthOf = (points: number): number => Math.floor(points * 256 / 6) / 256
