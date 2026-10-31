/**
 * A pivot cache's dates as the file holds them and as the model does: a time of day without a date, which the model
 * holds at day 0 of the 1900 system (`1900-01-00T08:30:00`, as a cell's serial 0.354 gives it), Excel writes on
 * 1899-12-30, the day 0 of its object model's dates (the set slicer-labels); a 1904 workbook's on 1904-01-01 in both.
 */
const FILE_ZERO = '1899-12-30T'
const MODEL_ZERO = '1900-01-00T'

/** A date of the file in the model's spelling. */
export const modelDate = (d: string): string => (d.startsWith(FILE_ZERO) ? MODEL_ZERO + d.slice(FILE_ZERO.length) : d)

/** A date of the model in the file's spelling. */
export const fileDate = (d: string): string => (d.startsWith(MODEL_ZERO) ? FILE_ZERO + d.slice(MODEL_ZERO.length) : d)
