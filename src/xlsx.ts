/**
 * The `./xlsx` subpath: an xlsx file read into a book (`readXlsx`) and a book written as an xlsx file with its
 * computed values (`writeXlsx`), with no dependencies. Neither computes: reading gives the values Excel stored
 * (`cached`) beside the book, writing takes a result of `compute` or a session. See `docs/xlsx.md`. `fromLegacy` turns a
 * file's formula of the old kind into the dynamic formula Excel 365 shows, as `readXlsx` does (`docs/legacy.md`).
 */
export { fromLegacy, type LegacyContext } from './legacy/convert.js'
export { readXlsx } from './xlsx/read/book.js'
export { writeXlsx } from './xlsx/write/book.js'
export type { WriteXlsxOptions, XlsxIssue, XlsxIssueCode, XlsxRead, XlsxWritten } from './xlsx/types.js'
