/**
 * The `./xlsx` subpath: an xlsx file read into a book (`readXlsx`) and a book written as an xlsx file with the values
 * of its formulas (`writeXlsx`), with no dependencies. Neither computes: reading gives the values Excel stored
 * (`cached`, in the shape of a result) beside the book, writing takes a snapshot of a book and its result (`compute`'s,
 * a session's `getSnapshot()`, or `cached`). See `docs/xlsx.md`. `fromLegacy` turns a file's formula of the old kind
 * into the dynamic formula Excel 365 shows, as `readXlsx` does (`docs/legacy.md`).
 */
export { fromLegacy, type LegacyContext } from './legacy/convert.js'
export { readXlsx } from './xlsx/read/book.js'
export { writeXlsx } from './xlsx/write/book.js'
export type { WriteXlsxOptions, XlsxCached, XlsxCachedCell, XlsxIssue, XlsxIssueCode, XlsxRead, XlsxSnapshot, XlsxWritten } from './xlsx/types.js'
