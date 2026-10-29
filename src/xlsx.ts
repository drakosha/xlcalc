/**
 * The `./xlsx` subpath: an xlsx file read into a book (`readXlsx`) and a book written as an xlsx file with the values
 * of its formulas (`writeXlsx`), with no dependencies. Neither computes: reading gives the values Excel stored
 * (`cached`, in the shape of a result) beside the book, writing takes a snapshot of a book and its result (`compute`'s,
 * a session's `getSnapshot()`, or `cached`). See `docs/xlsx.md`. `fromLegacy` turns a file's formula of the old kind
 * into the dynamic formula Excel 365 shows, as `readXlsx` does (`docs/legacy.md`). Each refuses options it does not take
 * (`strict-args.ts`).
 */
import { withDeclaredArities } from './declared-arity.js'
import { strict } from './strict-args.js'
import * as read from './xlsx/read/book.js'
import * as write from './xlsx/write/book.js'

export { fromLegacy, type LegacyContext } from './legacy/convert.js'
export const readXlsx = /* @__PURE__ */ strict(read.readXlsx)
// the writer tells a function's name from a defined name's by the functions' declared arguments without the engine
export const writeXlsx = /* @__PURE__ */ strict(/* @__PURE__ */ withDeclaredArities(write.writeXlsx))
export type { ReadXlsxOptions, WriteXlsxOptions, XlsxCached, XlsxCachedCell, XlsxIssue, XlsxIssueCode, XlsxLimits, XlsxRead, XlsxSnapshot, XlsxWritten } from './xlsx/types.js'
