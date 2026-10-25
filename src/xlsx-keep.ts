/**
 * The `./xlsx/keep` subpath: a book written over the xlsx file it was read from, so that what the model does not hold
 * (charts, pictures, pivot tables, the theme, rich text, page setup and the rest) stays as the file has it. `openKept`
 * reads the file as `readXlsx` does, with hidden names that carry the references of those parts; `writeKept` writes
 * what the book changed over the original bytes and refuses (`KeptRefusal`) a change that would leave those references
 * pointing at other cells. See `docs/xlsx-keep.md`. Each refuses options it does not take (`strict-args.ts`).
 */
import { strict } from './strict-args.js'
import * as open from './xlsx/keep/open.js'
import * as write from './xlsx/keep/book.js'

export const openKept = /* @__PURE__ */ strict(open.openKept)
export const writeKept = /* @__PURE__ */ strict(write.writeKept)
export { KeptRefusal } from './xlsx/keep/refusals.js'
export type { WriteKeptOptions } from './xlsx/keep/types.js'
export type { ReadXlsxOptions, XlsxIssue, XlsxIssueCode, XlsxLimits, XlsxRead, XlsxSnapshot, XlsxWritten } from './xlsx/types.js'
