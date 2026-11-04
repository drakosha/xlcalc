/**
 * The `./xlsx/keep` subpath: a book written over the xlsx file it was read from, so that what the model does not hold
 * (charts, pictures, pivot tables, the theme, rich text, page setup and the rest) stays as the file has it. `openKept`
 * reads the file as `readXlsx` does, with hidden names that carry the references of those parts; `writeKept` writes
 * what the book changed over the original bytes and refuses (`KeptRefusal`) a change that would leave those references
 * pointing at other cells; `inspectKept` says where a file's sheets lie and which of their cells have formatted runs of
 * text. See `docs/xlsx-keep.md`. Each refuses options it does not take (`strict-args.ts`).
 */
import { strict } from './strict-args.js'
import * as inspect from './xlsx/keep/inspect.js'
import * as open from './xlsx/keep/open.js'
import * as write from './xlsx/keep/book.js'

export const openKept = /* @__PURE__ */ strict(open.openKept)
export const writeKept = /* @__PURE__ */ strict(write.writeKept)
export const inspectKept = /* @__PURE__ */ strict(inspect.inspectKept)
export { KeptRefusal, type KeptRefusalCode } from './xlsx/keep/refusals.js'
export type { InspectKeptOptions, KeptInspection, KeptSheet, WriteKeptOptions } from './xlsx/keep/types.js'
export type { ReadXlsxOptions, XlsxIssue, XlsxIssueCode, XlsxLimits, XlsxRead, XlsxSnapshot, XlsxWritten } from './xlsx/types.js'
