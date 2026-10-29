/**
 * The `./csv` subpath: a CSV file read into a book of one sheet as Excel opens it (`readCsv`), and a sheet of a book
 * written as a CSV file as Excel saves it (`writeCsv`), with no dependencies and none of the computation. See
 * `docs/csv.md`. Each refuses options it does not take (`strict-args.ts`).
 */
import { withDeclaredArities } from './declared-arity.js'
import { strict } from './strict-args.js'
import * as read from './csv/read.js'
import * as write from './csv/write.js'

// without the engine, the check of a formula's calls counts by the functions' declared arguments
export const readCsv = /* @__PURE__ */ strict(/* @__PURE__ */ withDeclaredArities(read.readCsv))
export const writeCsv = /* @__PURE__ */ strict(write.writeCsv)
export type { CsvIssue, CsvIssueCode, CsvLimits, CsvRead, CsvSnapshot, CsvWritten, ReadCsvOptions, WriteCsvOptions } from './csv/types.js'
