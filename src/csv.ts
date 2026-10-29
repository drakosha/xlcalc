/**
 * The `./csv` subpath: a CSV file read into a book of one sheet as Excel opens it (`readCsv`), and a sheet of a book
 * written as a CSV file as Excel saves it (`writeCsv`), with no dependencies and none of the computation. See
 * `docs/csv.md`. Each refuses options it does not take (`strict-args.ts`).
 */
import { strict } from './strict-args.js'
import * as read from './csv/read.js'
import * as write from './csv/write.js'

export const readCsv = /* @__PURE__ */ strict(read.readCsv)
export const writeCsv = /* @__PURE__ */ strict(write.writeCsv)
export type { CsvIssue, CsvIssueCode, CsvLimits, CsvRead, CsvSnapshot, CsvWritten, ReadCsvOptions, WriteCsvOptions } from './csv/types.js'
