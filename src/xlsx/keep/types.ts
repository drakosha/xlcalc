/** The types of the `./xlsx/keep` subpath: a book written over the xlsx file it was read from. */
import type { WriteXlsxOptions, XlsxLimits } from '../types.js'

/**
 * Options of `writeKept`: the language of its issues and of its refusals (`KeptRefusal`) and which formula values the
 * file stores, as `writeXlsx` takes them, and the limits of reading the original file, as `readXlsx` takes them.
 */
export interface WriteKeptOptions extends Pick<WriteXlsxOptions, 'locale' | 'values'> {
  /** How large the original file's parts may be once inflated (`ReadXlsxOptions.limits`). */
  limits?: XlsxLimits
  /**
   * Stores the values of the formulas that differ from the file's also when nothing else changed, where the original bytes
   * come back without it: for a file whose values a program that does not calculate will read (one a generator wrote
   * without them). With no value that differs, the original bytes come back all the same.
   */
  writeValues?: boolean
}

/** Options of `inspectKept`. */
export interface InspectKeptOptions {
  /** How large the file's parts may be once inflated (`ReadXlsxOptions.limits`). */
  limits?: XlsxLimits
}

/** A worksheet of a file as `inspectKept` gives it. */
export interface KeptSheet {
  /** Its name as the file has it. */
  name: string
  /** Its part in the archive (`xl/worksheets/sheet1.xml`). */
  part: string
  /** The cells whose text has formatted runs, which the workbook does not hold: written over, they lose them. */
  formattedText: string[]
}

/** What `inspectKept` finds. */
export interface KeptInspection {
  /** The worksheets, in the workbook's order. */
  sheets: KeptSheet[]
}
