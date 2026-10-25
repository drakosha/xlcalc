/** The types of the `./xlsx/keep` subpath: a book written over the xlsx file it was read from. */
import type { WriteXlsxOptions, XlsxLimits } from '../types.js'

/**
 * Options of `writeKept`: the language of its issues and which formula values the file stores, as `writeXlsx` takes
 * them, and the limits of reading the original file, as `readXlsx` takes them.
 */
export interface WriteKeptOptions extends Pick<WriteXlsxOptions, 'locale' | 'values'> {
  /** How large the original file's parts may be once inflated (`ReadXlsxOptions.limits`). */
  limits?: XlsxLimits
}
