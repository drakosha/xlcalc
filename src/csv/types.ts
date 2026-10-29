/** The types of the `./csv` subpath: a CSV file read into a book of one sheet, and a sheet of a book written as a CSV file. */
import type { Book, BookInput, ComputeResult } from '../contract.js'
import type { Locale } from '../locale/types.js'
import type { XlsxCached } from '../xlsx/types.js'

/**
 * What the book holds otherwise than Excel would, or what writing lost. The codes:
 *
 * - `formula`: fields Excel takes as formulas (`=A1*2`, `+A1`, `-A1`), kept as text, as `formulas` is not true; or,
 *   with `formulas: true`, fields whose formula does not parse, kept as text.
 * - `textLength`: a field longer than Excel's 32,767 characters, truncated.
 * - `size`: rows past row 1,048,576 or fields past column XFD, left out, as Excel leaves them.
 * - `sheets`: when writing, the sheets other than the one written.
 * - `encoding`: when writing, characters the code page lacks, written as `?`.
 */
export type CsvIssueCode = 'formula' | 'textLength' | 'size' | 'sheets' | 'encoding'

/** One thing lost or changed, and where: the sheet, and the cells it concerns. */
export interface CsvIssue {
  code: CsvIssueCode
  /** What was lost or changed, and why, in the locale of the call. */
  message: string
  sheet?: string
  cells?: string[]
}

/** Options of `readCsv`. */
export interface ReadCsvOptions {
  /**
   * The book's locale: each field is read as text typed into a cell in it (`1 234,5`, `15.03.2026` and `ИСТИНА` in
   * ru-RU), and it gives the list separator of `delimiter: 'locale'`, the code page of a file in no Unicode encoding and
   * the language of the issues. A loaded locale's tag or a locale object, en-US by default; a book in another locale
   * takes its tag (`BookInput.locale`).
   */
  locale?: string | Locale
  /**
   * The file's name (`sales.csv`): the sheet is named after it without its extension, as Excel names it (a name Excel
   * does not take repaired), and a `.tsv` file is read at tabs. `Sheet1` without one.
   */
  name?: string
  /**
   * The character between fields: one character other than a quote or a line break, or `'locale'`, the locale's list
   * separator as Excel takes it (`,` in en-US, `;` in ru-RU). A first line `sep=X` names the delimiter in any case and is
   * no row, as in Excel. Without one: tabs for a `.tsv` name, else a guess from the first lines (`CsvRead.delimiter`).
   */
  delimiter?: string
  /**
   * The file's encoding, by its WHATWG label (`utf-8`, `utf-16le`, `windows-1251`), when the host knows it. Without one:
   * a byte order mark, UTF-16 by its zero bytes, UTF-8 if the bytes are valid UTF-8, else the locale's Windows code page.
   */
  encoding?: string
  /**
   * Whether a field Excel takes as a formula (`=SUM(A1:A3)`, `+A1`, `-A1`, in the locale's formula text) is a formula,
   * as Excel opens a CSV file; false by default, so a file from elsewhere does not bring formulas in (CSV injection):
   * such a field stays text, and an issue lists them.
   */
  formulas?: boolean
  /**
   * Whether a field of digits alone that a number would change stays text: a leading zero (`00123`, a code, an article
   * number) or more than 15 digits (`1234567890123456789`). False by default, as Excel reads them as numbers.
   */
  codesAsText?: boolean
  /** The moment whose year completes a date without one (`15.3`), as `parseEntry` takes it: now by default. */
  now?: string
  /** What the file may hold, beyond which it is refused. */
  limits?: CsvLimits
}

/** The limits `readCsv` sets on a file (`ReadCsvOptions.limits`). */
export interface CsvLimits {
  /**
   * The fields that are not empty, a whole number of at least 1: a file with more is refused with an error before its
   * book is made. Unlimited by default; a host that reads files uploaded by users should set it.
   */
  cells?: number
}

/** A CSV file read into a book. */
export interface CsvRead {
  /** The book: one sheet, its cells as Excel opens the file, in the locale of the read. */
  book: Book
  issues: CsvIssue[]
  /** The delimiter the fields were read at. */
  delimiter: string
  /** The encoding the file was read in, by its WHATWG name (`utf-8`, `utf-16le`, `windows-1251`); null for text. */
  encoding: string | null
}

/** What `writeCsv` writes: a book and the values of its formulas, as `writeXlsx` takes them (a session's `getSnapshot()`, `{ book, result: compute(book) }`). */
export interface CsvSnapshot {
  book: BookInput
  /** The values of the book's formulas and their arrays; needed when the sheet written has formulas. */
  result?: ComputeResult | XlsxCached
}

/** Options of `writeCsv`. */
export interface WriteCsvOptions {
  /** The sheet written, by name: the book's active sheet by default. */
  sheet?: string
  /** The character between fields, as `readCsv` takes it: the list separator of the locale (`'locale'`) by default, as Excel writes. */
  delimiter?: string
  /**
   * `utf-8` (the default, as Excel's CSV UTF-8), `utf-16le`, or `ansi`, the locale's Windows code page (Excel's CSV),
   * which writes a character it lacks as `?`.
   */
  encoding?: 'utf-8' | 'utf-16le' | 'ansi'
  /** Whether a UTF-8 or UTF-16 file begins with its byte order mark: true by default, as Excel writes it. */
  bom?: boolean
  /** The locale cells show their values in and of `delimiter: 'locale'`, and of the issues: the book's by default. */
  locale?: string | Locale
}

/** What `writeCsv` returns: the file, its text, and what the file could not hold. */
export interface CsvWritten {
  data: Uint8Array
  text: string
  issues: CsvIssue[]
}
