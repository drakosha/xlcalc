/**
 * CSV files for Excel (`run.ts csv <set>`). A case is a CSV file, its text in an encoding, under a name, which Excel
 * opens as a person opens it (`open workbook` of a `.csv` file) and saves as xlsx, then as CSV; or a book, written by
 * `writeXlsx` with its values, which Excel saves as CSV. The answer is the sheet Excel named and the CSV it saved; the
 * xlsx it saved goes beside it, read back by `readXlsx` for each cell's value or formula and format.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import { readCsv, writeCsv } from '../../../src/csv.js'
import { displayRules } from '../../../src/locale/rules.js'
import { compute } from '../../../src/index.js'
import { encodeCsv } from '../../../src/csv/decode.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { ruRU } from '../../../src/locales/ru-RU.js'

/** How a case's text is written: UTF-8 with or without its mark, UTF-16 with its mark or UTF-16LE without, or a code page. */
export type CsvEncoding = 'utf-8' | 'utf-8-bom' | 'utf-16le-bom' | 'utf-16be-bom' | 'utf-16le' | 'windows-1251' | 'windows-1252'

export interface CsvCase {
  id: string
  /** The file's text and its encoding (`utf-8` by default). */
  text?: string
  encoding?: CsvEncoding
  /** The file's name, `book.csv` by default: Excel names the sheet after it. */
  name?: string
  /** Or a book Excel saves as CSV, written as xlsx with the engine's values. */
  book?: BookInput
}

export interface CsvCases {
  /** The locale Excel runs in, en-US by default. */
  locale?: string
  cases: CsvCase[]
}

export interface CsvSet {
  buildCsvCases(): CsvCases
}

/** What Excel did with a case's file: the active sheet's name as it opened it, and the CSV it saved, as base64. */
export interface CsvAnswer {
  sheet: string
  csv: string
}

/** A case's file name as Excel is to open it. */
export const fileName = (c: CsvCase): string => (c.book ? 'book.xlsx' : (c.name ?? 'book.csv'))

/** The bytes of a text in an encoding. */
export function encoded(text: string, encoding: CsvEncoding = 'utf-8'): Uint8Array {
  if (encoding === 'utf-8' || encoding === 'utf-8-bom') return encodeCsv(text, 'utf-8', encoding === 'utf-8-bom', displayRules(undefined).codePage).data
  if (encoding === 'utf-16le' || encoding === 'utf-16le-bom') return encodeCsv(text, 'utf-16le', encoding === 'utf-16le-bom', displayRules(undefined).codePage).data
  if (encoding === 'utf-16be-bom') {
    const le = encodeCsv(text, 'utf-16le', true, displayRules(undefined).codePage).data
    return le.map((_, i) => le[i ^ 1])
  }
  const page = displayRules(encoding === 'windows-1251' ? ruRU : undefined).codePage
  const { data, lost } = encodeCsv(text, 'ansi', false, page)
  if (lost) throw new Error(`${JSON.stringify(text)} has characters ${encoding} lacks`)
  return data
}

/** A case's file: its text encoded, or its book as xlsx with the engine's values. */
export async function caseFile(c: CsvCase): Promise<Uint8Array> {
  if (!c.book) return encoded(c.text ?? '', c.encoding)
  const { data, issues } = await writeXlsx({ book: c.book, result: compute(c.book) })
  if (issues.length) throw new Error(`${c.id}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
  return data
}

/** What the engine makes of a case: for a CSV file the sheet and its cells as `readCsv` reads it as Excel does, and the CSV `writeCsv` writes of a case's book (in the code page, as Excel's CSV). */
export interface CsvEngine {
  sheet?: string
  cells?: Record<string, CellInput>
  styles?: BookInput['styles']
  csv: string
}

/** The engine's answer to a case in `locale`, a date without a year taking the year of `now`. */
export function engineAnswer(c: CsvCase, data: Uint8Array, locale: string, now: string): CsvEngine {
  const loc = locale === 'ru-RU' ? ruRU : undefined
  const read = c.book ? undefined : readCsv(data, { locale: loc, name: fileName(c), delimiter: 'locale', formulas: true, now })
  const b = c.book ?? read!.book
  const csv = writeCsv({ book: b, result: compute(b) }, { encoding: 'ansi', ...(loc ? { locale: loc } : {}) }).text
  return read ? { sheet: read.book.sheets[0].name, cells: read.book.sheets[0].cells, ...(read.book.styles ? { styles: read.book.styles } : {}), csv } : { csv }
}
