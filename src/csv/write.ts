/**
 * `writeCsv`: a sheet of a book as a CSV file, as Excel saves it. Each cell is written as it shows its value (a value's,
 * or a formula's and its array's from the result) by its number format in the locale, from A1 to the last row and
 * column with a value, every row as wide as the widest; fields are quoted where they need it (`fields.ts`), rows end
 * in CRLF, and the text is encoded as asked (`decode.ts`).
 */
import type { CellInput, CellValue } from '../contract.js'
import { epochOf } from '../dates.js'
import { InputError } from '../input-error.js'
import { checkBoolean, checkOneOf, checkString, object, onlyOptions } from '../input/basic.js'
import { optionalLocale } from '../locale/load.js'
import { localeOf } from '../locale/registry.js'
import { displayRules } from '../locale/rules.js'
import { bookScalar, parseAddress } from '../model.js'
import { bookLooks } from '../shown.js'
import { sheetKey } from '../syntax/print.js'
import { activeSheetIndex } from '../xlsx/view.js'
import { writeArgs } from '../xlsx/write/snapshot.js'
import { encodeCsv, type WriteEncoding } from './decode.js'
import { quoted } from './fields.js'
import { CsvIssues } from './issues.js'
import { checkDelimiter } from './read.js'
import type { CsvSnapshot, CsvWritten, WriteCsvOptions } from './types.js'

const ENCODINGS: readonly WriteEncoding[] = ['utf-8', 'utf-16le', 'ansi']

/** Writes a sheet of a book (its active sheet by default) as a CSV file, its formulas by their values in `snapshot.result`. */
export function writeCsv(snapshot: CsvSnapshot, opts?: WriteCsvOptions): CsvWritten {
  const o = opts === undefined ? {} : object(opts, 'opts', "options such as { sheet: 'Sales', delimiter: ',' }")
  onlyOptions(o, 'opts', ['sheet', 'delimiter', 'encoding', 'bom', 'locale'])
  const has = (snapshot as { result?: unknown } | null)?.result !== undefined
  const { book, keys, cells } = writeArgs(snapshot as never, { values: has ? 'preserve' : 'omit' })
  const named = o.sheet === undefined ? undefined : sheetKey(checkString(o.sheet, 'opts.sheet'))
  const s = named === undefined ? activeSheetIndex(book) : book.sheets.findIndex((x) => sheetKey(x.name) === named)
  if (s < 0) throw new InputError('opts.sheet', `the workbook has no sheet ${JSON.stringify(o.sheet)}`)
  const sheet = book.sheets[s]
  const locale = optionalLocale(o.locale, 'opts.locale') ?? localeOf(book.locale)
  const given = o.delimiter === undefined ? 'locale' : checkDelimiter(o.delimiter, 'opts.delimiter')
  const delimiter = given === 'locale' ? locale.listSeparator : given
  if (o.encoding !== undefined) checkOneOf(o.encoding, 'opts.encoding', ENCODINGS)
  if (o.bom !== undefined) checkBoolean(o.bom, 'opts.bom')
  const epoch = epochOf(book.dateSystem)
  const rules = displayRules(locale, epoch)
  const looks = bookLooks(book, rules)
  // each cell's value: its own, its formula's from the result, or one an array of the result fills
  const values = new Map<string, { v: CellValue; own?: CellInput }>()
  const result = cells?.[s] ?? {}
  for (const a of Object.keys(result)) values.set(a, { v: result[a].value })
  for (const a of keys[s]) {
    const own = sheet.cells[a]
    if (own.formula !== undefined) {
      const r = result[a]
      if (r === undefined) throw new InputError('snapshot.result', `the formula of '${sheet.name}'!${a} has no value: pass { book, result } with the result of this workbook`)
      values.set(a, { v: r.value, own })
    } else if (own.value !== undefined && own.value !== null) values.set(a, { v: bookScalar(own.value, epoch), own })
  }
  const rows: string[][] = []
  let width = 0
  for (const [a, { v, own }] of values) {
    if (v === null) continue
    const { row, col } = parseAddress(a)
    width = Math.max(width, col + 1)
    ;(rows[row] ??= [])[col] = quoted(looks(s, row, col, v, own).text, delimiter)
  }
  let text = ''
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] ?? []
    for (let c = 0; c < width; c++) text += (c ? delimiter : '') + (row[c] ?? '')
    text += '\r\n'
  }
  const issues = new CsvIssues(sheet.name)
  if (book.sheets.length > 1) issues.add('sheets', { key: 'csvSheets', sheet: sheet.name })
  const encoding = (o.encoding as WriteEncoding | undefined) ?? 'utf-8'
  const { data, lost } = encodeCsv(text, encoding, o.bom !== false, rules.codePage)
  if (lost) issues.add('encoding', { key: 'csvCharsLost', page: rules.codePage.name })
  return { data, text, issues: issues.list(locale) }
}
