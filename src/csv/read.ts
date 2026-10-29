/**
 * `readCsv`: a CSV file as a book of one sheet, as Excel opens it. The text (`decode.ts`) is split into rows and fields
 * (`fields.ts`), and each field is read as text pasted into a cell in the book's locale (`readEntry`, the way of a
 * paste): a number, a date, a time, a percentage or money with the format Excel sets, a boolean, an error, text, or a
 * formula; an apostrophe before text stays part of it, a formula that does not parse is text, and a field with a line
 * feed wraps its cell. Where the book holds
 * a field otherwise than Excel would by the options, an issue says so.
 */
import type { BookInput, CellInput, EntryOptions } from '../contract.js'
import { readEntry } from '../entry/index.js'
import { checkEntryOptions } from '../input.js'
import { InputError, what } from '../input-error.js'
import { checkBoolean, checkNumber, checkString, object, onlyKeys, onlyOptions } from '../input/basic.js'
import { checkLocaleArg } from '../locale/load.js'
import { DEFAULT_LOCALE, localeOf } from '../locale/registry.js'
import { displayRules } from '../locale/rules.js'
import type { Locale } from '../locale/types.js'
import { formatAddress, MAX_COLS, MAX_ROWS, MAX_TEXT } from '../model.js'
import { fileBytes } from '../xlsx/read/book.js'
import { repairedName } from '../xlsx/read/sheet-names.js'
import { legacyConverter } from '../legacy/convert.js'
import { canonical, parse } from '../syntax.js'
import { decodeCsv, encodingName } from './decode.js'
import { csvRows, guessDelimiter, sepLine } from './fields.js'
import { CsvIssues } from './issues.js'
import type { CsvLimits, CsvRead, ReadCsvOptions } from './types.js'

const OPTIONS = ['locale', 'name', 'delimiter', 'encoding', 'formulas', 'codesAsText', 'now', 'limits'] as const

/** `opts.delimiter`: one character that is no quote and no line break, or `'locale'`. */
export function checkDelimiter(v: unknown, path: string): string {
  if (v === 'locale') return v
  if (typeof v !== 'string' || v.length !== 1 || /["\r\n]/.test(v)) throw new InputError(path, `expected one character other than a quote or a line break, or 'locale', got ${what(v)}`)
  return v
}

/** `opts.limits`: the cells, a whole number of at least 1. */
function checkLimits(v: unknown): CsvLimits {
  const o = object(v, 'opts.limits', 'limits such as { cells: 1000000 }')
  onlyKeys(o, 'opts.limits', ['cells'])
  if (o.cells !== undefined) checkNumber(o.cells, 'opts.limits.cells', 1, Number.MAX_SAFE_INTEGER, true)
  return o as CsvLimits
}

/** The sheet's name from the file's (`sales.csv` gives `Sales`, as Excel names it), and whether the file is one of tabs. */
function nameOf(file: string | undefined): { sheet: string; tabs: boolean } {
  if (file === undefined) return { sheet: 'Sheet1', tabs: false }
  const base = file.replace(/^.*[\\/]/, '')
  const ext = /\.(csv|tsv|txt)$/i.exec(base)
  return { sheet: repairedName(ext ? base.slice(0, ext.index) : base, 0), tabs: ext?.[1].toLowerCase() === 'tsv' }
}

/** Whether a field is digits alone that a number would change: a leading zero, or more than 15 digits. */
const isCode = (field: string) => /^\d+$/.test(field) && ((field.length > 1 && field[0] === '0') || field.length > 15)

/**
 * Reads a CSV file (its bytes, or its text) into a book of one sheet, as Excel opens the file in `opts.locale`
 * (en-US by default), with the delimiter and the encoding it was read in and what the book holds otherwise than Excel.
 */
export function readCsv(data: Uint8Array | ArrayBuffer | string, opts?: ReadCsvOptions): CsvRead {
  const o = opts === undefined ? {} : object(opts, 'opts', "options such as { locale: 'ru-RU', name: 'sales.csv' }")
  onlyOptions(o, 'opts', OPTIONS)
  const locale: Locale = o.locale === undefined ? localeOf(undefined) : checkLocaleArg(o.locale, 'opts.locale')
  const entry: EntryOptions = checkEntryOptions({ locale, ...(o.now !== undefined ? { now: o.now } : {}) })
  const file = nameOf(o.name === undefined ? undefined : checkString(o.name, 'opts.name'))
  const given = o.delimiter === undefined ? undefined : checkDelimiter(o.delimiter, 'opts.delimiter')
  const encoding = o.encoding === undefined ? undefined : encodingName(checkString(o.encoding, 'opts.encoding'), 'opts.encoding')
  for (const k of ['formulas', 'codesAsText'] as const) if (o[k] !== undefined) checkBoolean(o[k], `opts.${k}`)
  const limit = o.limits === undefined ? Infinity : (checkLimits(o.limits).cells ?? Infinity)
  const decoded = typeof data === 'string' ? { text: data.replace(/^\uFEFF/, ''), encoding: null, marked: false } : decodeCsv(fileBytes(data), displayRules(locale).codePage, encoding)
  // a first line `sep=X` names the delimiter, as Excel takes it, unless the host gives its own
  const sep = sepLine(decoded.text)
  const text = sep ? sep.rest : decoded.text
  // Excel reads a file marked as UTF-16 as Unicode text, at tabs
  const excel = () => (decoded.marked && decoded.encoding?.startsWith('utf-16') ? '\t' : locale.listSeparator)
  const delimiter = given !== undefined && given !== 'locale' ? given : (sep?.delimiter ?? (given === 'locale' ? excel() : file.tabs ? '\t' : (guessDelimiter(text) ?? locale.listSeparator)))
  const issues = new CsvIssues(file.sheet)
  const cells: Record<string, CellInput> = {}
  const rows = csvRows(text, delimiter)
  if (rows.length > MAX_ROWS || rows.some((r) => r.length > MAX_COLS)) issues.add('size', { key: 'csvOutside' })
  let count = 0
  let wraps = false
  for (let r = 0; r < Math.min(rows.length, MAX_ROWS); r++) {
    const row = rows[r]
    for (let c = 0; c < Math.min(row.length, MAX_COLS); c++) {
      let field = row[c]
      if (field === '') continue
      if (++count > limit) throw new Error(`csv: the file has more than ${limit} cells that are not empty (limits.cells)`)
      const at = formatAddress(r, c)
      if (field.length > MAX_TEXT) {
        field = field.slice(0, MAX_TEXT)
        issues.add('textLength', { key: 'csvTextCut' }, at)
      }
      const cell = cellOf(field, entry, o, issues, at)
      // a field with a line feed wraps its cell, as Excel opens it (a carriage return alone does not)
      cells[at] = field.includes('\n') ? { ...cell, style: 0 } : cell
      wraps ||= cell !== cells[at]
    }
  }
  const book: BookInput = { sheets: [{ name: file.sheet, cells }] }
  if (wraps) book.styles = [{ alignment: { wrap: true } }]
  if (locale.tag !== DEFAULT_LOCALE) book.locale = locale.tag
  return { book, issues: issues.list(locale), delimiter, encoding: decoded.encoding }
}

/**
 * A field as its cell's input: as pasted into the cell, and begun by `@` and a function's call as a formula (Excel takes
 * `@SUM(1)` so, as Lotus did; `@A1` is text). Excel enters the formula as one of the old kind, so `A1:A2` takes the
 * value in its row (`@A1:A2`, as `fromLegacy` reads it). A formula is kept as text unless `formulas`, a code with `codesAsText`.
 */
function cellOf(field: string, entry: EntryOptions, o: Record<string, unknown>, issues: CsvIssues, at: string): CellInput {
  if (o.codesAsText === true && isCode(field)) return { value: field }
  const formulas = o.formulas === true
  const lotus = /^@[\p{L}_][\p{L}\p{N}_.]*\(/u.test(field)
  const typed = lotus || (field.length > 1 && field[0] === '=')
  if (typed && !formulas) return text(field, issues, at)
  // pasted, a formula that does not parse is text; an error value (`#N/A`) is a formula of the error alone
  const r = readEntry(lotus ? `=${field.slice(1)}` : field, entry, 'pasted')
  const input = (r.ok && r.input) || { value: field }
  if (input.formula === undefined && typed) {
    issues.add('formula', { key: 'csvFormulaUnparsed' }, at)
    return { value: field }
  }
  if (input.formula === undefined || field[0] === '#') return input
  return formulas ? { ...input, formula: oldKind(input.formula) } : text(field, issues, at)
}

const toDynamic = legacyConverter()

/** A formula of the old kind as the dynamic formula Excel shows for it (`@A1:A2`), as written where it needs no `@`. */
function oldKind(formula: string): string {
  const ast = toDynamic(parse(formula))
  return ast ? canonical(ast) : formula
}

/** A field Excel takes as a formula, kept as text. */
function text(field: string, issues: CsvIssues, at: string): CellInput {
  issues.add('formula', { key: 'csvFormulaText' }, at)
  return { value: field }
}
