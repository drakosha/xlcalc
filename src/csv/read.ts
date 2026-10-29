/**
 * `readCsv`: a CSV file as a book of one sheet, as Excel opens it. The text (`decode.ts`) is split into rows and fields
 * (`fields.ts`), and each field is read as text pasted into a cell in the book's locale (`readEntry`, the way of a
 * paste): a number, a date, a time, a percentage or money with the format Excel sets, a boolean, an error, text, or a
 * formula; an apostrophe before text stays part of it, and a formula that does not parse is text. Where the book holds
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
  const decoded = typeof data === 'string' ? { text: data.replace(/^﻿/, ''), encoding: null } : decodeCsv(fileBytes(data), displayRules(locale).codePage, encoding)
  // a first line `sep=X` names the delimiter, as Excel takes it, unless the host gives its own
  const sep = sepLine(decoded.text)
  const text = sep ? sep.rest : decoded.text
  const delimiter =
    given !== undefined && given !== 'locale' ? given : (sep?.delimiter ?? (given === 'locale' ? locale.listSeparator : file.tabs ? '\t' : (guessDelimiter(text) ?? locale.listSeparator)))
  const issues = new CsvIssues(file.sheet)
  const cells: Record<string, CellInput> = {}
  const rows = csvRows(text, delimiter)
  if (rows.length > MAX_ROWS || rows.some((r) => r.length > MAX_COLS)) issues.add('size', { key: 'csvOutside' })
  let count = 0
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
      cells[at] = cellOf(field, entry, o, issues, at)
    }
  }
  const book: BookInput = { sheets: [{ name: file.sheet, cells }] }
  if (locale.tag !== DEFAULT_LOCALE) book.locale = locale.tag
  return { book, issues: issues.list(locale), delimiter, encoding: decoded.encoding }
}

/** A field as its cell's input: as pasted into the cell, a formula kept as text unless `formulas`, a code kept as text with `codesAsText`. */
function cellOf(field: string, entry: EntryOptions, o: Record<string, unknown>, issues: CsvIssues, at: string): CellInput {
  if (o.codesAsText === true && isCode(field)) return { value: field }
  const formulas = o.formulas === true
  const typed = field.length > 1 && field[0] === '='
  if (typed && !formulas) return text(field, issues, at)
  // pasted, a formula that does not parse is text; an error value (`#N/A`) is a formula of the error alone
  const r = readEntry(field, entry, 'pasted')
  const input = (r.ok && r.input) || { value: field }
  if (input.formula === undefined && typed) issues.add('formula', { key: 'csvFormulaUnparsed' }, at)
  return input.formula !== undefined && field[0] !== '#' && !formulas ? text(field, issues, at) : input
}

/** A field Excel takes as a formula, kept as text. */
function text(field: string, issues: CsvIssues, at: string): CellInput {
  issues.add('formula', { key: 'csvFormulaText' }, at)
  return { value: field }
}
