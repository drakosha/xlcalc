import type { BookInput } from '../contract.js'
import { MAX_ROW_HEIGHT } from '../column-width.js'
import { InputError } from '../input-error.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { MAX_PAGE_FIELD } from '../xlsx/print/types.js'
import { parseRange } from '../structure/area.js'
import {
  HEADER_ATTRS, HEADER_TEXT, MARGINS, OPTIONS, PAGE_FIELD_WRAP, PROPERTIES, SETUP, checkedFields, checkPrintFields,
} from '../xlsx/print/xml.js'
import { array, checkNumber, checkString, integer, object, onlyKeys } from './basic.js'

const SHEET_KEYS = ['areas', 'repeatRows', 'repeatColumns', 'setup', 'margins', 'options', 'headerFooter', 'rowBreaks', 'columnBreaks']

function span(input: unknown, path: string, limit: number, start = 1): number {
  const value = object(input, path, 'a span with first and last lines')
  onlyKeys(value, path, ['first', 'last'])
  const first = integer(value.first, `${path}.first`, start, limit)
  return integer(value.last, `${path}.last`, first, limit)
}

function checkBreaks(input: unknown, path: string, limit: number, across: number): void {
  for (const [i, value] of array(input, path, 'an array of page breaks').entries()) {
    const at = `${path}[${i}]`
    const br = object(value, at, 'a page break')
    onlyKeys(br, at, ['before', 'first', 'last', 'manual', 'pivot'])
    integer(br.before, `${at}.before`, 1, limit)
    const first = br.first === undefined ? 1 : integer(br.first, `${at}.first`, 1, across)
    if (br.last !== undefined) integer(br.last, `${at}.last`, first, across)
    for (const key of ['manual', 'pivot']) {
      if (br[key] !== undefined && typeof br[key] !== 'boolean') throw new InputError(`${at}.${key}`, 'expected a boolean')
    }
  }
}

export function checkWorksheetPrint(input: unknown, path: string): void {
  const settings = object(input, path, 'worksheet print settings')
  onlyKeys(settings, path, SHEET_KEYS)
  if (settings.areas !== undefined) {
    for (const [i, entry] of array(settings.areas, `${path}.areas`, 'an array of print areas').entries()) {
      const at = `${path}.areas[${i}]`
      if (!parseRange(checkString(entry, at).replace(/\$/g, '').toUpperCase())) throw new InputError(at, 'expected a sheet-local A1 range')
    }
  }
  if (settings.repeatRows !== undefined) span(settings.repeatRows, `${path}.repeatRows`, MAX_ROWS)
  if (settings.repeatColumns !== undefined) span(settings.repeatColumns, `${path}.repeatColumns`, MAX_COLS)
  if (settings.setup !== undefined) {
    const at = `${path}.setup`
    const fields = checkedFields(settings.setup, at, SETUP, Object.keys(PROPERTIES))
    checkPrintFields(PROPERTIES, fields, at)
    const first = fields.firstPageNumber === undefined ? undefined : integer(fields.firstPageNumber, `${at}.firstPageNumber`, -MAX_PAGE_FIELD, MAX_PAGE_FIELD)
    checkPrintFields(SETUP, typeof first === 'number' && first < 0 ? { ...fields, firstPageNumber: PAGE_FIELD_WRAP + first } : fields, at)
  }
  if (settings.options !== undefined) checkPrintFields(OPTIONS, checkedFields(settings.options, `${path}.options`, OPTIONS), `${path}.options`)
  if (settings.margins !== undefined) {
    const at = `${path}.margins`
    const fields = checkedFields(settings.margins, at, MARGINS)
    if (Object.keys(MARGINS).some((key) => fields[key] === undefined)) throw new InputError(at, 'expected all six page margins in inches')
    checkPrintFields(MARGINS, fields, at)
  }
  if (settings.headerFooter !== undefined) {
    const at = `${path}.headerFooter`
    const fields = checkedFields(settings.headerFooter, at, HEADER_ATTRS, HEADER_TEXT)
    checkPrintFields(HEADER_ATTRS, fields, at)
    for (const key of HEADER_TEXT) if (fields[key] !== undefined) checkString(fields[key], `${at}.${key}`)
  }
  if (settings.rowBreaks !== undefined) checkBreaks(settings.rowBreaks, `${path}.rowBreaks`, MAX_ROWS, MAX_COLS)
  if (settings.columnBreaks !== undefined) checkBreaks(settings.columnBreaks, `${path}.columnBreaks`, MAX_COLS, MAX_ROWS)
}

export function checkRowDefaults(input: unknown, path: string): void {
  const defaults = object(input, path, 'worksheet row defaults')
  onlyKeys(defaults, path, ['defaultRowHeight', 'visibleRows'])
  if (defaults.defaultRowHeight !== undefined) checkNumber(defaults.defaultRowHeight, `${path}.defaultRowHeight`, 0, MAX_ROW_HEIGHT)
  if (defaults.visibleRows !== undefined) {
    let start = 1
    for (const [i, value] of array(defaults.visibleRows, `${path}.visibleRows`, 'an array of row spans').entries()) {
      start = span(value, `${path}.visibleRows[${i}]`, MAX_ROWS, start) + 1
    }
  }
}

/** Checks the worksheet metadata in entry points that edit or write the workbook. */
export function checkBookPrinting(book: BookInput, path = 'book'): void {
  for (const [i, sheet] of book.sheets.entries()) {
    const at = `${path}.sheets[${i}]`
    if (sheet.print !== undefined) checkWorksheetPrint(sheet.print, `${at}.print`)
    if (sheet.rowDefaults !== undefined) checkRowDefaults(sheet.rowDefaults, `${at}.rowDefaults`)
  }
}
