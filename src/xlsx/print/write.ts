import type { BookInput } from '../../contract.js'
import { integer, object, onlyKeys } from '../../input/basic.js'
import { InputError } from '../../input-error.js'
import { MAX_COLS, MAX_ROWS, indexToCol } from '../../model.js'
import { parseRange } from '../../structure/area.js'
import { sheetKey } from '../../syntax/print.js'
import { areaText, sheetPrefix } from '../keep/refs.js'
import type { WorksheetPrintParts } from '../write/book.js'
import { escapeText } from '../xml/text.js'
import { MAX_PAGE_FIELD } from './types.js'
import { BREAK, HEADER_ATTRS, HEADER_TEXT, MARGINS, OPTIONS, PAGE_FIELD_WRAP, PROPERTIES, SETUP, checkedFields, leaf, writeFields } from './xml.js'

const SHEET_KEYS = ['sheet', 'areas', 'repeatRows', 'repeatColumns', 'setup', 'margins', 'options', 'headerFooter', 'rowBreaks', 'columnBreaks']

function span(input: unknown, path: string, limit: number): [number, number] {
  const value = object(input, path, 'a span with first and last lines')
  onlyKeys(value, path, ['first', 'last'])
  const first = integer(value.first, `${path}.first`, 1, limit)
  const last = integer(value.last, `${path}.last`, first, limit)
  return [first, last]
}

function namesFor(given: Record<string, unknown>, path: string, sheet: string, index: number): string[] {
  const prefix = sheetPrefix(sheet) + '!'
  const names: string[] = []
  const name = (kind: string, refs: string[]) => {
    if (refs.length) names.push(`<definedName name="_xlnm.${kind}" localSheetId="${index}">${escapeText(refs.map((ref) => prefix + ref).join(','))}</definedName>`)
  }
  if (given.areas !== undefined) {
    if (!Array.isArray(given.areas)) throw new InputError(`${path}.areas`, 'expected an array of ranges')
    const refs = given.areas.map((range: unknown, i: number) => {
      const rect = typeof range === 'string' ? parseRange(range.replace(/\$/g, '').toUpperCase()) : null
      if (!rect) throw new InputError(`${path}.areas[${i}]`, 'expected a sheet-local A1 range')
      return areaText(rect)
    })
    name('Print_Area', refs)
  }
  const titles: string[] = []
  if (given.repeatRows !== undefined) {
    const [first, last] = span(given.repeatRows, `${path}.repeatRows`, MAX_ROWS)
    titles.push(`$${first}:$${last}`)
  }
  if (given.repeatColumns !== undefined) {
    const [first, last] = span(given.repeatColumns, `${path}.repeatColumns`, MAX_COLS)
    titles.push(`$${indexToCol(first - 1)}:$${indexToCol(last - 1)}`)
  }
  name('Print_Titles', titles)
  return names
}

function headerFooter(value: unknown, path: string): string {
  if (value === undefined) return ''
  const given = checkedFields(value, path, HEADER_ATTRS, HEADER_TEXT)
  const attrs = writeFields(HEADER_ATTRS, given, path)
  const content = HEADER_TEXT.map((tag) => {
    const text = given[tag]
    if (text === undefined) return ''
    if (typeof text !== 'string') throw new InputError(`${path}.${tag}`, 'expected header or footer text')
    return `<${tag}>${escapeText(text)}</${tag}>`
  }).join('')
  return content ? `<headerFooter${attrs}>${content}</headerFooter>` : leaf('headerFooter', attrs)
}

function breaks(value: unknown, path: string, tag: string, limit: number, across: number): string {
  if (value === undefined) return ''
  if (!Array.isArray(value)) throw new InputError(path, 'expected an array of page breaks')
  let manualCount = 0
  const content = value.map((entry: unknown, i: number) => {
    const itemPath = `${path}[${i}]`
    const given = object(entry, itemPath, 'a page break')
    onlyKeys(given, itemPath, ['before', 'first', 'last', 'manual', 'pivot'])
    const before = integer(given.before, `${itemPath}.before`, 1, limit)
    const first = given.first === undefined ? undefined : integer(given.first, `${itemPath}.first`, 1, across)
    const last = given.last === undefined ? across : integer(given.last, `${itemPath}.last`, first ?? 1, across)
    const attrs = writeFields(BREAK, { id: before - 1, min: first === undefined ? undefined : first - 1, max: last - 1, man: given.manual, pt: given.pivot }, itemPath)
    if (given.manual === true) manualCount++
    return leaf('brk', attrs)
  }).join('')
  return `<${tag} count="${value.length}" manualBreakCount="${manualCount}">${content}</${tag}>`
}

function worksheetParts(given: Record<string, unknown>, path: string, sheet: string, index: number): WorksheetPrintParts {
  const setupPath = `${path}.setup`
  const optionsPath = `${path}.options`
  const marginsPath = `${path}.margins`
  const setup = given.setup === undefined ? {} : checkedFields(given.setup, setupPath, SETUP, Object.keys(PROPERTIES))
  const firstNumber = setup.firstPageNumber === undefined ? undefined
    : integer(setup.firstPageNumber, `${setupPath}.firstPageNumber`, -MAX_PAGE_FIELD, MAX_PAGE_FIELD)
  const firstPageNumber = firstNumber === undefined || firstNumber >= 0 ? firstNumber : PAGE_FIELD_WRAP + firstNumber
  const options = given.options === undefined ? {} : checkedFields(given.options, optionsPath, OPTIONS)
  const margins = given.margins === undefined ? {} : checkedFields(given.margins, marginsPath, MARGINS)
  if (given.margins !== undefined && Object.keys(MARGINS).some((key) => margins[key] === undefined)) {
    throw new InputError(marginsPath, 'expected all six page margins in inches')
  }
  return {
    properties: leaf('pageSetUpPr', writeFields(PROPERTIES, setup, setupPath)),
    settings: leaf('printOptions', writeFields(OPTIONS, options, optionsPath)) +
      leaf('pageMargins', writeFields(MARGINS, margins, marginsPath)) +
      leaf('pageSetup', writeFields(SETUP, { ...setup, firstPageNumber }, setupPath)) +
      headerFooter(given.headerFooter, `${path}.headerFooter`) +
      breaks(given.rowBreaks, `${path}.rowBreaks`, 'rowBreaks', MAX_ROWS, MAX_COLS) +
      breaks(given.columnBreaks, `${path}.columnBreaks`, 'colBreaks', MAX_COLS, MAX_ROWS),
    names: namesFor(given, path, sheet, index),
  }
}

export function printParts(book: BookInput, print: unknown): ReadonlyMap<string, WorksheetPrintParts> {
  if (!Array.isArray(print)) throw new InputError('print', 'expected settings by sheet, as an array')
  const sheets = new Map(book.sheets.map((sheet, index) => [sheetKey(sheet.name), { name: sheet.name, index }]))
  const parts = new Map<string, WorksheetPrintParts>()
  print.forEach((entry: unknown, i) => {
    const path = `print[${i}]`
    const given = object(entry, path, 'print settings for a sheet')
    onlyKeys(given, path, SHEET_KEYS)
    const sheet = typeof given.sheet === 'string' ? sheets.get(sheetKey(given.sheet)) : undefined
    if (!sheet) throw new InputError(`${path}.sheet`, `unknown worksheet ${JSON.stringify(given.sheet)}`)
    if (parts.has(sheet.name)) throw new InputError(`${path}.sheet`, `repeated worksheet ${JSON.stringify(sheet.name)}`)
    parts.set(sheet.name, worksheetParts(given, path, sheet.name, sheet.index))
  })
  return parts
}
