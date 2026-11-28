/** SpreadsheetML attributes shared by the print reader and writer. */
import { InputError } from '../../input-error.js'
import { object, onlyKeys } from '../../input/basic.js'
import { escapeAttr } from '../xml/text.js'
import { MAX_PAGE_FIELD } from './types.js'

type Scalar = boolean | number | string
type Parse = (text: string) => Scalar | undefined
type Fields = Readonly<Record<string, Parse>>

const bool: Parse = (text) => {
  if (text === '1' || text === 'true') return true
  if (text === '0' || text === 'false') return false
}

const unsigned: Parse = (text) => {
  const value = Number(text)
  if (/^\+?\d+$/.test(text) && value <= 0xffffffff) return value
}

export const PAGE_FIELD_WRAP = 0x100000000
const pageCount: Parse = (text) => {
  const value = unsigned(text)
  if (typeof value === 'number' && value <= MAX_PAGE_FIELD) return value
}

const pageNumber: Parse = (text) => {
  const value = unsigned(text)
  if (typeof value === 'number' && (value <= MAX_PAGE_FIELD || value >= PAGE_FIELD_WRAP - MAX_PAGE_FIELD)) return value
}

const copies: Parse = (text) => pageCount(text) || undefined
const scale: Parse = (text) => {
  const value = unsigned(text)
  if (typeof value === 'number' && (value === 0 || value >= 10 && value <= 400)) return value
}

const margin: Parse = (text) => {
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(text)) return undefined
  const value = Number(text)
  if (Number.isFinite(value) && value >= 0) return value
}

const measure: Parse = (text) => /^\d+(?:\.\d+)?(?:mm|cm|in|pt|pc|pi)$/.test(text) ? text : undefined
const oneOf = (...values: string[]): Parse => (text) => values.includes(text) ? text : undefined

export const PROPERTIES = { autoPageBreaks: bool, fitToPage: bool }
export const SETUP = {
  paperSize: unsigned, paperHeight: measure, paperWidth: measure, scale,
  fitToWidth: pageCount, fitToHeight: pageCount, pageOrder: oneOf('downThenOver', 'overThenDown'),
  orientation: oneOf('default', 'portrait', 'landscape'), usePrinterDefaults: bool, blackAndWhite: bool, draft: bool,
  cellComments: oneOf('none', 'asDisplayed', 'atEnd'), errors: oneOf('displayed', 'blank', 'dash', 'NA'),
  firstPageNumber: pageNumber, useFirstPageNumber: bool, horizontalDpi: unsigned, verticalDpi: unsigned, copies,
}
export const MARGINS = { left: margin, right: margin, top: margin, bottom: margin, header: margin, footer: margin }
export const OPTIONS = { horizontalCentered: bool, verticalCentered: bool, headings: bool, gridLines: bool, gridLinesSet: bool }
export const HEADER_ATTRS = { differentOddEven: bool, differentFirst: bool, scaleWithDoc: bool, alignWithMargins: bool }
export const HEADER_TEXT = ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'] as const

export const BREAK = { id: unsigned, min: unsigned, max: unsigned, man: bool, pt: bool }

/** Read known attributes, reporting malformed values to the caller. */
export function readFields<T>(fields: Fields, attrs: Record<string, string>, invalid: () => void): T {
  const values: Record<string, Scalar> = {}
  for (const [key, parse] of Object.entries(fields)) {
    const raw = attrs[key]
    if (raw === undefined) continue
    const value = parse(raw.trim())
    if (value === undefined) invalid()
    else values[key] = value
  }
  return values as T
}

export function checkedFields(value: unknown, path: string, fields: Fields, others: readonly string[] = []): Record<string, unknown> {
  const given = object(value, path, 'print settings')
  onlyKeys(given, path, [...Object.keys(fields), ...others])
  return given
}

export function fieldText(value: unknown): string {
  return typeof value === 'boolean' ? (value ? '1' : '0') : String(value)
}

export function checkPrintFields(fields: Fields, given: Record<string, unknown>, path: string): void {
  for (const [key, parse] of Object.entries(fields)) {
    const value = given[key]
    if (value !== undefined && parse(fieldText(value)) !== value) throw new InputError(`${path}.${key}`, 'invalid print setting')
  }
}

export function writeFields(fields: Fields, given: Record<string, unknown>): string {
  return Object.keys(fields).flatMap((key) => given[key] === undefined ? [] : [` ${key}="${escapeAttr(fieldText(given[key]))}"`]).join('')
}

export const leaf = (name: string, attrs: string): string => attrs ? `<${name}${attrs}/>` : ''
