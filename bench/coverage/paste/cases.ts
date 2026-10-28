/** The cases of the step `paste`: a book, a range copied and pasted with Paste Special at a target. */
import type { BookInput, RangeRef } from '../../../src/contract.js'
import { colToIndex, formatAddress } from '../../../src/model.js'

/**
 * What Paste Special pastes, as the harness names Excel's `XlPasteType` (`notes` is `xlPasteComments`); `copy`, a plain
 * copy (`Range.Copy` with a destination); `painter`, the Format Painter of the Home tab (Excel for Mac alone).
 */
export type PasteWhat = 'copy' | 'all' | 'values' | 'formulas' | 'formats' | 'validation' | 'formulasAndNumberFormats' | 'valuesAndNumberFormats' | 'columnWidths' | 'allExceptBorders' | 'notes' | 'painter'

/** Paste Special's operation (`XlPasteSpecialOperation`), none when left out. */
export type PasteOperation = 'add' | 'subtract' | 'multiply' | 'divide'

export interface PasteCase {
  id: string
  /** The key of its book in `PasteCases.books`. */
  book: string
  from: RangeRef
  /** The target as selected before the paste: a cell, or a range the copies tile. */
  to: RangeRef
  what: PasteWhat
  transpose?: boolean
  operation?: PasteOperation
  skipBlanks?: boolean
}

export interface PasteCases {
  books: Record<string, BookInput>
  cases: PasteCase[]
  /** By book, the cells (on any sheet) whose formulas go to Excel as it saves a formula typed into a cell, not as dynamic array formulas. */
  plain?: Record<string, string[]>
}

/** What Excel answered: the error of a paste it refused (its saved file is the book as it was). */
export interface PasteAnswer {
  error?: string
}

export interface PasteSet {
  buildPasteCases(): PasteCases
}

/**
 * What the Format Painter selects before a key takes the selection onto the target, as the painter applies at a change
 * of the selection by the keyboard, not by the object model: the target and a column to its right, shrunk by Shift+Left;
 * whole rows and a row below them, shrunk by Shift+Up.
 */
export function painterSelection(range: string): { select: string; key: 'left' | 'up' } {
  const rows = /^(\d+):(\d+)$/.exec(range)
  if (rows) return { select: `${rows[1]}:${Number(rows[2]) + 1}`, key: 'up' }
  const end = /([A-Z]+)(\d*)$/.exec(range)!
  const next = formatAddress(0, colToIndex(end[1]) + 1).replace(/\d+$/, '')
  return { select: `${range.includes(':') ? range.slice(0, range.indexOf(':')) : range}:${next}${end[2]}`, key: 'left' }
}

/** The paste as the script takes it; the Format Painter's target as it selects it and the key. */
export function pasteArgs(c: PasteCase): string {
  const p = c.what === 'painter' ? painterSelection(c.to.range) : null
  return JSON.stringify([c.from.sheet, c.from.range, c.to.sheet, p ? p.select : c.to.range, c.what, c.transpose ? '1' : '0', c.operation ?? 'none', c.skipBlanks ? '1' : '0', p ? p.key : ''])
}
