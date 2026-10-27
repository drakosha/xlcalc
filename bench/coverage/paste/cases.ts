/** The cases of the step `paste`: a book, a range copied and pasted with Paste Special at a target. */
import type { BookInput, RangeRef } from '../../../src/contract.js'

/** What Paste Special pastes, as the harness names Excel's `XlPasteType`; `copy`, a plain copy (`Range.Copy` with a destination). */
export type PasteWhat = 'copy' | 'all' | 'values' | 'formulas' | 'formats' | 'validation' | 'formulasAndNumberFormats' | 'valuesAndNumberFormats' | 'columnWidths' | 'allExceptBorders'

export interface PasteCase {
  id: string
  /** The key of its book in `PasteCases.books`. */
  book: string
  from: RangeRef
  /** The target as selected before the paste: a cell, or a range the copies tile. */
  to: RangeRef
  what: PasteWhat
  transpose?: boolean
}

export interface PasteCases {
  books: Record<string, BookInput>
  cases: PasteCase[]
}

/** What Excel answered: the error of a paste it refused (its saved file is the book as it was). */
export interface PasteAnswer {
  error?: string
}

export interface PasteSet {
  buildPasteCases(): PasteCases
}

/** The paste as the script takes it. */
export const pasteArgs = (c: PasteCase): string => JSON.stringify([c.from.sheet, c.from.range, c.to.sheet, c.to.range, c.what, c.transpose ? '1' : '0'])
