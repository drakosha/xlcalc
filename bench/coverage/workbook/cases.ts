/**
 * Workbook cases for Excel (`run.ts workbook <set>`): a small book each, with what its file holds beyond the
 * model (hidden rows and columns, filters, the 1904 date system, iterative calculation, cached values), the
 * steps Excel takes after opening it and the rectangles read at each `read` step.
 */
import type { FileParts } from '../case-book.js'
import type { Val } from '../formulajs.js'
import type { BookInput, RangeRef } from '../../../src/contract.js'

/** One step after Excel opens the book (it computes the book on opening, as the file has no values). */
export type Step =
  /** Excel's full calculation (`calculate full`), a recalculation (`calculate`), a read of the rectangles. */
  | 'full'
  | 'calc'
  | 'read'
  /** A number or a text typed into a cell, or null to clear it (a session's `setCells`). */
  | { set: { sheet: string; cell: string; value: number | string | null } }
  /** A formula typed into a cell, in the engine's spelling (a session's `setCells`). */
  | { formula: { sheet: string; cell: string; formula: string } }
  /** Rows (`5:6`) or columns (`C:D`) hidden or shown (a session's `setLayout`). */
  | { hide: { sheet: string; rows?: string; columns?: string; hidden: boolean } }
  /** Excel's autofilter on a range: a column (1 from its left) keeps the rows equal to `criterion`. */
  | { filter: { sheet: string; range: string; field: number; criterion: string } }

export interface WorkbookCase {
  id: string
  /** What the case checks. */
  about: string
  book: BookInput
  file?: FileParts
  read: RangeRef[]
  /** Rectangles whose cells' shown text (`Range.Text`) each read takes. */
  show?: RangeRef[]
  /** `['full', 'read']` when left out. */
  steps?: Step[]
  /**
   * A held-out case: the answer of each read predicted before Excel ran it, the values of non-blank cells by
   * `Sheet!A1` and the text of the cells to show by `show:Sheet!A1`.
   */
  predicted?: Record<string, Val | string>[]
}

export interface WorkbookSet {
  buildWorkbookCases(): WorkbookCase[]
}

/** Excel's answer: the application's iteration settings after opening the book, and each read (the shown text by `Sheet!A1` when the case has `show`). */
export interface WorkbookAnswer {
  settings: { iteration: boolean; count: number; delta: number }
  reads: { values: Record<string, Val>; formulas: Record<string, string>; shown?: Record<string, string> }[]
}
