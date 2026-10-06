/**
 * Spill cases for Excel (`run.ts spill <set>`): a small book each, its formulas written as dynamic array
 * formulas, and the rectangles whose values are read after Excel computes it. No case has a circular
 * reference (Excel would show a warning window).
 */
import type { Val } from '../formulajs.js'
import type { BookInput, CellInput, RangeRef, SheetInput } from '../../../src/contract.js'

export interface SpillCase {
  id: string
  /** What the case checks. */
  about: string
  /** Formulas in the engine's spelling (`E2#`, no `_xlfn.`), written as Excel 365 stores them. */
  book: BookInput
  /** Rectangles read after the computation. */
  read: RangeRef[]
  /** A volatile case: its numbers are compared by type only. */
  volatile?: true
  /** A held-out case: its answer (non-blank cells of `read`, `Sheet!A1`) predicted before Excel computed it. */
  predicted?: Record<string, Val>
}

export interface SpillSet {
  buildSpillCases(): SpillCase[]
}

/** A cell: `=…` a formula, `{ format }` or `{ style }` an empty cell with a look, else a value. */
export type CellSpec = number | string | boolean | CellInput

const cell = (v: CellSpec): CellInput => (typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1) } : typeof v === 'object' ? v : { value: v })

/** A sheet from `{ A1: spec }`, with merged areas. */
export const sheet = (name: string, cells: Record<string, CellSpec>, merges?: string[]): SheetInput => ({
  name,
  cells: Object.fromEntries(Object.entries(cells).map(([a, v]) => [a, cell(v)])),
  ...(merges ? { merges } : {}),
})

/** A case on one sheet `S`, read in `ranges` of S. */
export const one = (id: string, about: string, cells: Record<string, CellSpec>, ranges: string[], extra: Partial<SpillCase> & { merges?: string[] } = {}): SpillCase => {
  const { merges, book, ...rest } = extra
  return { id, about, book: { ...book, sheets: [sheet('S', cells, merges), ...(book?.sheets ?? [])] }, read: ranges.map((range) => ({ sheet: 'S', range })), ...rest }
}

/** Numbers as answers, for predictions. */
export const n = (v: number): Val => ({ t: 'n', v })
export const e = (v: Extract<Val, { t: 'e' }>['v']): Val => ({ t: 'e', v })
export const s = (v: string): Val => ({ t: 's', v })
export const b = (v: boolean): Val => ({ t: 'b', v })
