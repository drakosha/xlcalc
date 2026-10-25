/** Formulas computed for tests of functions: their values, and an error value by its code. */
import { expect } from 'vitest'
import { compute, type BookInput, type CellInput, type CellValue, type Clock, type ComputeResult } from '../src/index.js'
import { values } from './clock-books.js'

/** An error value by its code. */
export const err = (code: Extract<CellValue, { error: unknown }>['error']): CellValue => ({ error: code })

/**
 * Formulas computed into Z1, Z2, … of sheet S, beside `cells`, at `now`: their values in order. With `second`, the book
 * has a sheet T too, of `other`.
 */
export const valuesAt =
  (now: string, second = false) =>
  (formulas: string[], cells: Record<string, CellInput> = {}, other: Record<string, CellInput> = {}): CellValue[] => {
    const f: Record<string, CellInput> = {}
    formulas.forEach((x, i) => (f[`Z${i + 1}`] = { formula: x }))
    const r = compute({ sheets: [{ name: 'S', cells: { ...cells, ...f } }, ...(second ? [{ name: 'T', cells: other }] : [])] }, { clock: { now, seed: 1 } })
    return formulas.map((_, i) => r.sheets[0].cells[`Z${i + 1}`].value)
  }

/** The values of sheet S of `cells` computed at `now`, by address. */
export const sheetValuesAt =
  (now: string) =>
  (cells: Record<string, CellInput>): Record<string, CellValue> =>
    values(compute({ sheets: [{ name: 'S', cells }] }, { clock: { now, seed: 1 } }))

/** A book and the same book without some of its parts (`bare`): the same values and hash, another input hash. The book's result. */
export function sameValuesApart(book: BookInput, bare: BookInput, clock: Clock): ComputeResult {
  const [r, b] = [compute(book, { clock }), compute(bare, { clock })]
  expect(r.hash).toBe(b.hash)
  expect(r.sheets).toEqual(b.sheets)
  expect(r.inputHash).not.toBe(b.inputHash)
  return r
}
