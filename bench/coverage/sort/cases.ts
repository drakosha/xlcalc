/**
 * Cases of the Sort command (`run.ts sort <set>`, `bench/windows/run.ts sort <set>`): a column of cells sorted by
 * itself, without a header, ascending or descending. A cell is written as Excel shows its formula (`=NA()`, `10`, `b`,
 * `TRUE`, '' for a blank), and each cell of a case is told from the others by it, so the order Excel leaves is read by
 * the formulas.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import { indexToCol } from '../../../src/model.js'

export interface SortCase {
  id: string
  cells: string[]
  descending?: true
}

export interface SortSet {
  cases: SortCase[]
}

/** The answers of a set: each case's column, as Excel shows its formulas after the sort. */
export interface SortAnswers {
  meta: Record<string, unknown>
  answers: Record<string, string[]>
}

/** A cell from the text Excel shows as its formula; a blank is none. */
function cellOf(t: string): CellInput | undefined {
  if (t === '') return undefined
  if (t.startsWith('=')) return { formula: t.slice(1) }
  if (t === 'TRUE' || t === 'FALSE') return { value: t === 'TRUE' }
  return { value: /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : t }
}

/** The book of a set: case i in column i of sheet S, from row 1. */
export function sortBook(set: SortSet): BookInput {
  const cells: Record<string, CellInput> = {}
  set.cases.forEach((c, i) =>
    c.cells.forEach((t, r) => {
      const x = cellOf(t)
      if (x) cells[`${indexToCol(i)}${r + 1}`] = x
    }),
  )
  return { sheets: [{ name: 'S', cells }] }
}

/** The range case i sorts. */
export const rangeOf = (set: SortSet, i: number): string => `${indexToCol(i)}1:${indexToCol(i)}${set.cases[i].cells.length}`
