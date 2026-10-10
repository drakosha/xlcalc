/**
 * Books and edits that the asynchronous modes are compared on with the synchronous one: the coverage measurement (in
 * en-US and ru-RU), the semantics cases, the synthetic book, the spill book, seeded books, iterative calculation,
 * the 1904 date system; and a book of a long chain, where one edit recomputes thousands of formulas.
 */
import type { BookInput, CellReplace, CellInput } from '../src/contract.js'
import { synthBook } from '../src/synth.js'
import { CASES, excelBook } from './excel-cases.js'
import { randomBook } from './seeded-books.js'
import { CASES as SEMANTICS, semanticsBook } from './semantics-cases.js'
import { spillBook } from './session-spill-edits.js'

export const CLOCK = { now: '2026-10-04T12:00:00Z', seed: 7 }

export const BOOKS: [string, BookInput][] = [
  ['the coverage measurement', excelBook(CASES)],
  ['the coverage measurement in ru-RU', { ...excelBook(CASES), locale: 'ru-RU' }],
  ['the semantics cases', semanticsBook(SEMANTICS)],
  ['the synthetic book', synthBook({ seed: 1, rows: 340 }).book],
  ['the spill book', spillBook()],
  ['a seeded book', randomBook(11)],
  ['iterative calculation', { sheets: [{ name: 'S', cells: { B1: { value: 1000 }, B2: { value: 0.1 }, B3: { formula: 'B2*(B1+B4)/2' }, B4: { formula: 'B1+B3-100' } } }], iteration: {} }],
  ['the 1904 date system', { dateSystem: 1904, sheets: [{ name: 'S', cells: { A1: { formula: 'DATE(2025,1,31)' }, A2: { formula: 'TODAY()-A1' }, A3: { formula: 'RAND()' } } }] }],
]

/** `n` formulas over A1, each reading the one before: an edit of A1 recomputes all of them. */
export function chainBook(n: number): BookInput {
  const cells: Record<string, CellInput> = { A1: { value: 1 }, B1: { formula: 'A1+1' } }
  for (let r = 2; r <= n; r++) cells[`B${r}`] = { formula: `B${r - 1}+A$1` }
  return { sheets: [{ name: 'S', cells }] }
}

export const setA1 = (value: number): CellReplace[] => [{ sheet: 'S', cell: 'A1', input: { value } }]

/** A book whose formula in B1 is heavy (a matrix product of `n` × `n`) once A1 is above 1. */
export const slowBook = (n = 700): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: `IF(A1>1,SUM(MMULT(SEQUENCE(${n},${n}),SEQUENCE(${n},${n}))),A1)` }, B2: { formula: 'B1+1' } } }] })

/** A book where every edit of A1 recomputes a matrix product of `n` × `n` (300: about 150 ms). */
export const heavyBook = (n = 300): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: `SUM(MMULT(SEQUENCE(${n},${n}),SEQUENCE(${n},${n})))+A1` } } }] })

/** The median time of `n` calls of `f`: how fast the machine runs now, for deadlines a test sets at any speed. */
export async function medianMs(n: number, f: (i: number) => unknown): Promise<number> {
  const xs: number[] = []
  for (let i = 0; i < n; i++) {
    const t = performance.now()
    await f(i)
    xs.push(performance.now() - t)
  }
  return xs.sort((a, b) => a - b)[n >> 1]
}
