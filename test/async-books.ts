/**
 * Books and edits that the asynchronous modes are compared on with the synchronous one: the coverage measurement (in
 * en-US and ru-RU), the semantics cases, the synthetic book, the spill book, seeded books, iterative calculation,
 * the 1904 date system; and a book of a long chain, where one edit recomputes thousands of formulas.
 */
import { expect } from 'vitest'
import type { BookInput, CellReplace, CellInput } from '../src/contract.js'
import { CalcError, createSession, type AsyncSession, type Execution, type Session } from '../src/index.js'
import type { SynthBook } from '../src/synth.js'
import { median } from '../bench/stats.mjs'
import { edits } from './session-edits.js'
import { synthBook } from '../src/synth.js'
import { CASES, excelBook } from './excel-cases.js'
import { randomBook } from './seeded-books.js'
import { CASES as SEMANTICS, semanticsBook } from './semantics-cases.js'
import { spillBook } from './session-spill-edits.js'
import './locales.js'

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
  return median(xs)
}

/** The CalcError a call fails with; a call that does not fail fails the test. */
export const fails = async (p: Promise<unknown>): Promise<CalcError> => {
  try {
    await p
  } catch (e) {
    expect(e).toBeInstanceOf(CalcError)
    return e as CalcError
  }
  throw new Error('expected CalcError')
}

/** The longest wait of the event loop while `p` is pending, and its runs; a call that never yields leaves it none. */
export async function loopDuring<T>(p: Promise<T>): Promise<{ value: T; worst: number; ticks: number }> {
  let last = performance.now()
  let worst = 0
  let ticks = 0
  let done = false
  const tick = () => {
    const now = performance.now()
    worst = Math.max(worst, now - last)
    last = now
    ticks++
    if (!done) setImmediate(tick)
  }
  setImmediate(tick)
  const value = await p.finally(() => (done = true))
  return { value, worst: Math.max(worst, performance.now() - last), ticks }
}

/** A session `s` and a synchronous one over the synthetic book `synth`, given 40 edits: the same changes and hash after each; the synchronous one. */
export async function sameOverEdits(s: AsyncSession, synth: SynthBook, execution: Execution): Promise<Session> {
  const sync = createSession(synth.book, { clock: CLOCK, execution })
  for (const batch of edits(5, 40, synth)) {
    expect(await s.setCells(batch)).toEqual(sync.setCells(batch))
    expect((await s.getResult()).hash).toBe(sync.getResult().hash)
  }
  return sync
}

/**
 * A call stopped by its soft deadline on the heavy book's session `s`, and the call after it, which computes the book
 * again outside its deadline: the deadline, a quarter of a cold computation of the book as fast as the machine runs now.
 */
export async function afterStopped(s: AsyncSession): Promise<number> {
  const timeoutMs = (await medianMs(3, (i) => s.setCells(setA1(10 + i)))) / 4
  expect((await fails(s.setCells(setA1(2), { timeoutMs }))).code).toBe('soft_timeout')
  expect((await s.setCells([{ sheet: 'S', cell: 'C1', input: { value: 5 } }], { timeoutMs })).changes).toEqual([{ sheet: 'S', cell: 'C1', value: 5, text: '5' }])
  expect((await s.getBook()).sheets[0].cells.A1).toEqual({ value: 12 })
  return timeoutMs
}
