/**
 * Selections in Excel for Windows (`bench/windows/run.ts select <set>`): a set names its books, each the sheet `S` with
 * its cells, and its cases, each a range of one book and the sides of it whose labels name the rest, as Excel's Create
 * Names from Selection (`Range.CreateNames`); the answer is the names Excel made, each with what it refers to, or
 * Excel's error. `predict.ts` writes the engine's answers (`Session.getNamesFromSelection`) before Excel runs a set.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { BookInput, CellInput, CellStyle, SelectionLabels } from '../../../src/contract.js'
import { createSession, validateFormula } from '../../../src/index.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../../..')

/** A book of a set: the sheet `S` and the styles its cells take. */
export interface SelectBook {
  cells: Record<string, CellInput>
  styles?: CellStyle[]
}

export interface SelectCase {
  id: string
  book: string
  range: string
  labels: SelectionLabels
}

/** The names made, each as `[name, what it refers to]` without `=`, in their canonical spelling, sorted by name; or Excel's error. */
export type SelectAnswer = { names: [string, string][] } | { error: string }

export interface SelectSet {
  BOOKS: Record<string, SelectBook>
  buildSelectCases(): SelectCase[]
}

/** The stored answers of a set. */
export const selectFile = (set: string) => join(ROOT, 'test/excel/win/select', `${set}.json`)

/** The book as Excel and the engine get it. */
export const selectBook = (b: SelectBook): BookInput => ({ ...(b.styles ? { styles: b.styles } : {}), sheets: [{ name: 'S', cells: b.cells }] })

/** A reference as it compares: canonical where it parses. */
function canon(ref: string): string {
  const r = validateFormula(ref.replace(/^=/, ''))
  return r.ok ? r.canonical : ref
}

/** Names as they compare: canonical references, sorted by name. */
export const namesAnswer = (names: [string, string][]): SelectAnswer => ({ names: names.map(([n, ref]): [string, string] => [n, canon(ref)]).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)) })

/** The engine's answers to the cases of one book. */
export function selectedInBook(b: SelectBook, cases: SelectCase[]): Record<string, SelectAnswer> {
  const s = createSession(selectBook(b), { clock: { now: '2026-10-28T12:00:00Z', seed: 1 } })
  try {
    return Object.fromEntries(cases.map((c) => [c.id, namesAnswer(s.getNamesFromSelection('S', c.range, c.labels).map((n) => [n.name, n.ref]))]))
  } finally {
    s.close()
  }
}
