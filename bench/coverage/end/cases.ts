/**
 * Excel's `Range.End` (Control and an arrow key), for Excel for Windows (`bench/windows/run.ts end <set> [ru|<tag>]`):
 * a set names its books, each the sheet `S` with its cells, looks, columns, hidden rows and columns, merged areas, an
 * autofilter and texts typed once Excel opened it, and its cases, each a cell of one book and a direction: the cell
 * `End` comes to.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { BookInput, CellInput, CellStyle, ColumnInput, EndDirection, LookRange, RowInput } from '../../../src/contract.js'
import { createSession } from '../../../src/index.js'
import { filteredRows } from '../find/cases.js'
import { localeDir } from '../locales.js'
import { groupCollector } from '../collect.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../../..')

/** A book of a set: the sheet `S`. */
export interface EndBook {
  cells: Record<string, CellInput>
  styles?: CellStyle[]
  looks?: LookRange[]
  columns?: Record<string, ColumnInput>
  rows?: Record<string, RowInput>
  merges?: string[]
  /** An autofilter over `range` keeping the rows whose `field`-th column holds `keep`: Excel hides the others as it applies it, the engine's book holds them hidden on a filtered sheet. */
  filter?: { range: string; field: number; keep: string }
  /** Texts typed into cells once Excel opened the book (`Formula2Local`), for what a file does not hold (an apostrophe alone); the engine's book holds the cell of `cells`. */
  typed?: Record<string, string>
}

export interface EndCase {
  id: string
  book: string
  cell: string
  direction: EndDirection
}

/** The cell `End` came to, or the error Excel answered. */
export type EndAnswer = { cell: string } | { error: string }

export interface EndSet {
  BOOKS: Record<string, EndBook>
  buildEndCases(): EndCase[]
}

export interface StoredEnd {
  meta: Record<string, string | number>
  books: Record<string, EndBook>
  cases: EndCase[]
  answers: Record<string, EndAnswer>
}

/** The stored answers of a set: `test/excel/win/end/<set>.json`, in a directory of its tag for a locale other than en-US. */
export const endFile = (set: string, tag = 'en-US') => join(ROOT, 'test/excel/win/end', localeDir(tag), `${set}.json`)

/** The book as Excel gets it, before its filter and its typed texts. */
export function sentBook(b: EndBook): BookInput {
  const { cells, styles, looks, columns, rows, merges } = b
  return { ...(styles ? { styles } : {}), sheets: [{ name: 'S', cells, ...(looks ? { looks } : {}), ...(columns ? { columns } : {}), ...(rows ? { rows } : {}), ...(merges ? { merges } : {}) }] }
}

/** The book as the engine takes it in a locale: with a filter, a filtered sheet with the rows it hides hidden. */
export function endBook(b: EndBook, locale?: string): BookInput {
  const book = sentBook(b)
  const sheet = book.sheets[0]
  const rows = { ...sheet.rows, ...Object.fromEntries(filteredRows(b).map((r) => [String(r), { hidden: true }])) }
  return { ...book, ...(locale ? { locale } : {}), sheets: [{ ...sheet, ...(b.filter ? { rows, filtered: true } : {}) }] }
}

/** The engine's answers to the cases of one book in a locale (`getEnd`), by case. */
export function endsInBook(b: EndBook, cases: EndCase[], tag: string): Record<string, string> {
  const s = createSession(endBook(b, tag), { clock: { now: '2026-10-13T12:00:00Z', seed: 1 } })
  try {
    return Object.fromEntries(cases.map((c) => [c.id, s.getEnd('S', c.cell, c.direction)]))
  } finally {
    s.close()
  }
}

/** Cells holding their address as text. */
export const named = (...xs: string[]): Record<string, CellInput> => Object.fromEntries(xs.map((a) => [a, { value: a }]))

const SHORT: Record<string, EndDirection> = { u: 'up', d: 'down', l: 'left', r: 'right' }

/** Cases of one group: ids `<group>-<n>`; `list` adds them from text, `A1:d B5:r` (`u`, `d`, `l`, `r`). */
export function collector() {
  const { out, add: put } = groupCollector<EndCase>()
  const add = (group: string, book: string, cell: string, direction: EndDirection) => put(group, { book, cell, direction })
  const list = (group: string, book: string, text: string) => {
    for (const x of text.split(' ')) {
      const [cell, d] = x.split(':')
      add(group, book, cell, SHORT[d])
    }
  }
  return { out, add, list }
}
