/**
 * Excel's Find (`Range.Find` and `FindNext`), for Excel for Windows (`bench/windows/run.ts find <set> [ru|<tag>]`): a set
 * names its books, each the sheet `S` with its cells, columns, hidden rows and columns, merged areas, an autofilter and
 * texts typed once Excel opened it, and its cases, each a query over one book: what to find, a range or the whole
 * sheet, Look in (formulas or values), Match entire cell contents, Search by columns and Match case. Excel finds from
 * the cell after the range's last one, so its first cell comes first, and goes on with `FindNext` until it comes back
 * to a cell found before: the cells in that order, each with its `Formula`, `FormulaLocal` and `Text`, which show what matched.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { BookInput, CellInput, ColumnInput, RowInput } from '../../../src/contract.js'
import { formatAddress } from '../../../src/model.js'
import { parseRange } from '../../../src/structure/area.js'
import { localeDir } from '../locales.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../../..')

/** A book of a set: the sheet `S`. */
export interface FindBook {
  cells: Record<string, CellInput>
  columns?: Record<string, ColumnInput>
  rows?: Record<string, RowInput>
  merges?: string[]
  /** An autofilter over `range` keeping the rows whose `field`-th column holds `keep`: Excel hides the others as it applies it, the engine's book holds them hidden on a filtered sheet. */
  filter?: { range: string; field: number; keep: string }
  /** Texts typed into cells once Excel opened the book (`Formula2Local`), for what a file does not hold (a text typed after an apostrophe); the engine's book holds the cell of `cells`. */
  typed?: Record<string, string>
}

/** What a case finds and how, as `Range.Find` takes it. */
export interface FindQuery {
  what: string
  /** The range searched, the whole sheet without it. */
  range?: string
  lookIn: 'formulas' | 'values'
  /** Match entire cell contents (`xlWhole`); a part of the cell's text without it (`xlPart`). */
  whole?: boolean
  /** Search by columns (`xlByColumns`); by rows without it. */
  byColumns?: boolean
  matchCase?: boolean
}

export interface FindCase {
  id: string
  book: string
  query: FindQuery
}

/** A cell Excel found, with what its object model reads of it. */
export interface FoundCell {
  cell: string
  formula: string
  formulaLocal: string
  text: string
}

/**
 * The cells Excel found in its order; `more` when it found more than `MAX_FOUND`, `next` the error `FindNext` answered,
 * `back` the cell `FindNext` came back to when it was not the first (it went round again from the second). Or the error
 * Find answered.
 */
export type FindAnswer = { found: FoundCell[]; more?: true; next?: string; back?: string } | { error: string }

export interface FindSet {
  BOOKS: Record<string, FindBook>
  buildFindCases(): FindCase[]
}

/** The cells a case reads at most: Find over blank cells would go on through the sheet. */
export const MAX_FOUND = 60

export interface StoredFind {
  meta: Record<string, string | number>
  books: Record<string, FindBook>
  cases: FindCase[]
  answers: Record<string, FindAnswer>
}

/** The stored answers of a set: `test/excel/win/find/<set>.json`, in a directory of its tag for a locale other than en-US. */
export const findFile = (set: string, tag = 'en-US') => join(ROOT, 'test/excel/win/find', localeDir(tag), `${set}.json`)

/** The rows an autofilter hides: those of its range below the header whose cell of the field is not the kept text. */
export function filteredRows(b: FindBook): number[] {
  if (!b.filter) return []
  const x = parseRange(b.filter.range)!
  const col = x.c1 + b.filter.field - 1
  const out: number[] = []
  for (let r = x.r1 + 1; r <= x.r2; r++) if (b.cells[formatAddress(r, col)]?.value !== b.filter.keep) out.push(r + 1)
  return out
}

/** The book as Excel gets it, before its filter and its typed texts. */
export function sentBook(b: FindBook): BookInput {
  return { sheets: [{ name: 'S', cells: b.cells, ...(b.columns ? { columns: b.columns } : {}), ...(b.rows ? { rows: b.rows } : {}), ...(b.merges ? { merges: b.merges } : {}) }] }
}

/** The book as the engine takes it in a locale: with a filter, a filtered sheet with the rows it hides hidden. */
export function findBook(b: FindBook, locale?: string): BookInput {
  const sheet = sentBook(b).sheets[0]
  const rows = { ...sheet.rows, ...Object.fromEntries(filteredRows(b).map((r) => [String(r), { hidden: true }])) }
  return { ...(locale ? { locale } : {}), sheets: [{ ...sheet, ...(b.filter ? { rows, filtered: true } : {}) }] }
}

/** Cases of one group: ids `<group>-<n>`. */
export function collector() {
  const out: FindCase[] = []
  const count = new Map<string, number>()
  const add = (group: string, book: string, query: FindQuery) => {
    const n = (count.get(group) ?? 0) + 1
    count.set(group, n)
    out.push({ id: `${group}-${n}`, book, query })
  }
  return { out, add }
}
