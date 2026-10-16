/**
 * Excel's Replace (`Range.Replace`), for Excel for Windows (`bench/windows/run.ts replace <set> [ru|<tag>]`): a set names
 * its books, as the find sets' (the sheet `S`, its cells, columns, hidden rows and columns, merged areas, an autofilter,
 * texts typed once Excel opened it) with the cells read after each case, and its cases, each one Replace over one book:
 * what to find, read with Find's wildcards as Excel's Replace always reads it, what to put, a range or the whole sheet,
 * Match entire cell contents and Match case. Each case runs in a book of its own; then its cells are read, and the book
 * is saved and its file read back for what each cell holds.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CellInput } from '../../../src/contract.js'
import { createSession } from '../../../src/index.js'
import { wildSource } from '../../../src/criteria.js'
import { folded } from '../../../src/find-text.js'
import { heldOf } from '../entry.js'
import { findBook, type FindBook } from '../find/cases.js'
import { localeDir } from '../locales.js'
import type { UiCell } from '../ui/cases.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../../..')

/** A book of a set, and the cells read after a case besides those it holds (a spill's, a merged area's). */
export interface ReplaceBook extends FindBook {
  read?: string[]
}

/** One Replace, as `Range.Replace` takes it. */
export interface ReplaceQuery {
  what: string
  by: string
  /** The range, the whole sheet without it. */
  range?: string
  /** Match entire cell contents (`xlWhole`); a part of the cell's text without it (`xlPart`). */
  whole?: boolean
  matchCase?: boolean
}

export interface ReplaceCase {
  id: string
  book: string
  query: ReplaceQuery
}

/** What Replace answered (true, false, or its error), the cells after it, and their columns' widths (`ColumnWidth`, whether it is the standard one; the sets from `rules-held-2`). */
export interface ReplaceAnswer {
  replaced: boolean | { error: string }
  cells: Record<string, UiCell>
  columns?: Record<string, { columnWidth: number; standard: boolean }>
}

export interface ReplaceSet {
  BOOKS: Record<string, ReplaceBook>
  buildReplaceCases(): ReplaceCase[]
}

export interface StoredReplace {
  meta: Record<string, string | number>
  books: Record<string, ReplaceBook>
  cases: ReplaceCase[]
  answers: Record<string, ReplaceAnswer>
}

/** The stored answers of a set: `test/excel/win/replace/<set>.json`, in a directory of its tag for a locale other than en-US. */
export const replaceFile = (set: string, tag = 'en-US') => join(ROOT, 'test/excel/win/replace', localeDir(tag), `${set}.json`)

/** The cells read after a case: the book's own and those it names, each once. */
export const readOf = (b: ReplaceBook): string[] => [...new Set([...Object.keys(b.cells), ...(b.typed ? Object.keys(b.typed) : []), ...(b.read ?? [])])]

/** The columns of the cells read after a case, by their letters, each once. */
export const columnsOf = (b: ReplaceBook): string[] => [...new Set(readOf(b).map((a) => a.replace(/\d+$/, '')))]

/** The ranges of a query, each area in turn, or the whole sheet. */
export const rangesOf = (q: ReplaceQuery) => (q.range ?? 'A1:XFD1048576').split(',').map((range) => ({ sheet: 'S', range }))

const clock = { now: '2026-10-13T12:00:00Z', seed: 1 }

/** The engine's answer to a case: `replaceInCells` with Replace's options over the query's range, each of its areas in turn, or the whole sheet; the cells read after it, as the book holds them. */
export function replaceInBook(b: ReplaceBook, q: ReplaceQuery, tag: string): Record<string, CellInput | null> {
  const s = createSession(findBook(b, tag), { clock })
  try {
    s.replaceInCells(rangesOf(q), q.what, q.by, { wildcards: true, wholeCell: q.whole, caseSensitive: q.matchCase, skipHidden: true })
    const cells = s.getBook().sheets[0].cells
    return Object.fromEntries(readOf(b).map((a) => [a, cells[a] ?? null]))
  } finally {
    s.close()
  }
}

/**
 * The model of a grid's Replace All (`xlcalc-sheet/src/local/find.ts`), with the matches `?` reads as Find does: the
 * cells Find finds in formulas, hidden ones skipped as it skips them, each one's text as Find reads it with every match
 * of the query (`*` taking as much as it can) replaced, an apostrophe the cell opens with kept, written as typed
 * (`setCellTexts`, no column widened); a text the session refuses refuses all, nothing written. The cells read after
 * it, as the book holds them; or the refusal.
 */
export function modelReplace(b: ReplaceBook, q: ReplaceQuery, tag: string): { refused?: string; cells: Record<string, CellInput | null> } {
  const s = createSession(findBook(b, tag), { clock })
  try {
    const opts = { wildcards: true, wholeCell: q.whole, caseSensitive: q.matchCase }
    const fold = (t: string) => (q.matchCase ? t : folded(t))
    const re = new RegExp(q.whole ? `^(?:${wildSource(fold(q.what))})$` : wildSource(fold(q.what)), 'g')
    const replace = (t: string) => {
      let [out, last] = ['', 0]
      for (const m of fold(t).matchAll(re)) if (m[0] !== '') [out, last] = [out + t.slice(last, m.index) + q.by, m.index + m[0].length]
      return last ? out + t.slice(last) : null
    }
    const found = s.findCells(rangesOf(q), q.what, { ...opts, in: 'edit', skipHidden: true, fields: ['editText', 'formula', 'value'], maxCells: 10000 })
    const entries = found.cells.flatMap((c) => {
      const text = c.formula === undefined && typeof c.value === 'string' ? c.value : c.editText!
      const next = replace(text)
      return next === null ? [] : [{ sheet: 'S', cell: c.cell, text: (c.editText === `'${text}` ? "'" : '') + next }]
    })
    const r = s.setCellTexts(entries, { widen: false })
    const cells = s.getBook().sheets[0].cells
    return { ...(r.ok ? {} : { refused: `${r.cell} ${r.code}` }), cells: Object.fromEntries(readOf(b).map((a) => [a, cells[a] ?? null])) }
  } finally {
    s.close()
  }
}

/** The cells `after` holds otherwise than the book in the locale `tag` held them (`heldOf`: the value or formula and format), null for one left empty. */
export function changedCells(b: ReplaceBook, after: Record<string, CellInput | null | undefined>, tag: string): Record<string, ReturnType<typeof heldOf>> {
  const before = createSession(findBook(b, tag), { clock }).getBook().sheets[0].cells
  return Object.fromEntries(readOf(b).flatMap((a) => (JSON.stringify(heldOf(before[a])) === JSON.stringify(heldOf(after[a] ?? undefined)) ? [] : [[a, heldOf(after[a] ?? undefined)]])))
}
