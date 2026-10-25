/**
 * `replaceInCells` against Excel for Windows' Replace (`test/excel/win/replace/`, the step `replace` of `bench/windows`
 * over the sets of `bench/coverage/replace/`): each case's Replace over its book in each locale Excel ran it in, with the
 * options of Excel's dialog (`wildcards`, `wholeCell`, `caseSensitive`, `skipHidden`) over the case's range, each of its
 * areas in turn, or the whole sheet. What each cell read after holds (its value or formula and its format, a formula
 * compared by its canonical text, a structured reference with its table's name or without it) is what the file Excel
 * saved holds, and no column Excel left at the standard width has one. Known differences are in `REPLACE_KNOWN` with
 * their reasons; the predictions committed before Excel ran each set are kept with how many cases Excel confirmed.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { held } from '../bench/coverage/entry/compare.js'
import { findBook } from '../bench/coverage/find/cases.js'
import { columnsOf, readOf, replaceInBook, type ReplaceBook, type ReplaceCase, type StoredReplace } from '../bench/coverage/replace/cases.js'
import { createSession, type CellInput } from '../src/index.js'
import { localeSets } from './stored-sets.js'
import './locales.js'

const REPLACE_REASONS = {
  prefix: "a text that reads as a number, from a file without the prefix character: Excel enters it again as typed and it becomes a number; the book keeps no mark of an apostrophe, and replaceInCells keeps a text that reads as something else a text, as typed after an apostrophe",
  oneCell: "a range of one cell: Excel's Range.Replace replaces in the whole sheet, as Find searches it; replaceInCells replaces in the ranges given",
} as const

/** `[<tag>:]<set>:<id>` → why the engine differs from Excel there, in that locale or in every one. */
const REPLACE_KNOWN: Record<string, keyof typeof REPLACE_REASONS> = {
  'rules-held:entry-3': 'prefix',
  'rules-held:entry-4': 'prefix',
  'ru-RU:rules-held:entry-8': 'prefix',
  'rules-held:look-26': 'oneCell',
}

/** How many cases of each set's prediction Excel for Windows confirmed, by `<tag>:<set>`: `rules-held` by a grid's model, the others by `replaceInCells`. */
const CONFIRMED: Record<string, number> = {
  'en-US:rules-held': 51,
  'ru-RU:rules-held': 50,
  'en-US:rules-held-2': 20,
  'ru-RU:rules-held-2': 20,
  'en-US:rules-held-3': 9,
  'ru-RU:rules-held-3': 9,
}

const RUNS = localeSets<StoredReplace>('win/replace')

/** A cell as compared: its value or formula and format; a structured reference without its table's name, as the engine writes one in its table. */
const key = (c: CellInput | null | undefined, b: ReplaceBook, tag: string) => (b.tables ?? []).reduce((t, x) => t.replaceAll(`${x.name}[`, '['), JSON.stringify(held(c ?? null, tag)))

/** A book's cells as the engine holds them in a locale, before any case. */
function cellsBefore(b: ReplaceBook, tag: string): Record<string, CellInput> {
  const s = createSession(findBook(b, tag))
  try {
    return s.getBook().sheets[0].cells
  } finally {
    s.close()
  }
}

/** The cells read after a case that hold otherwise than the book did before it, by their keys. */
function changed(b: ReplaceBook, after: (a: string) => CellInput | null | undefined, tag: string): Record<string, string> {
  const before = cellsBefore(b, tag)
  return Object.fromEntries(readOf(b).flatMap((a) => (key(after(a), b, tag) === key(before[a], b, tag) ? [] : [[a, key(after(a), b, tag)]])))
}

describe('Replace (Excel for Windows)', () => {
  it('runs in both locales, and every known difference is a stored case', () => {
    expect(new Set(RUNS.map((r) => r.tag))).toEqual(new Set(['en-US', 'ru-RU']))
    const ids = new Set(RUNS.flatMap((r) => r.stored.cases.flatMap((c) => [`${r.set}:${c.id}`, `${r.tag}:${r.set}:${c.id}`])))
    expect(Object.keys(REPLACE_KNOWN).filter((k) => !ids.has(k))).toEqual([])
  })
  it.each(RUNS.map((r) => [`${r.tag}:${r.set}`, r] as const))('%s: what each cell holds after it, as Excel left it, but the known differences', (_, run) => {
    const off = run.stored.cases.flatMap((c: ReplaceCase) => {
      const b = run.stored.books[c.book]
      const a = run.stored.answers[c.id]
      const got = replaceInBook(b, c.query, run.tag)
      const cells = readOf(b).filter((x) => key(got[x], b, run.tag) !== key(a.cells[x].cell, b, run.tag))
      // the engine widens no column: Excel left every column at the width it had
      const widened = columnsOf(b).filter((k) => a.columns && !a.columns[k].standard && b.columns?.[k]?.width === undefined)
      const known = REPLACE_KNOWN[`${run.tag}:${run.set}:${c.id}`] ?? REPLACE_KNOWN[`${run.set}:${c.id}`]
      const same = !cells.length && !widened.length
      return same === !known ? [] : [{ id: c.id, query: c.query, cells: cells.map((x) => [x, key(a.cells[x].cell, b, run.tag), key(got[x], b, run.tag)]), widened, known }]
    })
    expect(off).toEqual([])
  })
  it.each(RUNS.map((r) => [`${r.tag}:${r.set}`, r] as const))('%s: the predictions written before Excel ran it', (k, run) => {
    const predicted = JSON.parse(readFileSync(new URL(`../bench/coverage/replace/${run.set}.predicted.json`, import.meta.url), 'utf8')).locales[run.tag] as Record<string, { changed: Record<string, CellInput | null> }>
    const confirmed = run.stored.cases.filter((c) => {
      const b = run.stored.books[c.book]
      const p = predicted[c.id].changed
      const excel = changed(b, (x) => run.stored.answers[c.id].cells[x].cell, run.tag)
      const before = cellsBefore(b, run.tag)
      return JSON.stringify(changed(b, (x) => (x in p ? p[x] : before[x]), run.tag)) === JSON.stringify(excel)
    })
    expect(confirmed.length).toBe(CONFIRMED[k])
  })
})
