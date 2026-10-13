/**
 * `findCells` against Excel for Windows' Find (`test/excel/win/find/`, the step `find` of `bench/windows` over the sets
 * of `bench/coverage/find/`): each case's query over its book in each locale Excel ran it in, with Find's options
 * (`wildcards`, `wholeCell`, `byColumns`, `caseSensitive`, `skipHidden`; Look in Formulas as `in: 'edit'`, Values as
 * `in: 'text'`) over the case's range, each of its areas in turn, or the whole sheet. The cells and their order are
 * Excel's once round: `Range.Find` from after the range's last cell, or from the case's cell (`after`), and `FindNext`
 * until a cell found before. Known differences are in `FIND_KNOWN` with
 * their reasons; the predictions committed before Excel ran each set are kept with how many cases Excel confirmed.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { findInBook, type FindCase, type StoredFind } from '../bench/coverage/find/cases.js'
import { localeDir, TAGS } from '../bench/coverage/locales.js'
import './locales.js'

const FIND_REASONS = {
  empty: "an empty query: Excel's Find finds empty cells, the range's or the sheet's from its first; findCells finds what cells hold, and nothing for an empty query",
  tilde: "a query of `~` alone, which reads as nothing: Excel finds only the cell it starts after (the sheet's last, empty); findCells nothing, as for an empty query",
  oneCell: "a range of one cell: Excel's Range.Find searches the whole sheet from after it, as the dialog does with one cell selected, and FindNext refuses it; findCells searches the ranges given",
  columnsRound: "by columns on a filtered sheet Excel's FindNext going round passes the first column of the first block over the whole sheet (and over B2:E11 in rules-held-3, where `back` names the cell it came back to), so Find from a cell after those cells never gives them; findCells gives every cell once round",
} as const

/** `<set>:<id>` → why the engine differs from Excel there, in every locale. */
const FIND_KNOWN: Record<string, keyof typeof FIND_REASONS> = {
  'rules-held:empty-1': 'empty',
  'rules-held:empty-2': 'empty',
  'rules-held:empty-4': 'empty',
  'rules-held:tilde-11': 'tilde',
  'rules-held:tilde-12': 'tilde',
  'rules-held:order-7': 'oneCell',
  'rules-held:order-8': 'oneCell',
  'rules-held-4:filtered-5': 'columnsRound',
  'rules-held-4:filtered-6': 'columnsRound',
  'rules-held-4:filtered-7': 'columnsRound',
  'rules-held-4:filtered-8': 'columnsRound',
  'rules-held-5:filtered-13': 'columnsRound',
  'rules-held-5:filtered-14': 'columnsRound',
  'rules-held-5:filtered-15': 'columnsRound',
  'rules-held-5:filtered-16': 'columnsRound',
  'rules-held-5:filtered-17': 'columnsRound',
  'rules-held-5:filtered-18': 'columnsRound',
  'rules-held-5:filtered-19': 'columnsRound',
  'rules-held-5:filtered-20': 'columnsRound',
  'rules-held-5:filtered-21': 'columnsRound',
}

/** How many cases of each set's prediction Excel for Windows confirmed, by `<tag>:<set>`: `rules-held` by the grid's own model, the others by `findCells`. */
const CONFIRMED: Record<string, number> = {
  'en-US:rules-held': 136,
  'ru-RU:rules-held': 136,
  'en-US:rules-held-2': 45,
  'ru-RU:rules-held-2': 45,
  'en-US:rules-held-3': 32,
  'ru-RU:rules-held-3': 32,
  'en-US:rules-held-4': 22,
  'ru-RU:rules-held-4': 22,
  'en-US:rules-held-5': 15,
  'ru-RU:rules-held-5': 15,
}

const RUNS = TAGS.flatMap((tag) => {
  const dir = new URL(`./excel/win/find/${localeDir(tag)}`, import.meta.url)
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((f) => f.endsWith('.json'))
    .map((f) => ({ tag, set: f.slice(0, -5), stored: JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredFind }))
})

/** The cells Excel found, once round; null when it answered an error (a cell to begin after outside the range, which findCells refuses too) or found more than it read. */
function excelCells(run: (typeof RUNS)[number], c: FindCase): string[] | null {
  const a = run.stored.answers[c.id]
  return 'found' in a && !a.more ? a.found.map((x) => x.cell) : null
}

describe('Find (Excel for Windows)', () => {
  it('runs in both locales, and every known difference is a stored case', () => {
    expect(new Set(RUNS.map((r) => r.tag))).toEqual(new Set(['en-US', 'ru-RU']))
    const ids = new Set(RUNS.flatMap((r) => r.stored.cases.map((c) => `${r.set}:${c.id}`)))
    expect(Object.keys(FIND_KNOWN).filter((k) => !ids.has(k))).toEqual([])
  })
  it.each(RUNS.map((r) => [`${r.tag}:${r.set}`, r] as const))('%s: the cells Excel found, in its order, but the known differences', (_, run) => {
    const off = run.stored.cases.flatMap((c) => {
      const want = excelCells(run, c)
      const got = findInBook(run.stored.books[c.book], c.query, run.tag)
      const same = JSON.stringify(got) === JSON.stringify(want)
      const known = FIND_KNOWN[`${run.set}:${c.id}`]
      // a known difference is still one, and every other case is Excel's
      return same === !known ? [] : [{ id: c.id, query: c.query, excel: want, engine: got, known }]
    })
    expect(off).toEqual([])
  })
  it.each(RUNS.map((r) => [`${r.tag}:${r.set}`, r] as const))('%s: the predictions written before Excel ran it', (key, run) => {
    const predicted = JSON.parse(readFileSync(new URL(`../bench/coverage/find/${run.set}.predicted.json`, import.meta.url), 'utf8')).locales[run.tag] as Record<string, string[]>
    const confirmed = run.stored.cases.filter((c) => JSON.stringify(predicted[c.id]) === JSON.stringify(excelCells(run, c)))
    expect(confirmed.length).toBe(CONFIRMED[key])
  })
})
