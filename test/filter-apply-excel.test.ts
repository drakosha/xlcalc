/**
 * The filters of the books Excel made (`test/excel/win/filter-criteria-made/`, `bench/windows/filter-criteria.ts`),
 * applied again: each filter's rows shown, its criteria applied by `applyFilter` without criteria (Reapply), and the
 * rows hidden compared with those Excel hid, the filter mode too.
 */
import { readFileSync, readdirSync } from 'node:fs'
import { expect, it } from 'vitest'
import { createSession, type BookInput, type FilterCriteria, type SheetInput } from '../src/index.js'
import { lines } from '../src/layout/spans.js'
import { readXlsx } from '../src/xlsx.js'

const made = new URL('./excel/win/filter-criteria-made/', import.meta.url)


/** Books whose rows Excel left as they were after an edit that moved the filtered cells: no Reapply gives them. */
const STALE: Record<string, string> = {
  'sheet-moved-part': 'the cells of the filtered column moved out of the filter after it was applied',
}

const hidden = (s: SheetInput) => lines(s.rows, 'r').filter(([, , v]) => v.hidden).flatMap(([a, b]) => Array.from({ length: b - a + 1 }, (_, i) => a + i + 1))

interface Case {
  name: string
  sheet: number
  target: { sheet: string } | { table: string }
}

const cases: Case[] = []
const books = new Map<string, BookInput>()
for (const file of readdirSync(made).filter((n) => n.endsWith('.xlsx')).sort()) {
  const name = file.slice(0, -5)
  if (STALE[name]) continue
  const { book } = await readXlsx(readFileSync(new URL(file, made)))
  books.set(name, book)
  book.sheets.forEach((s, i) => {
    const filters: [Case['target'], FilterCriteria][] = []
    if (s.autoFilter?.criteria) filters.push([{ sheet: s.name }, s.autoFilter.criteria])
    for (const t of s.tables ?? []) if (t.filter) filters.push([{ table: t.name }, t.filter])
    for (const [target, criteria] of filters) {
      const kinds = (criteria.columns ?? []).flatMap((c) => (c.filter ? [c.filter.type] : []))
      if (kinds.length) cases.push({ name, sheet: i, target })
    }
  })
}

it('finds the filters of every kind', () => {
  expect(cases.length).toBe(113)
})

it.each(cases.map((c) => [c.name, JSON.stringify(c.target), c] as const))('%s %s: the rows Excel hid', (_, __, c) => {
  const book = books.get(c.name)!
  const shown = { ...book, sheets: book.sheets.map((s, i) => (i === c.sheet ? { ...s, rows: undefined, filtered: undefined } : s)) }
  const session = createSession(shown, { clock: { now: '2026-11-06T12:00:00Z', seed: 1 } })
  expect(session.editStructure({ op: 'applyFilter', ...c.target }).ok).toBe(true)
  const s = session.getBook().sheets[c.sheet]
  expect([hidden(s), !!s.filtered]).toEqual([hidden(book.sheets[c.sheet]), !!book.sheets[c.sheet].filtered])
})
