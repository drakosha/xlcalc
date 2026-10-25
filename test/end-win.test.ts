/**
 * `getEnd` against Excel for Windows' `Range.End` (`test/excel/win/end/`, the step `end` of `bench/windows` over the
 * sets of `bench/coverage/end/`): each case's cell and direction over its book in each locale Excel ran it in, in a
 * synchronous session, in the same thread and in a thread of the pool. Known differences would be in `END_KNOWN` with
 * their reasons (none); the predictions committed before Excel ran each set are kept with how many cases Excel confirmed.
 */
import { readFileSync } from 'node:fs'
import { afterAll, describe, expect, it } from 'vitest'
import { endBook, type EndAnswer, type EndSet, type StoredEnd } from '../bench/coverage/end/cases.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import { localeSets } from './stored-sets.js'
import './locales.js'

afterAll(closeCalcs)

/** `<set>:<id>` → why the engine differs from Excel there, in every locale. */
const END_KNOWN: Record<string, string> = {}

/** How many cases of each set's prediction by `getEnd` Excel for Windows confirmed, by `<tag>:<set>`. */
const CONFIRMED: Record<string, number> = {
  'en-US:rules-held': 167,
  'en-US:rules-held-2': 19,
  'ru-RU:ru-held': 4,
}

const RUNS = localeSets<StoredEnd>('win/end')

const excelCell = (a: EndAnswer) => ('cell' in a ? a.cell : a.error)

describe('Range.End (Excel for Windows)', () => {
  it('runs in both locales, Excel answered every case with a cell, and every known difference is a stored case', () => {
    expect(new Set(RUNS.map((r) => r.tag))).toEqual(new Set(['en-US', 'ru-RU']))
    expect(RUNS.flatMap((r) => Object.entries(r.stored.answers).filter(([, a]) => !('cell' in a)))).toEqual([])
    const ids = new Set(RUNS.flatMap((r) => r.stored.cases.map((c) => `${r.set}:${c.id}`)))
    expect(Object.keys(END_KNOWN).filter((k) => !ids.has(k))).toEqual([])
  })
  describe.each(KINDS)('%s', (kind) => {
    it.each(RUNS.map((r) => [`${r.tag}:${r.set}`, r] as const))('%s: the cell Excel came to, but the known differences', async (_, run) => {
      const open = await opener(kind)
      const off: unknown[] = []
      for (const [name, book] of Object.entries(run.stored.books)) {
        const cases = run.stored.cases.filter((c) => c.book === name)
        if (!cases.length) continue
        const s = await open(endBook(book, run.tag), { clock: CLOCK })
        for (const c of cases) {
          const [want, got] = [excelCell(run.stored.answers[c.id]), await s.getEnd('S', c.cell, c.direction)]
          const known = END_KNOWN[`${run.set}:${c.id}`]
          // a known difference is still one, and every other case is Excel's
          if ((got === want) === !known) continue
          off.push({ id: c.id, cell: c.cell, direction: c.direction, excel: want, engine: got, known })
        }
        await s.close()
      }
      expect(off).toEqual([])
    }, 60_000)
  })
  it.each(RUNS.map((r) => [`${r.tag}:${r.set}`, r] as const))('%s: the set as it is now, and the predictions written before Excel ran it', async (key, run) => {
    const set = (await import(`../bench/coverage/end/${run.set}.ts`)) as EndSet
    expect([run.stored.books, run.stored.cases]).toEqual([set.BOOKS, set.buildEndCases()])
    const predicted = JSON.parse(readFileSync(new URL(`../bench/coverage/end/${run.set}.predicted.json`, import.meta.url), 'utf8')).locales[run.tag] as Record<string, string>
    const confirmed = run.stored.cases.filter((c) => predicted[c.id] === excelCell(run.stored.answers[c.id]))
    expect(confirmed.length).toBe(CONFIRMED[key])
  })
})
