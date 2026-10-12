/**
 * The outline commands against Excel for Windows (`test/excel/win/outline/`, the step `outline` of `bench/windows` over
 * the sets of `bench/coverage/outline/`): each case's book, then each command as `editStructure`'s outline operation (lines
 * hidden or shown by hand as `setLayout` changes them), and after each step the window's lines (level, hidden, and a
 * summary line's `ShowDetail` as `outlineGroups` gives its group collapsed) and the file `writeXlsx` writes (levels,
 * hidden lines, collapsed marks, settings) as Excel read and saved them; a refusal where Excel refused the command. The
 * reader reads the book Excel saved at the end with the groups the engine holds. Known differences are in
 * `OUTLINE_KNOWN` with their reasons; the predictions committed before Excel ran each set are kept with how many cases
 * Excel confirmed.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { isDeepStrictEqual } from 'node:util'
import { engineSteps, engineViews, excelView, type OutlineAnswer, type OutlineBook, type OutlineCase, type OutlineView } from '../bench/coverage/outline/cases.js'
import { localeDir, TAGS } from '../bench/coverage/locales.js'
import { outlineGroups } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'

interface StoredOutline {
  books: Record<string, OutlineBook>
  cases: OutlineCase[]
  answers: Record<string, OutlineAnswer>
}

const OUTLINE_REASONS = {
  xfd: "columns grouped in a span of several columns that reaches XFD, as a file's `<col max=\"16384\">` holds them: Excel's level buttons leave them as they are and its object model refuses Ungroup over any of them; the engine takes them as other columns (XFD grouped alone, by Group or a file, both hide and ungroup)",
  byHand: 'rows hidden and shown by hand (`setLayout`): Excel reads the marks again only for the groups whose rows changed, so a summary row shown by hand over a group it was hidden with stays without a mark; the engine marks every summary row whose group is all hidden, as Excel opens a file',
} as const

/** `<set>:<id>` → why the engine differs from Excel there. */
const OUTLINE_KNOWN: Record<string, keyof typeof OUTLINE_REASONS> = {
  'rules-held:edge-7': 'xfd',
  'rules-held-2:levels-7': 'xfd',
  'rules-held-2:levels-9': 'xfd',
  'rules-held-2:levels-10': 'xfd',
  'rules-held-2:levels-11': 'xfd',
  'rules-held-3:xfd-4': 'xfd',
  'rules-held-3:xfd-5': 'xfd',
  'rules-held-3:xfd-6': 'xfd',
  'rules-held-3:levels-2': 'xfd',
  'rules-held-2:hand-3': 'byHand',
  'rules-held-3:hand-3': 'byHand',
  'rules-held-3:hand-5': 'byHand',
}

/** How many cases of each set's prediction Excel for Windows confirmed, by `<tag>:<set>`: the first two by the grid's own model, the others by the engine. */
const CONFIRMED: Record<string, number> = {
  'en-US:rules-held': 49,
  'ru-RU:ru-held': 2,
  'en-US:rules-held-2': 37,
  'en-US:rules-held-3': 9,
}

const RUNS = TAGS.flatMap((tag) => {
  const dir = new URL(`./excel/win/outline/${localeDir(tag)}`, import.meta.url)
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((f) => f.endsWith('.json'))
    .map((f) => ({ tag, set: f.slice(0, -5), stored: JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredOutline }))
})

const predicted = (set: string): Record<string, OutlineView[]> => JSON.parse(readFileSync(new URL(`../bench/coverage/outline/${set}.predicted.json`, import.meta.url), 'utf8')).cases

describe('outline commands (Excel for Windows)', () => {
  it('runs in en-US and ru-RU; every confirmed count names a stored set, every known difference a stored case', () => {
    expect(new Set(RUNS.map((r) => r.tag))).toEqual(new Set(['en-US', 'ru-RU']))
    expect(Object.keys(CONFIRMED).filter((k) => !RUNS.some((r) => `${r.tag}:${r.set}` === k))).toEqual([])
    const ids = new Set(RUNS.flatMap((r) => r.stored.cases.map((c) => `${r.set}:${c.id}`)))
    expect(Object.keys(OUTLINE_KNOWN).filter((k) => !ids.has(k))).toEqual([])
  })

  for (const run of RUNS)
    describe(`${run.tag} ${run.set}`, () => {
      it("each step as Excel read and saved it, each refusal as Excel's; a known difference differs", async () => {
        for (const c of run.stored.cases) {
          const excel = run.stored.answers[c.id].steps.map((s) => excelView(s, c.window))
          const engine = await engineViews(run.stored.books[c.book], c)
          if (OUTLINE_KNOWN[`${run.set}:${c.id}`]) expect(isDeepStrictEqual(engine, excel), c.id).toBe(false)
          else expect([c.id, engine]).toEqual([c.id, excel])
        }
      })

      it('the reader reads the book Excel saved with the groups the engine holds', async () => {
        for (const c of run.stored.cases) {
          if (OUTLINE_KNOWN[`${run.set}:${c.id}`]) continue
          const saved = new URL(`./excel/win/outline/${localeDir(run.tag)}${run.set}/${c.id}.xlsx`, import.meta.url)
          const read = (await readXlsx(new Uint8Array(readFileSync(saved)))).book.sheets[0]
          expect([c.id, outlineGroups(read)]).toEqual([c.id, outlineGroups(engineSteps(run.stored.books[c.book], c).at(-1)!.book.sheets[0])])
        }
      })

      it('the prediction committed before Excel ran the set', () => {
        const p = predicted(run.set)
        const confirmed = run.stored.cases.filter((c) => isDeepStrictEqual(p[c.id], run.stored.answers[c.id].steps.map((s) => excelView(s, c.window)))).length
        expect(confirmed).toBe(CONFIRMED[`${run.tag}:${run.set}`] ?? run.stored.cases.length)
      })
    })
})
