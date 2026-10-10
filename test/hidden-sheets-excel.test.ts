/**
 * Hidden sheets against Excel for Mac and Excel for Windows (`test/excel/hidden/`, `test/excel/win/hidden/`, from the
 * step `hidden`): every file the writer wrote opens without the repair window, with the first visible sheet active; Excel
 * read each sheet's visibility and the read cells' values the engine holds after the case, refused what the engine
 * refuses, and the reader reads every file Excel saved with the visibility Excel read. The engine's predictions
 * (`bench/coverage/hidden/<set>.predicted.json` and the set's `RAW_GUESS`) were committed before Excel ran the set; the
 * ones Excel did not confirm are the set's `MISSED`.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredHidden } from '../bench/coverage/hidden.js'
import { engineAnswer, excelRefused, hiddenFile, visibleOf, type HiddenAnswer, type HiddenCase, type HiddenCases } from '../bench/coverage/hidden/cases.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import { readXlsx } from '../src/xlsx.js'

interface HiddenModule {
  buildHiddenCases(): HiddenCases
  RAW_GUESS?: Record<string, { repaired: boolean; opened?: string }>
  MISSED?: readonly string[]
}

const PLATFORMS = [
  ['Excel for Mac', ''],
  ['Excel for Windows', 'win/'],
] as const

/** The stored sets of a platform. */
function stored(root: string): (StoredHidden & { name: string })[] {
  const dir = new URL(`./excel/${root}hidden/`, import.meta.url)
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((f) => f.endsWith('.json'))
    .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredHidden) }))
}

/** What a prediction said that Excel's answer does not hold, as text; none when it held. */
function missed(set: HiddenCases, c: HiddenCase, predicted: Record<string, unknown>, guess: HiddenModule['RAW_GUESS'], answer: HiddenAnswer | undefined, rejected: boolean): string[] {
  const g = guess?.[c.id]
  if (g) return g.repaired !== rejected ? [`repaired ${g.repaired}, Excel ${rejected}`] : g.opened !== undefined && g.opened !== answer?.opened ? [`opened ${g.opened}, Excel ${answer?.opened}`] : []
  if (!answer) return ['Excel did not load the book']
  const p = predicted as { refused?: string; opened?: string; sheets?: unknown; values?: unknown }
  const out: string[] = []
  if (!!p.refused !== excelRefused(set, c, answer)) out.push(`refused ${p.refused ?? 'no'}, Excel ${answer.error ?? 'took it'}`)
  if (p.opened !== undefined && p.opened !== answer.opened) out.push(`opened ${p.opened}, Excel ${answer.opened}`)
  if (p.sheets && JSON.stringify(p.sheets) !== JSON.stringify(answer.sheets)) out.push(`sheets ${JSON.stringify(p.sheets)}, Excel ${JSON.stringify(answer.sheets)}`)
  if (p.values && JSON.stringify(p.values) !== JSON.stringify(answer.values)) out.push(`values ${JSON.stringify(p.values)}, Excel ${JSON.stringify(answer.values)}`)
  return out
}

describe.each(PLATFORMS)('hidden sheets against %s', (_, root) => {
  it('every set is answered', () => expect(stored(root).map((s) => s.name)).toEqual(['sheets-held']))
  for (const set of stored(root)) {
    describe(set.name, async () => {
      const mod = (await import(`../bench/coverage/hidden/${set.name}.ts`)) as HiddenModule
      const built = mod.buildHiddenCases()
      const predicted = JSON.parse(readFileSync(new URL(`../bench/coverage/hidden/${set.name}.predicted.json`, import.meta.url), 'utf8')) as Record<string, Record<string, unknown>>

      it('the stored cases are the generator\'s, sent as the writer writes them now', async () => {
        expect(set.cases).toEqual(built.cases)
        expect(set.books).toEqual(built.books)
        for (const c of built.cases) expect(set.sent[c.id], c.id).toBe(await partsHash(await hiddenFile(c, built)))
      })

      it('every written file opens without the repair window; the one whose every sheet is hidden does not', () => {
        expect(Object.keys(set.rejected ?? {})).toEqual(built.cases.filter((c) => c.raw === 'allHidden').map((c) => c.id))
        expect(built.cases.filter((c) => !set.answers[c.id] && !set.rejected?.[c.id]).map((c) => c.id)).toEqual([])
      })

      it.each(built.cases.filter((c) => set.answers[c.id]).map((c) => [c.id, c] as const))('%s: Excel holds what the engine holds', async (_, c) => {
        const a = set.answers[c.id]
        // a file changed by hand is no book the writer writes: Excel shows the hidden sheet its active tab names
        if (!c.raw) {
          const e = engineAnswer(built, c)
          expect(excelRefused(built, c, a), `refused: the engine ${e.refused ?? 'no'}, Excel ${a.error ?? 'answered nothing'}`).toBe(!!e.refused)
          expect(a.sheets).toEqual(e.sheets)
          expect(a.values).toEqual(e.values)
          // the writer makes the first visible sheet active, which Excel shows on opening
          expect(a.opened).toBe(built.books[c.book].sheets.find((s) => !s.hidden)!.name)
        }
        const saved = await readXlsx(new Uint8Array(readFileSync(new URL(`./excel/${root}hidden/${set.name}/${c.id}.xlsx`, import.meta.url))))
        expect(saved.book.sheets.map((s): [string, number] => [s.name, visibleOf(s)])).toEqual(a.sheets)
        expect(saved.issues.filter((i) => i.code === 'hidden')).toEqual([])
      })

      it('the predictions Excel did not confirm are the set\'s MISSED', () => {
        const misses = built.cases.filter((c) => missed(built, c, predicted[c.id], mod.RAW_GUESS, set.answers[c.id], !!set.rejected?.[c.id]).length).map((c) => c.id)
        expect(misses).toEqual([...(mod.MISSED ?? [])])
      })
    })
  }
})
