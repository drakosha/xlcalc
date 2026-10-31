/**
 * Slicers and timelines against Excel for Windows, the set `slicer` (`test/excel/win/pivot/slicer.json` and its saved
 * books). Each case made again through a session's operations (`bench/coverage/pivot/slicer-predict.ts`) gives the cells
 * Excel showed in every PivotTable and the items its slicers showed (name, selected, has data, in their order) or a
 * timeline's dates and level. Each book Excel saved holds a slicer's choice as its field's hidden items on every
 * PivotTable it filters, and a timeline's dates as a filter of its field, as the session writes them; laid out again
 * from its cache, a book's slicers get back the items Excel saved, in its order.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { beforeAll, describe, expect, it } from 'vitest'
import type { PivotPhase } from '../bench/coverage/pivot/cases.js'
import { predict } from '../bench/coverage/pivot/slicer-predict.js'
import { buildPivotCases } from '../bench/coverage/pivot/slicer.js'
import { registerPivot, type BookInput } from '../src/index.js'
import { refreshPivots } from '../src/pivot/engine.js'
import { pivotEngine } from '../src/pivot/index.js'
import { fieldIndex, selectionLevel, timelineBounds, withSlicerChoice, withTimelineChoice } from '../src/structure/slicer-choice.js'
import { readXlsx } from '../src/xlsx.js'

const DIR = new URL('./excel/win/pivot/', import.meta.url)
const answers = (JSON.parse(readFileSync(new URL('slicer.json', DIR), 'utf8')) as { answers: Record<string, { phases: PivotPhase[] }> }).answers

const rounded = (v: unknown): unknown => (typeof v === 'number' ? Number(v.toPrecision(15)) : Array.isArray(v) ? v.map(rounded) : v)

const books = async () => Promise.all(readdirSync(new URL('slicer/', DIR)).map(async (f) => [f, (await readXlsx(readFileSync(new URL(`slicer/${f}`, DIR)))).book] as const))

beforeAll(() => registerPivot(pivotEngine))

describe('slicers and timelines against Excel for Windows', () => {
  for (const c of buildPivotCases())
    it(c.id, async () => {
      const got = await predict(c)
      const want = answers[c.id].phases
      expect(got.length).toBe(want.length)
      got.forEach((phase, k) => {
        for (const [name, area] of Object.entries(want[k].pivots)) {
          if ('error' in area) throw new Error(`${c.id}: Excel's error ${area.error}`)
          expect(phase.pivots[name].area).toBe(area.area.replace(/\$/g, ''))
          expect(rounded(phase.pivots[name].values)).toEqual(rounded(area.values))
        }
        expect(phase.slicers).toEqual(want[k].slicers ?? {})
      })
    })

  it("each book holds the choice as the session writes it, and its slicers' items as laid out again", async () => {
    for (const [file, book] of await books()) {
      // the PivotTables as the session leaves them after a choice: the same as Excel saved
      let chosen: BookInput = book
      for (const sc of book.slicerCaches ?? []) chosen = withSlicerChoice(chosen, sc)
      for (const tc of book.timelineCaches ?? []) chosen = withTimelineChoice(chosen, tc)
      expect(chosen.sheets.map((s) => s.pivotTables), file).toEqual(book.sheets.map((s) => s.pivotTables))
      for (const tc of book.timelineCaches ?? []) {
        const pc = book.pivotCaches!.find((c) => c.id === tc.cache)!
        expect(timelineBounds(pc, fieldIndex(pc, tc.field)), file).toEqual(tc.bounds)
        for (const t of book.sheets.flatMap((s) => s.timelines ?? []).filter((x) => x.cache === tc.name)) expect(selectionLevel(tc.selection!), file).toBe(t.selectionLevel)
      }
      const r = refreshPivots(book, undefined, () => [], 0, false)
      if (!r.ok) throw new Error(`${file}: ${r.message}`)
      expect(r.book.slicerCaches, file).toEqual(book.slicerCaches)
    }
  })
})
