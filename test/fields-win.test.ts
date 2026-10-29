/**
 * Fields of a workbook the model holds for a host, against Excel for Windows (`test/excel/win/fields/`,
 * `bench/coverage/fields/`): Excel read from each file `writeXlsx` or `writeKept` wrote what the workbook holds, and the
 * reader reads it back from the file Excel saved: a split window with the same lines, which Excel moves to its screen's
 * nearest pixel.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { bookCalculation, bookFields, bookNotes, bookPanes, bookTurned, caseBook, fieldsFile, savedDir, seenPanes, splitOf, type FieldsAnswer, type FieldsSet } from '../bench/coverage/fields/cases.js'
import { readXlsx } from '../src/xlsx.js'
import type { SplitPanes } from '../src/contract.js'

/** A split's lines where Excel for Windows saves them: on its screen's nearest pixel (0.75 points at 96 DPI). */
function onScreen(s: SplitPanes, dpi: number): SplitPanes {
  const pixel = 72 / dpi
  const out = { ...s }
  for (const axis of ['x', 'y'] as const) if (out[axis] !== undefined) out[axis] = Math.round(out[axis] / pixel) * pixel
  return out
}

describe.each(['sheets', 'book', 'notes', 'rotation', 'split'])('the set %s', (set) => {
  const stored = JSON.parse(readFileSync(fieldsFile(set), 'utf8')) as { meta: { dpi: number }; answers: Record<string, FieldsAnswer> }
  it('Excel reads what the workbook holds, and saves it so', async () => {
    const cases = ((await import(`../bench/coverage/fields/${set}.ts`)) as FieldsSet).buildFieldsCases()
    expect(Object.keys(stored.answers).sort()).toEqual(cases.map((c) => c.id).sort())
    for (const c of cases) {
      const want = bookFields(caseBook(c))
      expect(stored.answers[c.id].sheets, c.id).toEqual(want)
      const saved = await readXlsx(new Uint8Array(readFileSync(join(savedDir(set), `${c.id}.xlsx`))))
      expect(bookFields(saved.book), c.id).toEqual(want)
      // the calculation option, read since the set book
      const calc = stored.answers[c.id].calculation
      if (calc !== undefined) expect([calc, bookCalculation(saved.book)], c.id).toEqual([bookCalculation(caseBook(c)), calc])
      // the notes, read since the set notes
      const notes = stored.answers[c.id].notes
      if (notes !== undefined) expect([notes, bookNotes(saved.book)], c.id).toEqual([bookNotes(caseBook(c)), notes])
      // the turned cells, read since the set rotation
      const turned = stored.answers[c.id].turned
      if (turned !== undefined) expect([turned, bookTurned(saved.book)], c.id).toEqual([bookTurned(caseBook(c)), turned])
      // the windows, read since the set split: as the workbook holds them, and the split saved as it was written
      const panes = stored.answers[c.id].panes
      if (panes !== undefined) {
        expect([panes.map((p) => seenPanes(p.panes)), bookPanes(saved.book)], c.id).toEqual([bookPanes(caseBook(c)), bookPanes(caseBook(c))])
        expect(splitOf(saved.book), c.id).toEqual(splitOf(caseBook(c)).map((s) => s && onScreen(s, stored.meta.dpi)))
      }
    }
  })
})
