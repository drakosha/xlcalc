/**
 * Windows split into panes against Excel for Mac (`test/excel/fields/`, `bench/coverage/fields.ts`): Excel read in each
 * file `writeXlsx` or `writeKept` wrote each sheet's window as the workbook holds it, and the reader reads the same
 * split back from the file Excel saved, its lines on whole points.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import type { MacFieldsAnswer } from '../bench/coverage/fields.js'
import { bookPanes, caseBook, fieldsFile, onScreen, savedDir, seenPanes, splitOf, type FieldsSet } from '../bench/coverage/fields/cases.js'
import { readXlsx } from '../src/xlsx.js'

describe.each(['split'])('the set %s', (set) => {
  const stored = JSON.parse(readFileSync(fieldsFile(set, 'fields'), 'utf8')) as { answers: Record<string, MacFieldsAnswer> }
  it('Excel reads each window as the workbook holds it, and saves it so', async () => {
    const cases = ((await import(`../bench/coverage/fields/${set}.ts`)) as FieldsSet).buildFieldsCases()
    expect(Object.keys(stored.answers).sort()).toEqual(cases.map((c) => c.id).sort())
    for (const c of cases) {
      const want = caseBook(c)
      const saved = await readXlsx(new Uint8Array(readFileSync(join(savedDir(set, 'fields'), `${c.id}.xlsx`))))
      expect([stored.answers[c.id].panes.map((p) => seenPanes(p.panes)), bookPanes(saved.book)], c.id).toEqual([bookPanes(want), bookPanes(want)])
      // Excel for Mac drops a line to the whole point before it: 150.5 saved as 150 (and 100.5 set through its object model)
      expect(splitOf(saved.book), c.id).toEqual(splitOf(want).map((s) => s && onScreen(s, 1, Math.floor)))
    }
  })
})
