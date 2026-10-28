/**
 * Fields of a workbook the model holds for a host, against Excel for Windows (`test/excel/win/fields/`,
 * `bench/coverage/fields/`): Excel read from each file `writeXlsx` or `writeKept` wrote what the workbook holds, and the
 * reader reads it back from the file Excel saved.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { bookCalculation, bookFields, bookNotes, caseBook, fieldsFile, savedDir, type FieldsAnswer, type FieldsSet } from '../bench/coverage/fields/cases.js'
import { readXlsx } from '../src/xlsx.js'

describe.each(['sheets', 'book', 'notes'])('the set %s', (set) => {
  const stored = JSON.parse(readFileSync(fieldsFile(set), 'utf8')) as { answers: Record<string, FieldsAnswer> }
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
    }
  })
})
