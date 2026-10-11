/**
 * References another writer may spell otherwise than Excel, read as Excel for Mac and for Windows read them: a rule's
 * cells in lower case or with `$` (the sets `cond/sqref-held` and `valid/sqref-held`, where Excel took every spelling),
 * a note's cell with a leading zero, `$`, in lower case or as a range of one cell (`parts/refs-held`: each on A1, and a
 * file with a note on a range of cells Excel does not open).
 */
import { describe, expect, it } from 'vitest'
import { rawFile } from '../bench/coverage/xlsx/sent.js'
import type { BookInput } from '../src/contract.js'
import { createSession } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'

const SHEET = 'xl/worksheets/sheet1.xml'

/** S!A1:B2 under a conditional format and a rule of data validation, whose `sqref` the file spells `to`. */
async function spelled(to: string) {
  const book: BookInput = {
    sheets: [
      {
        name: 'S',
        cells: { A1: { value: 5 }, B2: { value: 1 } },
        conditionalFormats: [{ type: 'expression', ranges: ['A1:B2'], formula: 'TRUE', style: { fill: { color: '#FF0000' } } }],
        validations: [{ range: 'A1:B2', type: 'whole', operator: 'greaterThan', formula1: '2' }],
      },
    ],
  }
  return readXlsx(await rawFile((await writeXlsx({ book })).data, [{ part: SHEET, from: 'sqref="A1:B2"', to: `sqref="${to}"` }]))
}

/** A note on A1 whose cell the file spells `ref`. */
async function noted(ref: string) {
  const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1, note: { author: 'Ann', text: 'on A1' } } } }] }
  return readXlsx(await rawFile((await writeXlsx({ book })).data, [{ part: 'xl/comments1.xml', from: '<comment ref="A1" ', to: `<comment ref="${ref}" ` }]))
}

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  it.each(['a1:b2', '$A$1:$B$2', '$a$1:B$2'])("a rule's cells spelled %s are A1:B2, for a conditional format and for data validation", async (to) => {
    const r = await spelled(to)
    expect(r.issues).toEqual([])
    const s = r.book.sheets[0]
    expect(s.conditionalFormats?.map((x) => x.ranges)).toEqual([['A1:B2']])
    expect(s.validations?.map((x) => x.range)).toEqual(['A1:B2'])
    const session = createSession(r.book, { execution })
    expect(session.getCells('S', 'B2').B2?.appearance?.style?.fill?.color).toBe('#FF0000')
    expect(session.validateEntry('S', 'B2', { value: 1 }).ok).toBe(false)
    session.close()
  })

  it.each(['A01', 'A1:A1', '$A$1', 'a1'])('a note on %s is on A1', async (ref) => {
    const r = await noted(ref)
    expect(r.issues).toEqual([])
    const session = createSession(r.book, { execution })
    expect(session.getCells('S', 'A1').A1?.note).toEqual({ author: 'Ann', text: 'on A1' })
    session.close()
  })

  it('a note on a range of cells, which Excel does not open, is left out and reported', async () => {
    const r = await noted('A1:B2')
    expect(r.issues.map((i) => [i.code, i.cells])).toEqual([['comment', ['A1:B2']]])
    const session = createSession(r.book, { execution })
    expect(session.getCells('S', 'A1').A1?.note).toBeUndefined()
    session.close()
  })
})
