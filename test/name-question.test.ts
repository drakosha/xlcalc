/**
 * A file whose names hold `?` just before what the reader of a file's formulas looks at (the table set
 * `name-question-held`, Excel for Mac and for Windows): `a?_xlfn.b` keeps its `_xlfn.` and `T?[x]` is the table `T?`,
 * so the book read back holds the formulas written, with no issue, and computes the values Excel computed.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { takenBook } from '../bench/coverage/stored.js'
import type { StoredTables } from '../bench/coverage/tables.js'
import { buildTableCases } from '../bench/coverage/tables/name-question-held.js'
import { compute } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { fromFile } from '../src/xlsx/formulas/from-file.js'
import { shownValue } from './spill-shown.js'

const answers = (path: string) => (JSON.parse(readFileSync(new URL(path, import.meta.url), 'utf8')) as StoredTables).answers
const EXCEL = { Mac: answers('./excel/tables/name-question-held.json'), Windows: answers('./excel/win/tables/name-question-held.json') }

it('reads a word with ? before a prefix or a bracket as one word', () => {
  expect(fromFile('a?_xlfn.b+1')).toEqual({ formula: 'a?_xlfn.b+1', found: [] })
  expect(fromFile('_xlfn.LET(_xlpm.x?,1,x?_xlpm.y)')).toEqual({ formula: 'LET(x?,1,x?_xlpm.y)', found: [] })
  expect(fromFile('SUM(T?[x])+ROWS(T?[[#Data],[x]])')).toEqual({ formula: 'SUM(T?[x])+ROWS(T?[[#Data],[x]])', found: [] })
})

describe.each(['closures', 'eval'] as const)('a file with ? in names and a table name, read back (%s)', (execution) => {
  it.each(buildTableCases().map((c) => [c.id, c] as const))('%s', async (id, c) => {
    const written = takenBook(c.book)
    const read = await readXlsx((await writeXlsx({ book: written })).data)
    expect(read.issues).toEqual([])
    expect(read.book.sheets[0].cells).toEqual(written.sheets[0].cells)
    const result = compute(read.book, { execution })
    for (const [platform, a] of Object.entries(EXCEL))
      for (const [at, v] of Object.entries(a[id].values)) expect(shownValue(result, 'S', at.slice(2)), `${platform} ${at}`).toEqual(v)
  })
})
