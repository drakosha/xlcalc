/**
 * The engine's files of PivotTables in Excel for Windows (set `write-1`, `bench/coverage/pivot/write-1.ts`): each book
 * Excel saved in `basic`, read and written anew by `writeXlsx`, opened by Excel. Read at once, without a refresh, its
 * PivotTable shows what Excel showed for its own book (the cells, the layout as `TableRange1` and `TableRange2`, the
 * values' area), and so does it after `RefreshTable`: what the engine wrote is what Excel lays out from the cache.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { PivotAnswer } from '../bench/coverage/pivot/cases.js'

const answers = (set: string): Record<string, PivotAnswer> => JSON.parse(readFileSync(`test/excel/win/pivot/${set}.json`, 'utf8')).answers
const ran = existsSync('test/excel/win/pivot/write-1.json')

describe.runIf(ran)('write-1: the books of basic written by writeXlsx, in Excel for Windows', () => {
  const excel = answers('basic')
  const ours = ran ? answers('write-1') : {}
  it.each(Object.keys(excel))('%s: as opened and after a refresh, as Excel showed its own', (id) => {
    const want = excel[id].phases.at(-1)!.pivots
    const got = ours[id]
    expect(got, id).toBeDefined()
    expect(got.phases.map((p) => p.errors ?? [])).toEqual([[], []])
    for (const p of got.phases) expect(p.pivots).toEqual(want)
  })
})
