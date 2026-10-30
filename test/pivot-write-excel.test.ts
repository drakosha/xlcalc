/**
 * The engine's files of PivotTables in Excel for Windows: each book Excel saved for a pivot set (`basic` for the set
 * `write-1`), read and written anew by `writeXlsx` (`bench/coverage/pivot/ours.ts`), opened by Excel. Read at once,
 * without a refresh, each PivotTable shows what Excel showed for its own book (the cells, `TableRange1`, `TableRange2`,
 * the values' area), and so does it after `RefreshTable`, which answers true: what the engine wrote is what Excel lays
 * out from the cache.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { PivotAnswer } from '../bench/coverage/pivot/cases.js'

const answers = (set: string): Record<string, PivotAnswer> => JSON.parse(readFileSync(`test/excel/win/pivot/${set}.json`, 'utf8')).answers

/** The sets of the engine's files and the sets of Excel's books they were written from. */
const SETS: [ours: string, excel: string][] = [['write-1', 'basic']]

describe.each(SETS.filter(([ours]) => existsSync(`test/excel/win/pivot/${ours}.json`)))('%s: the books of %s written by writeXlsx, in Excel for Windows', (set, from) => {
  const excel = answers(from)
  const ours = answers(set)
  it.each(Object.keys(excel))('%s: as opened and after a refresh, as Excel showed its own', (id) => {
    // Excel's own book, each PivotTable by its name
    const want = excel[id].phases.at(-1)!.pivots
    const got = ours[id]
    expect(got, id).toBeDefined()
    expect(got.phases.map((p) => p.errors ?? [])).toEqual([[], []])
    expect(got.phases[1].results).toEqual({ 'op 1': true })
    for (const p of got.phases) expect(Object.fromEntries(Object.entries(p.pivots).map(([k, v]) => [k.slice(k.indexOf('!') + 1), v]))).toEqual(want)
  })
})
