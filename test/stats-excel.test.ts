/**
 * Workbook Statistics against Excel 16 for Mac (`test/excel/stats/`, `bench/coverage/stats.ts`): the window's rows for
 * each case's sheet and workbook are what `Session.getWorkbookStatistics` gives over the book the reader reads from the
 * case's file. Known differences are pinned in `KNOWN` with their reasons.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { engineStats, statsDifferences, statsFile, type StatsAnswer, type StatsSet } from '../bench/coverage/stats/cases.js'

/** Excel for Mac shows the notes of the current sheet and never the workbook's, which the engine counts. */
const bookNotes = (n: number) => [`workbook Notes: not shown, engine ${n}`]

/** `<set>:<case id>` → the difference lines pinned between Excel and the engine, each with its reason. */
const KNOWN: Readonly<Record<string, string[]>> = { 'fit:sheets': bookNotes(2), 'fit:notes': bookNotes(4) }

const SETS = ['fit']

describe.each(SETS)('the set %s', (set) => {
  const stored = JSON.parse(readFileSync(statsFile(set), 'utf8')) as { answers: Record<string, StatsAnswer> }
  it('every case is answered, and the window shows what the engine counts', async () => {
    const cases = ((await import(`../bench/coverage/stats/${set}.ts`)) as StatsSet).buildStatsCases()
    expect(Object.keys(stored.answers).sort()).toEqual(cases.map((c) => c.id).sort())
    for (const c of cases) expect(statsDifferences(stored.answers[c.id], await engineStats(c)), c.id).toEqual(KNOWN[`${set}:${c.id}`] ?? [])
  })
})
