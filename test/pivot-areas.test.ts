/**
 * `PivotLocation.areas` against the ranges Excel for Windows gave for the same pivot tables (the set `areas`, its
 * answers in `test/excel/win/pivot/areas.json`): each case played on the engine alone (`drive.ts`) or each book Excel
 * saved read as it is (`readXlsx`), as `bench/coverage/pivot/areas-predict.ts` predicts it. `table` is `TableRange1`,
 * `page` `PageRange`, `data` `DataBodyRange`; `rowLabels` is `RowRange` without its header row; `columnHeaders` the
 * table's rows above `DataBodyRange`, down to the end of `ColumnRange` where that goes lower (column labels alone).
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { PivotAreas } from '../src/contract.js'
import { parseRange, rectText } from '../src/structure/area.js'
import { predictedAreas } from '../bench/coverage/pivot/areas-predict.js'
import type { PivotArea, PivotAnswer, PivotCase, PivotSet } from '../bench/coverage/pivot/cases.js'

const SETS = ['areas', 'areas-held']

/**
 * The pivot tables whose areas Excel's ranges do not give: a pivot table with a report filter alone, whose
 * `TableRange2` the object model reads as null (the harness reads nothing), and one with no field, whose drop zones
 * have no header rows Excel names (`ColumnRange` and `RowRange` none) but its `DataBodyRange`, `A4:C20`.
 */
const OWN: Record<string, PivotAreas> = {
  'page-alone': { page: 'A1:B1', table: 'A3:C20', data: 'A4:C20' },
  empty: { table: 'A3:C20', data: 'A4:C20' },
}

/** The areas Excel's ranges give, in the library's terms. */
function excelAreas(a: PivotArea): PivotAreas {
  const x = parseRange(a.table)!
  const body = parseRange(a.body)!
  const rows = a.rows ? parseRange(a.rows) : null
  const columns = a.columns ? parseRange(a.columns) : null
  const end = Math.max(body.r1 - 1, columns?.r2 ?? -1)
  return {
    ...(a.page ? { page: a.page } : {}),
    table: a.table,
    ...(end >= x.r1 ? { columnHeaders: rectText({ ...x, r2: Math.min(end, x.r2) }) } : {}),
    ...(rows && rows.r2 > rows.r1 ? { rowLabels: rectText({ ...rows, r1: rows.r1 + 1 }) } : {}),
    data: a.body,
  }
}

describe.each(SETS)('the areas of the pivot tables of the set %s', (set) => {
  const answers = JSON.parse(readFileSync(new URL(`excel/win/pivot/${set}.json`, import.meta.url), 'utf8')) as { answers: Record<string, PivotAnswer> }
  it('are where Excel puts them', async () => {
    const { buildPivotCases } = (await import(`../bench/coverage/pivot/${set}.ts`)) as PivotSet
    const misses: string[] = []
    let compared = 0
    const cases = buildPivotCases() as PivotCase[]
    for (const c of cases) {
      const ours = await predictedAreas(c)
      for (const [key, got] of Object.entries(answers.answers[c.id].phases.at(-1)!.pivots)) {
        const want = OWN[c.id] ?? ('error' in got ? null : excelAreas(got))
        if (!want) {
          misses.push(`${c.id} ${key}: Excel ${(got as { error: string }).error}`)
          continue
        }
        compared++
        if (JSON.stringify(ours[key]) !== JSON.stringify(want)) misses.push(`${c.id} ${key}: ours ${JSON.stringify(ours[key])}, Excel ${JSON.stringify(want)}`)
      }
    }
    expect(misses).toEqual([])
    // every case, a saved book with two pivot tables twice
    expect(compared).toBeGreaterThanOrEqual(cases.length)
  })

  it.each(Object.keys(OWN).filter((id) => answers.answers[id]))('%s: Excel names no header rows, its DataBodyRange is the data', (id) => {
    const got = answers.answers[id].phases.at(-1)!.pivots.PT1
    if ('error' in got) expect(got.error).toContain('null-valued')
    else expect([got.rows, got.columns, got.body]).toEqual(['', '', OWN[id].data])
  })
})
