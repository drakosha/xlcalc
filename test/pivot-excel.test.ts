/**
 * PivotTables against Excel for Windows, over the books and answers of the sets of `bench/coverage/pivot`
 * (`test/excel/win/pivot/<set>.json`, a book per phase: `<case>-1.xlsx`…, the last `<case>.xlsx`). Each phase's book is read and
 * its PivotTables laid out again by `xlcalc/pivot` from the cache Excel saved; the first phase's refreshed from its
 * source too; and a phase that refreshes after edits of the source is made from the phase before: its PivotTables and
 * caches, the source as this phase's book holds it, refreshed. The cells of each PivotTable's area are compared with
 * those Excel showed, numbers to 15 significant digits, a refusal with Excel's; the cases each set lists as known
 * differences are left out.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import * as basic from '../bench/coverage/pivot/basic.js'
import type { PivotPhase, PivotSet } from '../bench/coverage/pivot/cases.js'
import * as filter from '../bench/coverage/pivot/filter.js'
import * as filterHeld from '../bench/coverage/pivot/filter-held.js'
import * as layout from '../bench/coverage/pivot/layout.js'
import * as layoutHeld from '../bench/coverage/pivot/layout-held.js'
import { replays, rounded, shown } from '../bench/coverage/pivot/phases.js'
import * as refresh from '../bench/coverage/pivot/refresh.js'
import * as sort from '../bench/coverage/pivot/sort.js'
import * as sortHeld from '../bench/coverage/pivot/sort-held.js'

const DIR = new URL('./excel/win/pivot/', import.meta.url)

/**
 * The cases of a set whose PivotTables differ from Excel's, by set, as `<case>/<phase>/<how>`. basic: empty cells with
 * "For empty cells show", which Excel shows as 0 (not measured further). sort: items of several custom lists at once,
 * which Excel orders by no order we found (`docs/development.md`). refresh: a refresh growing over cells or another
 * PivotTable, which Excel lays out as one cell of #SPILL! where the placement refuses it. layout: a field's subtotals by
 * several functions, which the reading of the file does not take yet.
 */
const KNOWN: Record<string, string[]> = {
  'layout-held': [],
  'filter-held': [],
  'sort-held': [],
  basic: ['shown-errors/0/saved', 'shown-errors/0/source'],
  filter: [],
  sort: ['months-down/0/saved', 'months-down/0/source'],
  layout: ['subtotals-columns', 'subtotals-inner', 'subtotals-three-compact', 'subtotals-three'].flatMap((c) => [`${c}/0/saved`, `${c}/0/source`]),
  refresh: ['over-values/1/saved', 'over-values/1/refresh', 'two-overlap/1/saved', 'two-overlap/1/refresh', 'two-overlap-all/1/saved', 'two-overlap-all/1/refresh'],
}

const SETS: Record<string, PivotSet> = { basic, filter, sort, refresh, layout, 'layout-held': layoutHeld, 'filter-held': filterHeld, 'sort-held': sortHeld }

async function differences(set: string): Promise<string[]> {
  const answers = JSON.parse(readFileSync(new URL(`${set}.json`, DIR), 'utf8')) as { answers: Record<string, { phases: PivotPhase[] }> }
  const out: string[] = []
  for (const c of SETS[set].buildPivotCases()) {
    const answer = answers.answers[c.id]
    if (!answer) continue
    for (const { at, result: r, refresh } of await replays(set, c, answer.phases.length)) {
      const want = answer.phases[Number(at.split('/')[1])]
      const excelRefused = refresh && (want.errors ?? []).length > 0
      if (!r.ok) {
        if (!excelRefused) out.push(`${at}: refused, ${r.message}`)
        continue
      }
      if (excelRefused) {
        out.push(`${at}: Excel refused (${want.errors!.join('; ')})`)
        continue
      }
      for (const [name, area] of Object.entries(want.pivots)) {
        if ('error' in area) continue
        const got = rounded(shown(r.book, 'P', area.area))
        if (JSON.stringify(got) !== JSON.stringify(rounded(area.values))) out.push(`${at} ${name}: ${JSON.stringify(got)} for ${JSON.stringify(area.values)}`)
      }
    }
  }
  return out.sort()
}

describe('PivotTables against Excel for Windows', () => {
  for (const set of Object.keys(KNOWN))
    it(set, async () => {
      const found = await differences(set)
      expect(found.filter((d) => !KNOWN[set].some((k) => d.startsWith(`${k}:`) || d.startsWith(`${k} `)))).toEqual([])
    }, 120_000)
})
