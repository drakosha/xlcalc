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
import * as calc from '../bench/coverage/pivot/calc.js'
import * as calcHeld from '../bench/coverage/pivot/calc-held.js'
import type { PivotPhase, PivotSet } from '../bench/coverage/pivot/cases.js'
import type { CellValue } from '../src/contract.js'
import * as filter from '../bench/coverage/pivot/filter.js'
import * as corpusShapes from '../bench/coverage/pivot/corpus-shapes.js'
import * as corpusShapes2 from '../bench/coverage/pivot/corpus-shapes-2.js'
import * as corpusShapes3 from '../bench/coverage/pivot/corpus-shapes-3.js'
import * as filterHeld from '../bench/coverage/pivot/filter-held.js'
import * as group from '../bench/coverage/pivot/group.js'
import * as groupHeld from '../bench/coverage/pivot/group-held.js'
import * as layout from '../bench/coverage/pivot/layout.js'
import * as layoutHeld from '../bench/coverage/pivot/layout-held.js'
import { replays, rounded, shown } from '../bench/coverage/pivot/phases.js'
// the sets of Excel ru-RU lay out in ru-RU
import './locales.js'
import * as refresh from '../bench/coverage/pivot/refresh.js'
import * as show from '../bench/coverage/pivot/show.js'
import * as slicer from '../bench/coverage/pivot/slicer.js'
import * as slicerHeld from '../bench/coverage/pivot/slicer-held.js'
import * as showHeld from '../bench/coverage/pivot/show-held.js'
import * as showItems from '../bench/coverage/pivot/show-items.js'
import * as sort from '../bench/coverage/pivot/sort.js'
import * as sortHeld from '../bench/coverage/pivot/sort-held.js'

const DIR = new URL('./excel/win/pivot/', import.meta.url)

/**
 * The cases of a set whose PivotTables differ from Excel's, by set, as `<case>/<phase>/<how>`. sort: items of several
 * custom lists at once, which Excel orders by no order we found (`docs/development.md`).
 */
const KNOWN: Record<string, string[]> = {
  'layout-held': [],
  'filter-held': [],
  'sort-held': [],
  basic: [],
  filter: [],
  group: [],
  calc: [],
  sort: ['months-down/0/saved', 'months-down/0/source'],
  layout: [],
  refresh: [],
  show: [],
  'show-held': [],
  'group-held': [],
  'calc-held': [],
  'show-items': [],
  slicer: [],
  'slicer-held': [],
  'ru-RU/basic': [],
  'corpus-shapes': [],
  'corpus-shapes-2': [],
  'corpus-shapes-3': [],
  'ru-RU/layout': [],
}

const SETS: Record<string, PivotSet> = { basic, filter, sort, refresh, layout, group, calc, 'layout-held': layoutHeld, 'filter-held': filterHeld, 'sort-held': sortHeld, show, 'show-items': showItems, 'show-held': showHeld, 'group-held': groupHeld, 'calc-held': calcHeld, slicer, 'slicer-held': slicerHeld, 'corpus-shapes': corpusShapes, 'corpus-shapes-2': corpusShapes2, 'corpus-shapes-3': corpusShapes3, 'ru-RU/basic': basic, 'ru-RU/layout': layout }

/**
 * The relative difference a set's numbers may have beyond their 15 significant digits: Show Values As divides by sums,
 * and a calculated field's formula takes them, whose last bit Excel's order of adding sets (`docs/limitations.md`),
 * which a share or a ratio carries into its 15th digit.
 */
const TOLERANCE: Record<string, number> = { show: 1e-14, 'show-items': 1e-14, 'show-held': 1e-14, calc: 1e-14, 'calc-held': 1e-14 }

/** The cells alike: the same values, numbers to 15 significant digits or within `tolerance` of each other. */
const alike = (got: CellValue[][], want: CellValue[][], tolerance: number) =>
  got.length === want.length &&
  got.every((r, i) => r.length === want[i].length && r.every((v, j) => {
    const w = want[i][j]
    return JSON.stringify(v) === JSON.stringify(w) || (typeof v === 'number' && typeof w === 'number' && Math.abs(v - w) <= tolerance * Math.abs(w))
  }))

async function differences(set: string): Promise<string[]> {
  const answers = JSON.parse(readFileSync(new URL(`${set}.json`, DIR), 'utf8')) as { answers: Record<string, { phases: PivotPhase[] }> }
  const tolerance = TOLERANCE[set] ?? 0
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
        if (!alike(got, rounded(area.values), tolerance)) out.push(`${at} ${name}: ${JSON.stringify(got)} for ${JSON.stringify(area.values)}`)
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
