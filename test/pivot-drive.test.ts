import { readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import { buildPivotCases } from '../bench/coverage/pivot/sort.js'
import { buildPivotCases as basic } from '../bench/coverage/pivot/basic.js'
import { buildPivotCases as show } from '../bench/coverage/pivot/show-items.js'
import { drive } from '../bench/coverage/pivot/drive.js'
import { rounded, shown } from '../bench/coverage/pivot/phases.js'

for (const [set, cases, ids] of [
  ['sort', buildPivotCases(), ['region-down', 'value-down', 'months-no-lists']],
  ['basic', basic(), ['page-one']],
  ['show-items', show(), ['hidden-run', 'page-total', 'page-run']],
] as const) {
  const answers = JSON.parse(readFileSync(new URL(`./excel/win/pivot/${set}.json`, import.meta.url), 'utf8')).answers
  for (const id of ids)
    it(`predicts ${set}/${id} from source rows without a saved layout`, () => {
      const c = cases.find(c => c.id === id)!
      const phase = drive(c).at(-1)!
      expect(phase.errors).toEqual([])
      const want = answers[id].phases.at(-1).pivots.PT1
      expect(rounded(shown(phase.book, 'P', want.area))).toEqual(rounded(want.values))
    })
}
