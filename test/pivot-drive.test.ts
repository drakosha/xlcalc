import { readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import { buildPivotCases } from '../bench/coverage/pivot/sort.js'
import { drive } from '../bench/coverage/pivot/drive.js'
import { rounded, shown } from '../bench/coverage/pivot/phases.js'

const answers = JSON.parse(readFileSync(new URL('./excel/win/pivot/sort.json', import.meta.url), 'utf8')).answers
for (const id of ['region-down', 'value-down', 'months-no-lists'])
  it(`predicts ${id} from source rows without a saved layout`, () => {
    const c = buildPivotCases().find(c => c.id === id)!
    const phase = drive(c).at(-1)!
    expect(phase.errors).toEqual([])
    const want = answers[id].phases.at(-1).pivots.PT1
    expect(rounded(shown(phase.book, 'P', want.area))).toEqual(rounded(want.values))
  })
